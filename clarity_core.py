import os
import tempfile
from multiprocessing import Pool, cpu_count

import cv2
import fitz  # PyMuPDF
import numpy as np
from PIL import Image
from skimage.filters import threshold_sauvola


def _process_page_worker(args):
    """
    这是在子进程中运行的函数（已修复）。
    它接收页面所需的所有数据，执行CPU密集型任务，并返回结果。
    采用“预处理->二值化->后处理”三步法。
    """
    page_num, image_bytes, temp_dir, target_dpi, window_size, k, page_size = args

    if image_bytes is None:
        return page_num, None

    try:
        # 1. 解码图像
        try:
            np_arr = np.frombuffer(image_bytes, np.uint8)
            image = cv2.imdecode(np_arr, cv2.IMREAD_GRAYSCALE)
            if image is None: raise IOError("无法用OpenCV解码图像")
        except Exception:
            from io import BytesIO
            pil_img = Image.open(BytesIO(image_bytes)).convert('L')
            image = np.array(pil_img)

        page_w_inch = page_size[0] / 72.0
        orig_h, orig_w = image.shape[:2]
        original_dpi = orig_w / page_w_inch if page_w_inch > 0 else 300

        scale_factor = target_dpi / original_dpi if original_dpi > 0 else 1.0

        if not (0.2 < scale_factor < 5.0):
            # 如果原始图片本身像素就很高，则不再放大
            if orig_w > page_size[0] * 2:
                scale_factor = 1.0

        new_width = int(orig_w * scale_factor)
        new_height = int(orig_h * scale_factor)

        resized_image = cv2.resize(image, (new_width, new_height), interpolation=cv2.INTER_AREA)

        denoised_image = cv2.medianBlur(resized_image, 3)

        threshold = threshold_sauvola(denoised_image, window_size=window_size, k=k)
        binary_np = denoised_image > threshold

        binary_image_8bit = (binary_np.astype(np.uint8)) * 255
        kernel = np.ones((2, 2), np.uint8)
        cleaned_image = cv2.morphologyEx(binary_image_8bit, cv2.MORPH_OPEN, kernel)

        pil_image = Image.fromarray(cleaned_image)

        pil_image = pil_image.convert('1')

        temp_tiff_path = os.path.join(temp_dir, f"page_{page_num}_{os.getpid()}.tiff")
        pil_image.save(temp_tiff_path, compression='group4')

        img_w, img_h = pil_image.size
        page_w, page_h = page_size
        doc = fitz.open()
        page = doc.new_page(width=page_w, height=page_h)
        ratio = min(page_w / img_w, page_h / img_h) if img_w > 0 and img_h > 0 else 0
        fit_w, fit_h = img_w * ratio, img_h * ratio
        x, y = (page_w - fit_w) / 2, (page_h - fit_h) / 2
        rect = fitz.Rect(x, y, x + fit_w, y + fit_h)
        page.insert_image(rect, filename=temp_tiff_path)

        pdf_bytes = doc.tobytes()
        doc.close()

        return page_num, pdf_bytes

    except Exception as e:
        import traceback
        print(f"子进程处理第 {page_num} 页时出错: {e}\n{traceback.format_exc()}")
        return page_num, None


def process_pdf(input_pdf_path, output_pdf_path, skip_pages, target_dpi, window_size, k, progress_callback=None):
    """
    核心处理函数 (已更新)。
    - 智能判断PDF类型，跳过文本型页面。
    - 使用安全的临时目录。
    - 传递 target_dpi 参数。
    """
    try:
        source_doc = fitz.open(input_pdf_path)
    except Exception as e:
        raise RuntimeError(f"无法打开或解析PDF文件: {input_pdf_path}. 错误: {e}")

    # 使用 tempfile 模块确保临时文件夹能被安全地自动清理
    with tempfile.TemporaryDirectory() as temp_dir:
        total_pages = len(source_doc)
        tasks = []

        if progress_callback: progress_callback.emit(5, "分析页面，准备任务...")

        # 步骤1: 准备任务列表，并进行智能判断
        for i in range(total_pages):
            page_num = i + 1
            if page_num in skip_pages:
                continue

            page = source_doc.load_page(i)

            if len(page.get_text("text", sort=True).strip()) > 50:
                print(f"第 {page_num} 页是文本型页面，跳过图像处理。")
                continue

            images = page.get_images(full=True)
            if not images:
                continue

            try:
                img_info = max(images, key=lambda img: img[2] * img[3])
                base_image = source_doc.extract_image(img_info[0])
                image_bytes = base_image["image"]
                page_size = (page.rect.width, page.rect.height)
                tasks.append((page_num, image_bytes, temp_dir, target_dpi, window_size, k, page_size))
            except Exception as e:
                print(f"提取第 {page_num} 页的图片时出错: {e}")
                continue

        processed_pages = {}
        if tasks:
            num_tasks = len(tasks)
            if progress_callback: progress_callback.emit(15, f"开始并行处理 {num_tasks} 个图像页面...")

            num_processes = max(1, cpu_count() - 1)
            with Pool(processes=num_processes) as pool:
                results_iterator = pool.imap_unordered(_process_page_worker, tasks)
                for i, result in enumerate(results_iterator):
                    page_num, pdf_bytes = result
                    processed_pages[page_num] = pdf_bytes
                    if progress_callback:
                        progress = 15 + int((i + 1) / num_tasks * 70)
                        progress_callback.emit(progress, f"已处理 {i + 1}/{num_tasks} 页...")

        if progress_callback: progress_callback.emit(90, "组装最终PDF...")
        output_doc = fitz.open()
        for i in range(total_pages):
            page_num = i + 1
            if page_num in processed_pages and processed_pages[page_num] is not None:
                processed_page_doc = fitz.open("pdf", processed_pages[page_num])
                output_doc.insert_pdf(processed_page_doc)
                processed_page_doc.close()
            else:
                output_doc.insert_pdf(source_doc, from_page=i, to_page=i)

        # 步骤4: 保存
        try:
            output_doc.save(output_pdf_path, garbage=4, deflate=True, clean=True)
            if progress_callback: progress_callback.emit(100, "保存成功！")
        finally:
            source_doc.close()
            output_doc.close()

if __name__ == '__main__':
    pass