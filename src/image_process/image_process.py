import os
import re
from PIL import Image, ImageDraw, ImageFont
from unidecode import unidecode

def convert_images_to_format(folder_path, format_no_logo, length, width):
    if not os.path.exists(folder_path):
        print(f"Thư mục {folder_path} không tồn tại.")
        return

    format_mapping = {"jpg": "JPEG", "jpeg": "JPEG", "webp": "WEBP"}
    pil_format = format_mapping.get(format_no_logo.lower(), format_no_logo)
    files = os.listdir(folder_path)

    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        if not os.path.isfile(file_path):
            continue

        try:
            with Image.open(file_path) as img:
                rgb_img = img.convert("RGB")
                resized_img = rgb_img.resize((length, width))
                new_file_path = os.path.splitext(file_path)[0] + f".{format_no_logo}"
                quality = 100
                resized_img.save(new_file_path, format=pil_format, optimize=True, quality=quality)

                while os.path.getsize(new_file_path) > 80 * 1024:  # Giới hạn 80KB
                    quality -= 5
                    if quality < 5:
                        print(f"Không thể nén file {file_name} xuống dưới giới hạn kích thước mà vẫn giữ kích thước {length}x{width}.")
                        break
                    resized_img.save(new_file_path, format=pil_format, optimize=True, quality=quality)

                if new_file_path != file_path:
                    os.remove(file_path)
                print(f"Đã chuyển đổi và nén: {file_name} -> {new_file_path}, Kích thước: {os.path.getsize(new_file_path)} bytes")
        except Exception as e:
            print(f"Lỗi khi xử lý file {file_name}: {e}")


def add_text_to_image(image_path, font_path, box_style="none"):
    try:
        import os
        from PIL import Image, ImageDraw, ImageFont

        image = Image.open(image_path)
        draw = ImageDraw.Draw(image)
        
        # Lấy tên file (không gồm phần mở rộng), rồi tách text theo dấu gạch dưới "_"
        filename = os.path.splitext(os.path.basename(image_path))[0]
        text = filename.replace("_", " ").strip().upper()

        # Tách text theo dấu phẩy để lấy các dòng
        lines = [line.strip() for line in text.split(",")]

        # Cỡ chữ tối đa và tối thiểu
        max_font_size = 32
        min_font_size = 12
        available_width = image.width - 30  # chừa lề 15px mỗi bên

        def get_max_text_width(lines_list, a_font):
            return max(draw.textlength(line, font=a_font) for line in lines_list)

        # Khoảng cách muốn đẩy text lên, khoảng ~1 cm (thường ~ 40 px)
        shift_distance = 15

        # --------------------------------------------------------------------
        # Trường hợp có đúng 2 dòng (theo yêu cầu đặc biệt):
        #   Dòng 1: màu vàng, to hơn dòng 2 là 5 đơn vị.
        #   Dòng 2: màu trắng.
        # --------------------------------------------------------------------
        margin_bottom = 0

        if len(lines) == 2:
            line1_font_size = max_font_size
            line2_font_size = line1_font_size - 5
            line1_font = ImageFont.truetype(font_path, line1_font_size)
            line2_font = ImageFont.truetype(font_path, line2_font_size)

            # Giảm dần cho đến khi vừa với khung (hoặc chạm min_font_size)
            while True:
                w1 = draw.textlength(lines[0], font=line1_font)
                w2 = draw.textlength(lines[1], font=line2_font)
                if (w1 <= available_width and w2 <= available_width) or (line2_font_size <= min_font_size):
                    break
                line1_font_size -= 1
                if line1_font_size < min_font_size + 5:
                    line1_font_size = min_font_size + 5
                line2_font_size = line1_font_size - 5
                if line2_font_size < min_font_size:
                    line2_font_size = min_font_size
                line1_font = ImageFont.truetype(font_path, line1_font_size)
                line2_font = ImageFont.truetype(font_path, line2_font_size)

            # Tính chiều cao box
            line1_height = line1_font.getbbox("hg")[3] + 5
            line2_height = line2_font.getbbox("hg")[3] + 5
            box_height = line1_height + line2_height + 20

            # Box sát đáy
            y_start = image.height - box_height - margin_bottom

            # Tạo layer để vẽ box (nếu cần)
            box_layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
            box_draw = ImageDraw.Draw(box_layer)

            if box_style == "gradient":
                fade_extra = 80  # Tùy chỉnh để tăng/giảm độ “lan tỏa” theo chiều dọc
                total_height = box_height + fade_extra

                for i in range(total_height):
                    # vì i chạy đến total_height (nhiều hơn box_height),
                    # gradient sẽ được “kéo dài” ra thêm
                    alpha = int(255 * (i / float(total_height)))
                    # tính toạ độ y, bắt đầu vẽ từ (y_start - fade_extra) lên đến y_start + box_height
                    y = y_start + i - fade_extra

                    # chỉ vẽ nếu nằm trong vùng ảnh
                    if 0 <= y < image.height:
                        color = (0, 0, 0, alpha)
                        box_draw.line([(0, y), (image.width, y)], fill=color)
        


            elif box_style == "linear_gradient":
                # Vẽ từ đáy -> lên trên để tạo linear gradient dọc
                fade_extra = 80  # nới thêm chiều cao gradient để đủ "fading"
                total_height = box_height + fade_extra
                for i in range(total_height):
                    alpha = int(255 * (1 - i / float(total_height)))
                    y = (y_start + box_height - 1) - i
                    if 0 <= y < image.height:
                        color = (0, 0, 0, alpha)
                        box_draw.line([(0, y), (image.width, y)], fill=color)

            # Bước thêm: Fade theo chiều ngang (trái–phải)
            if box_style in ["gradient"]:
                px = box_layer.load()
                for yy in range(y_start, y_start + box_height):
                    if 0 <= yy < image.height:
                        for xx in range(image.width):
                            r, g, b, a = px[xx, yy]
                            if a > 0:  # mới cần tính tiếp fade ngang
                                offset = xx / float(image.width)
                                # 70% width ở giữa => 15% mỗi bên để fade
                                if offset < 0.15:
                                    alphaH = offset / 0.15  # tăng dần 0->1
                                elif offset > 0.85:
                                    alphaH = 1 - (offset - 0.85) / 0.15  # giảm dần 1->0
                                else:
                                    alphaH = 1
                                new_a = int(a * alphaH)
                                px[xx, yy] = (r, g, b, new_a)

            if box_style != "none":
                image.paste(box_layer, (0, 0), box_layer)

            # Vẽ text hai dòng
            # Thay vì y_start + 10, ta trừ thêm shift_distance để đẩy text lên trên
            text_y = y_start + 10 - shift_distance

            # Dòng 1 (màu vàng)
            text_color_1 = (241, 196, 15, 255)
            text_width_1 = draw.textlength(lines[0], font=line1_font)
            text_x_1 = 15 + (available_width - text_width_1) // 2
            draw.text((text_x_1, text_y), lines[0], fill=text_color_1, font=line1_font)
            text_y += line1_height

            # Dòng 2 (màu trắng)
            text_color_2 = (255, 255, 255, 255)
            text_width_2 = draw.textlength(lines[1], font=line2_font)
            text_x_2 = 15 + (available_width - text_width_2) // 2
            draw.text((text_x_2, text_y), lines[1], fill=text_color_2, font=line2_font)

        else:
            # -------------------------
            # Trường hợp 1 dòng hoặc > 2 dòng
            # -------------------------
            font_size = max_font_size
            font = ImageFont.truetype(font_path, font_size)

            while get_max_text_width(lines, font) > available_width and font_size > min_font_size:
                font_size -= 1
                font = ImageFont.truetype(font_path, font_size)

            line_height = font.getbbox("hg")[3] + 5
            box_height = len(lines) * line_height + 20

            # Box sát đáy
            y_start = image.height - box_height - margin_bottom

            box_layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
            box_draw = ImageDraw.Draw(box_layer)

            if box_style == "gradient":
                fade_extra = 80  # Tùy chỉnh để tăng/giảm độ “lan tỏa” theo chiều dọc
                total_height = box_height + fade_extra

                for i in range(total_height):
                    # vì i chạy đến total_height (nhiều hơn box_height),
                    # gradient sẽ được “kéo dài” ra thêm
                    alpha = int(255 * (i / float(total_height)))
                    # tính toạ độ y, bắt đầu vẽ từ (y_start - fade_extra) lên đến y_start + box_height
                    y = y_start + i - fade_extra

                    # chỉ vẽ nếu nằm trong vùng ảnh
                    if 0 <= y < image.height:
                        color = (0, 0, 0, alpha)
                        box_draw.line([(0, y), (image.width, y)], fill=color)
        

            elif box_style == "linear_gradient":
                fade_extra = 80
                total_height = box_height + fade_extra
                for i in range(total_height):
                    alpha = int(255 * (1 - i / float(total_height)))
                    y = (y_start + box_height - 1) - i
                    if 0 <= y < image.height:
                        color = (0, 0, 0, alpha)
                        box_draw.line([(0, y), (image.width, y)], fill=color)

            # Bước thêm: Fade theo chiều ngang (trái–phải)
            if box_style in ["gradient"]:
                px = box_layer.load()
                for yy in range(y_start, y_start + box_height):
                    if 0 <= yy < image.height:
                        for xx in range(image.width):
                            r, g, b, a = px[xx, yy]
                            if a > 0:  # mới cần tính tiếp fade ngang
                                offset = xx / float(image.width)
                                # 70% width ở giữa => 15% mỗi bên để fade
                                if offset < 0.15:
                                    alphaH = offset / 0.15
                                elif offset > 0.85:
                                    alphaH = 1 - (offset - 0.85) / 0.15
                                else:
                                    alphaH = 1
                                new_a = int(a * alphaH)
                                px[xx, yy] = (r, g, b, new_a)

            if box_style != "none":
                image.paste(box_layer, (0, 0), box_layer)

            # Vẽ text
            text_color = (241, 196, 15, 255)
            text_y = y_start + 10 - shift_distance  # Đẩy text lên trên
            for line in lines:
                text_width = draw.textlength(line, font=font)
                text_x = 15 + (available_width - text_width) // 2
                draw.text((text_x, text_y), line, fill=text_color, font=font)
                text_y += line_height

        # Lưu ảnh
        image.save(image_path)

    except Exception as e:
        print(f"Lỗi khi xử lý ảnh {image_path}: {e}")
        


def process_text_in_images(folder_path, font_path, box_style="none"):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp")):
            # Tùy logic: ví dụ chỉ xử lý file có khoảng trắng trong tên
            if re.search(r'\s', filename.strip()):
                image_path = os.path.join(folder_path, filename)
                add_text_to_image(image_path, font_path, box_style)


def remove_special_characters(text):
    text = re.sub(r'[^a-zA-Z0-9\s-]', '', text)
    text = re.sub(r'[-\s]+', '-', text)
    return text

def rename_images(folder_path):
    image_extensions = ['.png', '.jpg', '.jpeg', '.webp']
    for filename in os.listdir(folder_path):
        if any(filename.lower().endswith(ext) for ext in image_extensions):
            name, ext = os.path.splitext(filename)
            new_name = unidecode(name.replace(' ', '-'))
            new_name = remove_special_characters(new_name).lower() + ext
            old_path = os.path.join(folder_path, filename)
            new_path = os.path.join(folder_path, new_name)
            os.rename(old_path, new_path)
            print(f"Đổi tên file: {filename} -> {new_name}")


from PIL import Image
import os

def compress_image(image, output_path, format_logo, max_size_kb=80):
    """Lưu ảnh với chất lượng nén sao cho không vượt quá max_size_kb KB"""
    quality = 100  # Bắt đầu với chất lượng cao nhất
    compression_level = 0  # Đối với PNG (0 = ít nén nhất, 9 = nén nhiều nhất)
    
    while True:
        if format_logo.lower() == 'jpg':
            image.save(output_path, format='JPEG', quality=quality)
        elif format_logo.lower() == 'png':
            image.save(output_path, format='PNG', compress_level=compression_level)
        else:
            image.save(output_path, format=format_logo.upper())

        # Kiểm tra kích thước file
        file_size_kb = os.path.getsize(output_path) / 1024  # Chuyển byte -> KB
        if file_size_kb <= max_size_kb:
            break  # Đạt yêu cầu, thoát vòng lặp

        if format_logo.lower() == 'jpg':
            quality -= 5  # Giảm quality nếu file quá lớn
            if quality < 10:  # Không giảm quá mức
                break
        elif format_logo.lower() == 'png':
            compression_level += 1  # Tăng mức nén
            if compression_level > 9:  # Giới hạn tối đa
                break

def add_logo_to_images(folder_path, logo_path, format_logo, length, width):
    try:
        logo = Image.open(logo_path)
        logo_width, logo_height = logo.size
        new_logo_width = 180
        new_logo_height = int((new_logo_width / logo_width) * logo_height)
        logo = logo.resize((new_logo_width, new_logo_height))

        output_folder = os.path.join(folder_path, "co_logo")
        os.makedirs(output_folder, exist_ok=True)

        for file_name in os.listdir(folder_path):
            if file_name.lower().endswith(('.jpg', '.jpeg', '.webp', '.png')):
                image_path = os.path.join(folder_path, file_name)
                with Image.open(image_path) as img:
                    if img.size != (length, width):
                        print(f"Skipping {file_name}: Size is not {length}x{width}")
                        continue

                    position = (20, 20)
                    img_with_logo = img.copy()
                    img_with_logo.paste(logo, position, logo if logo.mode == 'RGBA' else None)

                    output_file_name = f"{os.path.splitext(file_name)[0]}.{format_logo}"
                    output_path = os.path.join(output_folder, output_file_name)

                    # Gọi hàm nén để đảm bảo ảnh dưới 80 KB
                    compress_image(img_with_logo, output_path, format_logo, max_size_kb=80)

                    print(f"Saved with logo: {output_path} (<= 80KB)")

    except Exception as e:
        print(f"Error: {e}")
        


def add_frame_to_images(folder_path, frame_path, format_frame, expected_width, expected_height):
    """
    Hàm này sẽ chèn một khung (frame) lên trên ảnh, với yêu cầu:
      - Ảnh gốc phải đúng kích thước (width x height) = (expected_width, expected_height) mới xử lý
      - Khung có thể lớn hơn hoặc nhỏ hơn ảnh: 
        • Nếu khung lớn hơn, cắt bớt cho khớp (từ góc trên trái).
        • Nếu khung nhỏ hơn, giữ nguyên. Ảnh cuối cùng vẫn có kích thước (expected_width, expected_height).
      - Sau khi chèn khung, sẽ nén ảnh để đảm bảo dưới 80KB.
      - Code sửa để ghi đè lên file gốc.
    """
    try:
        frame = Image.open(frame_path)

        # Nếu frame lớn hơn ảnh thì cắt bớt
        if frame.width > expected_width or frame.height > expected_height:
            frame = frame.crop((0, 0, expected_width, expected_height))

        for file_name in os.listdir(folder_path):
            if file_name.lower().endswith(('.jpg', '.jpeg', '.webp', '.png')):
                image_path = os.path.join(folder_path, file_name)
                
                with Image.open(image_path) as img:
                    # Kiểm tra kích thước
                    if img.size != (expected_width, expected_height):
                        print(f"Skipping {file_name}: Size is not {expected_width}x{expected_height}")
                        continue

                    # Tạo ảnh mới để dán khung
                    final_image = img.copy()
                    # Nếu khung có alpha (RGBA) thì dùng mask
                    mask = frame if frame.mode == 'RGBA' else None
                    final_image.paste(frame, (0, 0), mask)

                    # Ghi đè file gốc
                    output_path = image_path  # Ghi đè lên file hiện tại
                    compress_image(final_image, output_path, format_frame, max_size_kb=80)

                    print(f"Overwrote with frame: {output_path} (<= 80KB)")
    except Exception as e:
        print(f"Error: {e}")