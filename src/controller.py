from src.generate_outline import generate_outline
from src.write_seo import write_seo
from src.normalize_seo import process_file_normalize
from src.fix_spin import fix_spin_file
import os
from docx import Document
import re
import json
from dotenv import load_dotenv


# Lấy thư mục chứa file thực thi (dù là .py hay .exe)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")

# Load file .env từ đúng thư mục
load_dotenv(ENV_PATH)


gpt_key1 = os.environ.get("CHATGPT_API_KEY1")
gpt_key2 = os.environ.get("CHATGPT_API_KEY2")
gpt_key3 = os.environ.get("CHATGPT_API_KEY3")
gpt_key4 = os.environ.get("CHATGPT_API_KEY4")
gpt_key5 = os.environ.get("CHATGPT_API_KEY5")
gpt_api_keys = [gpt_key1, gpt_key2, gpt_key3, gpt_key4, gpt_key5]


gemini_key1 = os.environ.get("GEMINI_API_KEY1")
gemini_key2 = os.environ.get("GEMINI_API_KEY2")
gemini_key3 = os.environ.get("GEMINI_API_KEY3")
gemini_key4 = os.environ.get("GEMINI_API_KEY4")
gemini_key5 = os.environ.get("GEMINI_API_KEY5")
gemini_api_keys = [gemini_key1, gemini_key2, gemini_key3, gemini_key4, gemini_key5]

claude_key1 = os.environ.get("CLAUDE_API_KEY1")
claude_key2 = os.environ.get("CLAUDE_API_KEY2")
claude_key3 = os.environ.get("CLAUDE_API_KEY3")
claude_key4 = os.environ.get("CLAUDE_API_KEY4")

claude_api_keys = [claude_key1, claude_key2, claude_key3, claude_key4]

from concurrent.futures import ThreadPoolExecutor
from itertools import cycle


def generate_outline_func(keywords_list, folder_file, number_of_h2, number_of_h3, outline_model_name, outline_model_type, learn_data, language, seo_category, outline_data):
    # Danh sách API keys và tạo vòng lặp
    if outline_model_type == "chatgpt":
        key_cycle = cycle(gpt_api_keys)  # Tạo vòng xoay API keys
    if outline_model_type == "gemini":
        key_cycle = cycle(gemini_api_keys)  # Tạo vòng xoay API keys

    # Sử dụng ThreadPoolExecutor để xử lý đa luồng
    with ThreadPoolExecutor(max_workers=10) as executor:  # Sử dụng 6 luồng
        futures = []

        # Gửi từng công việc vào các luồng
        for keywords in keywords_list:
            current_key = next(key_cycle)  # Lấy API key tiếp theo
            # Gửi công việc vào ThreadPoolExecutor
            
            futures.append(executor.submit(generate_outline, keywords, folder_file, number_of_h2, number_of_h3, outline_model_name, outline_model_type, current_key, learn_data, language, seo_category, outline_data))

        # Đợi và lấy kết quả từ các công việc
        for future in futures:
            try:
                future.result()
            except Exception as e:
                print(f"Error occurred: {e}")

def write_full_func(keywords_list, folder_file, write_model_name, write_model_type, language, seo_category, learn_data):

    # Tạo vòng xoay API keys
    if write_model_type == "chatgpt":
        key_cycle = cycle(gpt_api_keys)  # Tạo vòng xoay API keys
    if write_model_type == "gemini":
        key_cycle = cycle(gemini_api_keys)  # Tạo vòng xoay API keys
    if write_model_type == "claude":
        key_cycle = cycle(claude_api_keys)  # Tạo vòng xoay API keys 


    # Sử dụng ThreadPoolExecutor để xử lý đa luồng
    with ThreadPoolExecutor(max_workers=10) as executor:  # Sử dụng 4 luồng  
        futures = []

        # Gửi từng từ khóa vào luồng xử lý
        for keywords in keywords_list:
            current_key = next(key_cycle)  # Lấy API key tiếp theo từ vòng xoay
            # Gửi công việc vào ThreadPoolExecutor
            
            futures.append(executor.submit(write_seo, folder_file, keywords, write_model_name, write_model_type, current_key, language, seo_category, learn_data))

        # Đợi và xử lý kết quả từ các công việc
        for future in futures:
            try:
                future.result()
            except Exception as e:
                print(f"Error occurred while processing a keyword: {e}")

# Hàm kiểm tra nếu file đã tồn tại trong thư mục đích
def is_file_processed(file_path, output_folder):
    output_file_path = os.path.join(output_folder, os.path.basename(file_path))
    return os.path.exists(output_file_path)

def normalize_seo(folder_file, max_key, number_of_image, normalize_model_name, normalize_model_type, language, word_density, brand_key):
    # Xác định thư mục chứa các file .docx đầu vào và thư mục đầu ra
    input_folder = os.path.join(folder_file, "full_seos")
    output_folder = os.path.join(input_folder, "normalize_seos")

    # Kiểm tra và tạo thư mục output nếu chưa tồn tại
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Lấy danh sách tất cả các file .docx trong thư mục input
    docx_files = [file_name for file_name in os.listdir(input_folder) 
                  if os.path.isfile(os.path.join(input_folder, file_name)) and file_name.endswith('.docx')]

    # Tạo vòng xoay API keys
    if normalize_model_type == "chatgpt":
        key_cycle = cycle(gpt_api_keys)  # Tạo vòng xoay API keys
    if normalize_model_type == "gemini":
        key_cycle = cycle(gemini_api_keys)  # Tạo vòng xoay API keys
    # Sử dụng ThreadPoolExecutor để gọi process_file với 6 luồng
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = []

        # Gửi công việc vào các luồng
        for file_name in docx_files:
            file_path = os.path.join(input_folder, file_name)
            
            if is_file_processed(file_path, output_folder):
                print(f"File already exists in output folder: {file_path}. Skipping processing.")
                continue
            
            # Lấy API key tiếp theo từ vòng xoay
            current_key = next(key_cycle)
            
            # Gửi yêu cầu xử lý file vào luồng
            futures.append(executor.submit(process_file_normalize, file_path, output_folder, max_key, number_of_image, normalize_model_name, normalize_model_type, current_key, language, word_density, brand_key))

        # Đợi và xử lý kết quả
        for future in futures:
            try:
                future.result()  # Đợi các luồng hoàn thành công việc
            except Exception as e:
                print(f"Error occurred while processing file: {e}")
                
                
def fix_spin_folder(folder_path, model_name, model_type, temperature, is_spineditor, search_type):
    # Lấy danh sách file .docx
    docx_files = [file for file in os.listdir(folder_path) if file.endswith(".docx")]
    
    # Tạo vòng xoay API keys
    if model_type == "chatgpt":
        key_cycle = cycle(gpt_api_keys)  # Tạo vòng xoay API keys
    if model_type == "gemini":
        key_cycle = cycle(gemini_api_keys)  # Tạo vòng xoay API keys
    
    # Sử dụng ThreadPoolExecutor với 10 luồng
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = []
        
        for file_name in docx_files:
            file_path = os.path.join(folder_path, file_name)
            current_key = next(key_cycle)  # Lấy API key tiếp theo
            
            # Gửi công việc vào luồng
            futures.append(executor.submit(process_fix_spin, file_path, model_name, model_type, current_key, temperature, is_spineditor, search_type))
        
        # Đợi các luồng hoàn thành
        for future in futures:
            try:
                future.result()
            except Exception as e:
                print(f"Lỗi khi xử lý file: {e}")

def split_sentence(input_string):
    """
    Chia câu thành các phần nhỏ hơn với điều kiện:
    - Câu đầu tiên ít nhất 16 từ.
    - Các câu tiếp theo tối đa 15 từ.
    """
    words = re.findall(r'\w+|\W+', input_string)  # Bao gồm cả dấu câu
    word_only = re.findall(r'\w+', input_string)  # Đếm số từ thật sự

    if len(word_only) <= 30:
        return [input_string.strip()]

    result = []
    current_sentence = []
    word_count = 0

    for word in words:
        current_sentence.append(word)
        if re.match(r'\w+', word):  # Đếm từ (không tính dấu câu)
            word_count += 1

        if word_count >= 16 and len(result) == 0:
            result.append(''.join(current_sentence).strip())
            current_sentence = []
            word_count = 0

        elif word_count >= 15:
            result.append(''.join(current_sentence).strip())
            current_sentence = []
            word_count = 0

    if current_sentence:
        if result:
            result[-1] += ' ' + ''.join(current_sentence).strip()
        else:
            result.append(''.join(current_sentence).strip())

    return result

def split_paragraph(paragraph):
    """
    Tách đoạn văn thành các câu dựa trên dấu câu (.!?).
    """
    sentences = re.split(r'(?<=[.!?])\s+', paragraph)
    split_sentences = []

    for sentence in sentences:
        split_sentences.extend(split_sentence(sentence))

    return split_sentences

def remove_quotes(sentence):
    """
    Loại bỏ dấu nháy đơn và nháy kép khỏi câu.
    """
    return re.sub(r'[\'"]', '', sentence)

def split_word_file(doc_path):
    """
    Xử lý file Word và lưu các câu đã chuẩn hóa vào file TXT.
    """
    doc = Document(doc_path)
    all_sentences = []

    for para in doc.paragraphs:
        split_sentences = split_paragraph(para.text)
        for sentence in split_sentences:
            sentence = sentence.strip()  # Loại bỏ khoảng trắng đầu cuối
            # Loại bỏ dấu nháy đơn và nháy kép
            sentence = remove_quotes(sentence)
            # Chỉ lưu câu có từ 31 kí tự trở lên (tính cả khoảng trắng ở giữa các từ)
            if len(sentence) >= 31:
                all_sentences.append(sentence)

    folder_path = os.path.dirname(doc_path)
    file_name = os.path.splitext(os.path.basename(doc_path))[0]
    output_path = os.path.join(folder_path, file_name + '.txt')

    if os.path.exists(output_path):
        print(f"File '{output_path}' đã tồn tại. Không ghi đè.")
        return

    with open(output_path, 'w', encoding='utf-8') as f:
        for sentence in all_sentences:
            f.write(sentence + '\n')

    print(f"Đã lưu kết quả vào '{output_path}'")
    
def process_fix_spin(file_path, model_name, model_type, api_key, temperature, is_spineditor, search_type):
    """ Hàm xử lý từng file một trong đa luồng """
    try:
        split_word_file(file_path)
        print(f"Đang xử lý file: {os.path.basename(file_path)}")   
        fix_spin_file(file_path, model_name, model_type, api_key, temperature, is_spineditor, search_type)
    except Exception as e:
        print(f"Lỗi xử lý file {os.path.basename(file_path)}: {e}")
        

def count_words(data):
    """
    Hàm đếm số chữ trong input, hỗ trợ nhiều kiểu dữ liệu (string, list, dict, JSON,...)
    """
    if isinstance(data, str):
        return len(data.split())
    elif isinstance(data, (list, tuple, set)):
        return sum(count_words(item) for item in data)
    elif isinstance(data, dict):
        return sum(count_words(key) + count_words(value) for key, value in data.items())
    else:
        try:
            json_str = json.dumps(data, ensure_ascii=False)
            return count_words(json_str)
        except:
            return 0