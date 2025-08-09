import openai
import google.generativeai as genai
import anthropic



def call_chatbot(prompt, model_name, model_type, api_key, temperature = 2):
    """
    Unified function to call ChatGPT, Claude, or Gemini APIs based on model_type.

    Args:
        prompt (str): Input prompt for the chatbot.
        model_name (str): Specific model to be used.
        model_type (str): Type of the model ("chatgpt", "claude", "gemini").

    Returns:
        tuple: Response text and cost in VND.
    """
    try:
        if model_type == "chatgpt":
            return call_chatgpt(prompt, model_name, api_key)
        elif model_type == "claude":
            return call_claude(prompt, model_name, api_key)
        elif model_type == "gemini":
            return call_gemini(prompt, model_name, api_key, temperature)
        else:
            return f"Unsupported model type: {model_type}", 0
    except Exception as e:
        return f"An error occurred: {e}", 0




def call_chatgpt(prompt, model, api_key, temperature=1.0):
    
    response = openai.ChatCompletion.create( 
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        api_key=api_key
    )

    return response.choices[0].message['content']



def call_claude(prompt, model, api_key, temperature=1.0):
    # Khởi tạo client Anthropics với api_key
    client = anthropic.Anthropic(api_key=api_key)
    
    # Tạo message yêu cầu Claude trả lời với prompt đầu vào
    message = client.messages.create(
        model=model,  # Sử dụng model phù hợp, ví dụ: "claude-3-sonnet-20240229"
        max_tokens=3000,
        temperature=temperature,  # Thêm temperature ở đây
        messages=[{"role": "user", "content": prompt}]
    )
    
    # Trích xuất text từ các content blocks
    text_response = ""
    for content_block in message.content:
        if content_block.type == 'text':
            text_response += content_block.text
    
    # Trả về nội dung text thuần
    return text_response


def call_gemini(prompt, model_name, api_key, temperature=2.0):
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        response = model.generate_content(prompt, generation_config={"temperature": temperature})

        return response.text
    except Exception as e:
        return f"Error in call_gemini: {e}", 0