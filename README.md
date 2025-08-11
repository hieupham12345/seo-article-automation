# AutoSEO Writer
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

## 📑 Table of Contents
1. [Introduction](#introduction)
2. [Features](#features)
3. [Requirements](#requirements)
4. [Installation](#installation)
5. [Notes](#notes)
6. [Usage & Demo](#usage--demo)
7. [License](#license)
8. [Contact](#contact)

---

## 1. Introduction
**AutoSEO Writer** is a fully automated SEO content generation tool that leverages APIs from **ChatGPT**, **Gemini**, and **Claude** (with the option to add other providers).  
It produces complete, well-structured, and fully optimized SEO articles for various niches and languages.  

The tool is designed to:  
- **Generate outlines** automatically  
- **Write based on the generated outline**  
- **Bulk write** multiple articles at once  
- Deliver **SEO scores of 90–100%** for most content  
- Provide **100% SEO optimization** for certain Latin-based languages  
- **Fix spin/duplicated content** using Google Search checks  
- **Auto-optimize images for SEO** (alt text, compression, naming)  

> This is a **personal project** developed during my student years.  
> While I organized the code structure as best as I could, the lack of consistent comments may make it a bit hard to follow in some areas.  
> This repository contains the **local version** without API keys.  
> If you need the **business version** with server-side & client-side separation, please contact me.

---

## 2. Features
- **Multi-Provider AI Integration**: ChatGPT, Gemini, Claude (customizable to add more)  
- **Outline Generation**: Create detailed outlines automatically  
- **Full Article Writing**: Write content based on outline with SEO structure  
- **Bulk Writing Mode**: Generate multiple articles in one run  
- **High SEO Scores**: 90–100% on popular SEO tools  
- **Perfect SEO Optimization**: 100% for selected Latin-based languages  
- **Duplicate Content Check**: Google Search integration to fix spun content  
- **Auto SEO Image Optimization**: Add alt text, compress images, and rename automatically  
- **Markdown & Word Output**: Supports exporting to `.docx`  
- **CLI & Python Module**: Run from command line or integrate into your project  

---

## 3. Requirements
- **Python 3.10**  
- Required libraries (install via `requirements.txt`)  
- API keys for:  
  - OpenAI (ChatGPT)  
  - Google Gemini  .v.v

---

## 4. Installation
Clone the repository and install dependencies using pip:

```bash
git clone https://github.com/hieupham12345/seo-article-automation.git
cd seo-article-automation
# Create and activate a virtual environment (recommended)
python -m venv venv
source venv/bin/activate       # On Linux/macOS
.\venv\Scripts\activate        # On Windows

pip install -r requirements.txt
py -3.10 main_app.py
```

**Exporting the app with PyInstaller**
If you want to package the app into a standalone Windows executable, you can use PyInstaller with the following command (run inside your activated Python 3.10 environment):

```bash
py -3.10 -m PyInstaller --noconfirm --onedir --noconsole --strip --upx-dir "C:\upx" `
    --exclude-module "matplotlib" `
    --exclude-module "numpy" `
    --exclude-module "pandas" `
    --exclude-module "tensorflow" `
    --exclude-module "torch" `
    --name "SEO Local ONLY" main_app.py `
    --add-data ".\font;font" `
    --add-data ".\.env;."
```
Make sure you have UPX installed and the path C:\upx is correct for your system.
The --exclude-module options reduce executable size by skipping heavy unused libs.
The .env file and font folder are included as data for the executable to work properly.

---

## 5. Notes

### Adding API Keys

To securely manage your API keys, create a `.env` file in the root directory of the project (same level as `main.py`), then add your keys in the following format:

```env
# Gemini API keys
GEMINI_API_KEY1=your_gemini_api_key_1
GEMINI_API_KEY2=your_gemini_api_key_2
GEMINI_API_KEY3=your_gemini_api_key_3
GEMINI_API_KEY4=your_gemini_api_key_4
GEMINI_API_KEY5=your_gemini_api_key_5

# ChatGPT API keys
CHATGPT_API_KEY1=your_chatgpt_api_key_1
CHATGPT_API_KEY2=your_chatgpt_api_key_2
CHATGPT_API_KEY3=your_chatgpt_api_key_3
CHATGPT_API_KEY4=your_chatgpt_api_key_4
CHATGPT_API_KEY5=your_chatgpt_api_key_5

# Claude API keys
CLAUDE_API_KEY1=your_claude_api_key_1
CLAUDE_API_KEY2=your_claude_api_key_2
CLAUDE_API_KEY3=your_claude_api_key_3
CLAUDE_API_KEY4=your_claude_api_key_4

# Other APIs
ERPER_SEARCH_API_1=your_erper_search_api_key
GOOGLE_SEARCH_API=your_google_search_api_key
.v.v
```

### Adding new model here
<img width="1009" height="613" alt="image" src="https://github.com/user-attachments/assets/c4e40c7c-a2d1-4338-85a5-ccde524b12c5" />

---

## 6. Usage & Demo

---

## 7. License

This project is licensed under the **MIT License**.
See the [LICENSE](LICENSE) file for details.

---

## 8. Contact

For any questions or business inquiries, feel free to reach out:

* GitHub: [hieupham12345](https://github.com/hieupham12345)
* Email: [tpmbdhieuvanpham@gmail.com](mailto:tpmbdhieuvanpham@gmail.com)
