# AI Grammar Checker - Complete Guide

[中文说明](README_Chinese.md)

## 🎯 Project Overview

This is a modern AI grammar checker that reads Word documents, uses OpenAI or Gemini APIs for grammar checks, and saves the results as Excel files. The project includes a native PyQt6 desktop application.

## 📁 Project Structure

```
AI_grammar_checker/
├── desktop_app.py          # PyQt6 desktop application
├── requirements.txt        # Dependency file
├── config.json             # Auto-generated configuration
└── README.md               # This documentation
```

## 🚀 Quick Start

1. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

2. **Run the desktop application**

   ```bash
   python desktop_app.py
   ```

## ⚙️ Configuration

### API Key Setup

Use environment variables or a system key manager to store API keys:

**OpenAI API Key:**
1. Visit [OpenAI Platform](https://platform.openai.com/api-keys)
2. Create an API key
3. Set the environment variable `OPENAI_API_KEY`

**Gemini API Key:**
1. Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Obtain the API key
3. Set the environment variable `GEMINI_API_KEY`

**Example:**
```bash
# Linux/macOS
export OPENAI_API_KEY="your OpenAI key"
export GEMINI_API_KEY="your Gemini key"

# Windows PowerShell
setx OPENAI_API_KEY "your OpenAI key"
setx GEMINI_API_KEY "your Gemini key"
```
The application reads the keys from environment variables at runtime. `config.json` no longer stores actual keys; encrypt them locally if needed.

### Configuration File

The program generates a `config.json` file with the following settings:

```json
{
  "provider": "openai",
  "model": "gpt-3.5-turbo",
  "language": "Chinese",
  "max_retries": 3,
  "retry_delay": 1,
  "session_refresh_interval": 3,
  "additional_checks": []
}
```

## 📱 Interface Overview

### Desktop Application (PyQt6)

**Main Features:**
- 🖥️ Native desktop experience
- ⚡ Background multithreading
- 🎛️ Rich widget selection
- 💾 Local configuration management

**Layout:**
1. **Left Configuration Panel**
   - API settings
   - Model settings
   - Language settings
   - Advanced settings
   - Configuration management

2. **Right Operation Panel**
   - File selection and preview
   - Output path configuration
   - Extra check management
   - Progress display
   - Start processing button

### 🌐 Interface Language Switching
- Switch languages in the **Language Settings** group
- Supports **Chinese** and **English**; selection is saved to `config.json`

## 🔧 Feature Details

### 1. Document Processing
- Supports `.docx` Word documents
- Automatically splits text into paragraphs
- Skips empty paragraphs
- Shows document preview and statistics

### 2. AI Grammar Check
- Supports OpenAI GPT series models
- Supports Google Gemini series models
- Bilingual support (Chinese & English)
- Customizable prompt templates

### 3. Session Management
- Restarts AI sessions after N paragraphs
- Avoids degradation from overly long context
- Configurable session refresh interval

## 🎛️ Advanced Settings

### Performance Optimization
- **Max retries** for failed API calls
- **Retry delay** in seconds
- **Session refresh interval**

### Cost Control
- Choose economical models (e.g., `gpt-3.5-turbo`)
- Limit the number of extra checks
- Adjust session refresh interval

### Quality Optimization
- Use advanced models (e.g., GPT-4)
- Refine extra check requirements
- Adjust prompt templates

## 🔍 Usage Tips

### 1. Model Selection Recommendations

**Cost priority:**
- OpenAI: `gpt-3.5-turbo`
- Gemini: `gemini/gemini-pro`

**Quality priority:**
- OpenAI: `gpt-4o` or `gpt-4-turbo`
- Gemini: `gemini/gemini-1.5-pro`

### 2. Extra Check Examples

**Academic writing:**
- "Check the accuracy and professionalism of terminology"
- "Check the rigor of logical arguments"
- "Check citation format compliance"

**Business documents:**
- "Check the appropriateness of business terms"
- "Check formality of expression"
- "Check completeness of information"

**Creative writing:**
- "Check the vividness and impact of language"
- "Check variety of sentence structures"
- "Check usage of rhetorical devices"

### 3. Performance Optimization Tips

**Processing large documents:**
- Increase session refresh interval (e.g., 5-10 paragraphs)
- Reduce the number of extra checks
- Choose faster models

**Improving accuracy:**
- Use advanced models
- Decrease session refresh interval (e.g., 2-3 paragraphs)
- Provide detailed extra check descriptions

## ⚠️ Notes

### Security
1. **API key safety:** Do not input API keys in public or insecure environments
2. **Data privacy:** Uploaded document content is sent to AI providers
3. **Configuration:** API keys are not stored in `config.json`; encrypt locally if needed

### Cost Control
1. **API billing:** Each API call incurs cost
2. **Usage monitoring:** Set usage limits in provider dashboards
3. **Testing advice:** Start with small documents

### Technical Limitations
1. **Document format:** Only `.docx` is supported
2. **Network requirement:** Needs a stable connection
3. **Processing time:** Large documents take longer

## 🆘 FAQ

### Q: What if API calls fail?
**A:** Check the following:
- API key correctness
- Network connection
- API quota
- Model name

### Q: Processing is slow
**A:** Try:
- Choosing faster models
- Reducing extra checks
- Increasing session refresh interval
- Checking network connectivity

### Q: How to batch process multiple documents?
**A:** The GUI supports single documents. To batch process:
- Use the command-line script
- Run the GUI repeatedly
- Contact the developer for batch scripts

### Q: Are other languages supported?
**A:** Currently supports Chinese and English. For more languages:
- Modify prompt templates
- Customize language settings
- Contact the developer for contributions

## 📞 Technical Support
If you encounter issues:
1. Check this document
2. View error messages
3. Verify configuration
4. Check network connection

## 📈 Future Plans
- [ ] Support more document formats (PDF, TXT, etc.)
- [ ] Add support for more AI providers
- [ ] Implement batch processing
- [ ] Add result statistics and analysis
- [ ] Optimize processing speed and user experience
- [ ] Add multi-language interface support

---

**Start using the AI Grammar Checker and let AI boost your writing!** 🚀
