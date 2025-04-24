# AI-Driven Fair Assignment Evaluation Using NLP Techniques

## 📌 Overview

This project presents a scalable, AI-powered system for fair and efficient evaluation of subjective student assignments. Designed for both **typed and handwritten responses**, the system utilizes **Natural Language Processing (NLP)** and **Optical Character Recognition (OCR)** to deliver accurate, transparent, and meaningful feedback.

Developed as a capstone project by **Shreya Gupta** and **Harsh Porwal**, under the guidance of **Chandru Vignesh C**, this system leverages the **Tkinter GUI**, **Gemini OCR API**, **fuzzywuzzy**, and **NLTK WordNet** to revolutionize traditional grading.

---

## 💡 Features

- **Multi-subject support**: Evaluate answers for AI, NLP, Cybersecurity, and more.
- **Typed + Handwritten Inputs**: Use Gemini OCR to extract text from uploaded handwritten answer sheets.
- **GUI Interface**: Built with Tkinter for easy login, subject selection, and answer submission.
- **NLP-Based Evaluation**:
  - Fuzzy string matching using Levenshtein distance (via fuzzywuzzy).
  - Synonym detection with WordNet to account for semantic variations.
  - Grammar and keyword presence checks.
- **Custom Scoring Matrix**: Instructors can define score weights via Excel.
- **Detailed Feedback**:
  - Grammar correctness
  - Matched keywords and their order
  - Recognized synonyms
  - Total score out of 10
- **AI-Powered Suggestions**: Model answers generated using GPT/Gemini help students understand ideal phrasing and structure.

---

## 🔧 Tech Stack

- **Frontend**: Python Tkinter
- **Backend**: Python
- **OCR**: Gemini API
- **NLP Libraries**:
  - `fuzzywuzzy`
  - `nltk` (WordNet)
- **Data Processing**: `pandas` for Excel-based answer key
- **AI Suggestions**: GPT or Gemini-based model answers (placeholder, customizable)

---

## 🚀 Installation

```bash
# Step 1: Create a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

```bash
# Step 2: Install required dependencies
pip install -r requirements.txt
```

```bash
# Step 3: Configure the Gemini OCR API Key
# Create a .env file in the root directory and add the following:
```

```env
GEMINI_API_KEY=your_api_key_here
```

```bash
# Step 4: Place the instructor's Excel-based answer key in the 'answer_keys/' directory
# Ensure the format matches expected input structure
```

```bash
# Step 5: Run the application
python main.py
```

