# 🤖 Advanced RAG Chat System

A **Retrieval-Augmented Generation (RAG)**-based chat system built with **Streamlit**, allowing users to upload multiple document types (PDF, Word, Excel, PPT, CSV, TXT) and interact with them via **chat** or **deep research** mode.

---

## 📦 Setup & Installation

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/rag-chat-system.git
cd rag-chat-system
```

### 2. Create a virtual environment
```bash
python -m venv myenv
source myenv/bin/activate    # Linux/Mac
myenv\Scripts\activate       # Windows
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Environment variables
Create a `.env` file and add your Google API Key:
```env
GOOGLE_API_KEY=your_api_key_here
```
> Get your key from [Google AI Studio](https://ai.google.dev/).

### 5. Run the app
```bash
streamlit run app.py
```

---

## 🏗 Architecture Overview

```
                         ┌─────────────────────────┐
                         │   Streamlit Frontend    │
                         │  - Upload UI            │
                         │  - Chat Interface       │
                         │  - Research Mode        │
                         └──────────┬──────────────┘
                                    │
                         ┌──────────▼──────────────┐
                         │  Document Processor     │
                         │  - PDF, Word, PPT, etc. │
                         │  - Alternative fallback │
                         └──────────┬──────────────┘
                                    │
                         ┌──────────▼──────────────┐
                         │   Text Splitter (LC)    │
                         │   - Chunking            │
                         └──────────┬──────────────┘
                                    │
                         ┌──────────▼──────────────┐
                         │  Vector DB (Chroma)     │
                         │  - Embeddings (Gemini)  │
                         │  - Similarity Search    │
                         └──────────┬──────────────┘
                                    │
                         ┌──────────▼──────────────┐
                         │   LLM (Gemini via LC)   │
                         │  - Chat mode            │
                         │  - Deep Research mode   │
                         └─────────────────────────┘
```

---

## ✨ Features Demonstration

1. **📂 Document Upload**
   - Supports PDF, DOC/DOCX, PPT/PPTX, XLSX/CSV, TXT.
   - Bulk upload multiple files.
   - Automatic text extraction & chunking.

2. **⚡ Chat with Documents**
   - Ask natural language queries.
   - Context-aware answers based only on uploaded content.
   - Maintains chat history.

3. **🔬 Deep Research Mode**
   - Generates **structured, multi-part analysis**:
     - Executive Summary
     - Detailed Analysis
     - Supporting Evidence
     - Cross-references
     - Implications & Further Investigation

4. **📊 Processing Summary**
   - Chunks count, text length.
   - Dataframe view of processed files.

5. **💾 Export**
   - Download conversation history as CSV.

---

## 🛠 Technology Stack

- **Frontend/UI:** [Streamlit](https://streamlit.io)
- **LLM Integration:** [Google Gemini](https://ai.google.dev/) (via `langchain-google-genai`)
- **Vector Database:** [ChromaDB](https://www.trychroma.com/)
- **Text Splitting & RAG:** [LangChain](https://www.langchain.com/)
- **Document Processing:**
  - PDF → PyPDF2, PyMuPDF, pdfminer
  - DOCX → python-docx
  - DOC → docx2txt / antiword / LibreOffice fallback
  - PPTX → python-pptx
  - PPT → LibreOffice fallback
  - XLSX/CSV → pandas / openpyxl
  - TXT → utf-8/latin-1 readers
- **Utilities:** python-magic / mimetypes for file type detection

---

## 🚀 Deliverables

- Full source code (`app.py`, `alternative_processors.py`)
- `requirements.txt`
- This README

---

## 📌 Notes & Challenges

- Handling **.doc / .ppt** required fallbacks (`antiword`, `LibreOffice`) as textract is deprecated.
- Streamlit thread required **asyncio event loop fix** for gRPC/Gemini.
- Used **python-magic-bin** for Windows compatibility.
- Designed **modular processor architecture** for extensibility.
