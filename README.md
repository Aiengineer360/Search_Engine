# Desktop Search Engine

A Python-based desktop search engine application that indexes and searches through PDF, Word (DOCX), PowerPoint (PPTX), Excel (XLSX), and plain text (TXT) documents using TF-IDF ranking.

## Features

- **Multi-format Support**: Indexes content from:
  - PDF files
  - Word documents (.docx)
  - PowerPoint presentations (.pptx)
  - Excel spreadsheets (.xlsx)
  - Plain text files (.txt)
  
- **Advanced Search**:
  - TF-IDF ranking for relevant results
  - Stemming and stopword removal for better matching
  - Query term highlighting in results

- **User Interface**:
  - Add/remove documents
  - Paginated search results
  - Progress tracking for document processing

- **Persistence**:
  - Saves index between sessions
  - Automatic loading of previous index on startup

## Installation

### Prerequisites

- Python 3.6+
- Required packages:
  ```bash
  pip install PyPDF2 python-docx python-pptx openpyxl nltk
  ```

### Running the Application

1. Clone the repository:
   ```bash
   git clone https://github.com/Aiengineer360/Search_Engine.git
   cd desktop-search-engine
   ```

2. Run the application:
   ```bash
   python Desktop_SearchEngine.py
   ```

## Usage

1. **Adding Documents**:
   - Click "Add Document" and select files to index
   - Supported formats: PDF, DOCX, PPTX, XLSX, TXT

2. **Searching**:
   - Enter your query in the search box
   - Click "Search" to see relevant documents
   - Use pagination buttons to navigate results

3. **Managing Documents**:
   - Remove documents using the "Remove Document" button
   - The index automatically updates when adding/removing files

## Technical Details

- **Inverted Index**: Uses a TF-IDF weighted inverted index for efficient searching
- **Text Processing**:
  - Tokenization and stemming using NLTK
  - Stopword removal
- **Persistence**: Index is saved to `inverted_index.pkl` and documents to `documents.pkl`

## File Structure

```
desktop-search-engine/
├── search_engine.py      # Main application code
├── inverted_index.pkl    # Saved index (auto-generated)
├── documents.pkl         # Saved document data (auto-generated)
└── README.md            # This file
```
