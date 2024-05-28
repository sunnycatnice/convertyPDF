# Word to PDF Converter

![Churro Mascot](./assets/convertyPDF_churro_mascotte.png)

This project is a Word to PDF converter with an optional print functionality for the converted PDFs. The program is developed in Python and can be executed on both Windows and macOS.

## Requirements

- Python 3.x
- Windows (for context menu integration)
- macOS (for testing)
- [Pandoc](https://pandoc.org/installing.html) installed on the system

## Installation

1. **Clone the repository**:

   ```sh
   git clone <YOUR_REPOSITORY_URL>
   cd <YOUR_REPOSITORY_NAME>
   ```

2. **Create and activate the virtual environment**:

   ```sh
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install the dependencies**:

   ```sh
   pip install -r requirements.txt
   ```

## Usage

### Command Line Conversion

To convert one or more Word files to PDF without printing:

```sh
python app.py file1.docx file2.docx
```

### Optional Flags

- To print files after conversion:

  ```sh
  python app.py file1.docx file2.docx --print
  ```

- To apply additional formatting to the PDF:

  ```sh
  python app.py file1.docx file2.docx --formatted
  ```

## Features

- **Convert Word to PDF**: Convert `.docx` files to `.pdf` format.
- **Optional Printing**: Print the converted PDF files.
- **Formatting Option**: Apply additional formatting to the PDF files, including removing table borders, adding cell padding, and adjusting paragraph spacing.
