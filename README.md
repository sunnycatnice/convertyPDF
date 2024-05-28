# ConvertifyPDF

![Churro Mascot](./assets/convertyPDF_churro_mascotte.png)

ConvertifyPDF is a versatile CLI tool for converting various file formats, including Word, Excel, and PowerPoint, into PDF. Developed in Python, it offers several functionalities available via command-line arguments and can be executed on both Windows and macOS.

## Requirements

- Python 3.x
- [Pandoc](https://pandoc.org/installing.html) installed on the system
- Tested on macOS

## Installation

1. **Clone the repository**:

   ```sh
   git clone https://github.com/sunnycatnice/convertyPDF.git
   cd convertyPDF
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

- **Convert Files to PDF**: Convert various file formats, including Word, Excel, and PowerPoint, to PDF.
- **Optional Printing**: Print the converted PDF files.
- **Formatting Option**: Apply additional formatting to the PDF files, including removing table borders, adding cell padding, and adjusting paragraph spacing.

## Roadmap

### Phase 1: Basic Functionality

- [x] Convert Word files to PDF
- [x] Implement optional printing of PDF files
- [x] Add formatting options (removing borders, adding padding, adjusting spacing)

### Phase 2: Testing and Compatibility

- [ ] Test on Windows
- [ ] Test on additional macOS versions
- [ ] Ensure compatibility with different versions of Python (3.x)
- [ ] Convert multiple types of files (Excel, PowerPoint) into PDF with one single command

### Phase 3: Enhancements

- [ ] Add support for batch conversion with progress indicators
- [ ] Implement a graphical user interface (GUI) for ease of use
- [ ] Add logging for debugging and usage tracking

### Phase 4: Documentation and Community

- [ ] Create detailed user documentation
- [ ] Set up a contribution guide for open-source contributors
- [ ] Implement continuous integration (CI) for automated testing

## Contributing

Contributions are welcome! Please fork the repository and use a feature branch. Pull requests are warmly welcome. See the [CONTRIBUTING](CONTRIBUTING.md) file for details.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```

```
