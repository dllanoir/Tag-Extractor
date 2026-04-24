# TAGs Extractor

A high-performance, modular Python application with a modern GUI designed to extract engineering tags from PDF documents.

## Features

- **Modern GUI**: Clean and intuitive interface for easy file selection and processing.
- **High Performance**: Utilizes multi-processing parallelism to optimize extraction speed.
- **Modular Architecture**: Decoupled extraction logic from the user interface for better maintainability.
- **Portable**: Can be built into a single-file Windows executable.
- **Customizable**: Supports regex-based tag extraction (currently configured for 2 or 3 uppercase characters).

## Installation

### Prerequisites

- Python 3.8+
- Requirements listed in `requirements.txt`

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/dllanoir/Tag-Extractor.git
   cd "TAGs Extractor"
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the application:
```bash
python main.py
```

## Building the Executable

To create a portable Windows executable:
```bash
pyinstaller PDF_Tag_Extractor.spec
```

## License

[MIT](LICENSE) - Feel free to use and modify!
