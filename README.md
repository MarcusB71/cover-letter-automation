# Cover Letter Automation

## Project Overview

This project started as a means to automate the tedious process of filling out employer specific information within my cover letter. The application allows the user to select a file which will be parsed and then dynamically generates fields based on any text that is placed within a delimiter (default=[]). The fields will each have input boxes where the user can enter a string that will eventually replace each field's instance within the selected file. The user can then select a file type and click export to name and generate a new file with the replacements.

### Project Applications

While the project initially only supported static, pre-defined fields, it now parses the selected file and dynmically generates the fields. This means that the application can be used for more than just cover letter automation. If you have a file that you are often making minor changes to (ie. dates, employer info), then this app is for you! Just note that the app only supports reading from .docx and writing to .pdf and .docx

### Preview

![Preview of app](https://github.com/MarcusB71/cover-letter-automation/blob/main/Preview.png)

## How to Install and Run

To Install and run the Cover Letter Automation app, follow these steps:

1. Clone the repository to your local machine
2. Make sure you have Python installed on your system. You can download it from [python.org](https://www.python.org/downloads/).
3. Install the required Python packages by running the following commands in your terminal:

```
pip install python-docx
pip install docx2pdf
```

4. Run the application by executing the following command:

```
python main.py
```

## How to Use

1. Select a .docx file that contains text inside a square bracket delimiter (ex. [text])
2. Enter replacement text into the input fields which will later replace [text]
3. Select desired export file type
4. Click "Save as" when ready to finalize the export
5. Enter file name and click save

## Why Use This App

### Pros:

- **Preservation of Original Document:** Users can create a new document without altering the original .docx file.
- **Avoids Formatting Issues:** Eliminates the risk of formatting discrepancies that occur when manually copy + pasting, ensuring consistent font size, font style, and spacing
- **Efficiency:** Streamlines the process of making repetitive changes to documents, saving time and effort
- **Supports Multiple File Formats:** Supports reading from .docx files and exporting to .pdf and .docx formats

### Cons:

- **Initial Setup:** Users need to install and run the app (I suggest utilizing auto-py-to-exe to create a standalone executable)
- **Limited File Type Support:** The app currently only supports reading from .docx and writing to .docx or .pdf

# Contribution

Contributions are welcome. Please fork the repository and submit a pull request for review.
