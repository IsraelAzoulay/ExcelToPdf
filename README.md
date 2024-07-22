## Sheet2Receipt

### Description
A Python application that converts invoice data from Excel files into structured PDF documents. It processes Excel files from a specified directory and outputs well-formatted invoices with product details and total amounts.

### Features
- Automated extraction of Excel invoice data.
- Dynamic generation of PDF invoices with product details.
- Summation of total amounts for each invoice.
- Batch processing support for multiple Excel files.

### Technologies Used
- **Python**
- **FPDF**: For PDF generation.
- **Pandas & openpyxl**: For Excel file data manipulation.
- **pathlib & glob**: For file and directory management.

### How to Use
1. Ensure you have Python installed on your machine.
2. Clone the repository: git clone https://github.com/IsraelAzoulay/excel-to-pdf.git
3. Navigate to the project directory.
4. Install the required libraries using the command: pip install -r requirements.txt
5. Place your invoice Excel files in the "invoices" directory.
6. Run `main.py`: python main.py

7. Check the generated PDF invoices in the "PDFs" directory.

### Files in the Repository
- **main.py**: The main script that processes Excel files and generates PDF invoices.
- **requirements.txt**: Contains the required Python libraries for the project.
- **.gitignore**: Specifies files and directories that are to be ignored by Git.
- **invoices**: Directory containing input Excel files.
- **PDFs**: Directory containing generated PDF invoices.

### Contribution
Feel free to fork this repository, make changes, and submit pull requests. Any feedback or suggestions are welcome!

### License
This project is licensed under the MIT License.
