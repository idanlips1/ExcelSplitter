# ExcelSplitter
ExcelSplitter is a Java utility class that uses the Apache POI library to split large Excel files into smaller ones. Each split file contains a specified maximum number of rows, making it easier to handle large datasets. The tool is designed for cases where Excel files have more than 1000 rows, although this limit can be adjusted by the user.

Features

Splits large Excel files into multiple smaller files based on a row limit.
Copies cell values, styles, and handles various cell types (numeric, string, formula, boolean).
Supports Excel .xlsx format using SXSSFWorkbook for streaming large workbooks.
Requirements

Java 8 or higher
Apache POI library (poi-ooxml and poi packages)
Installation

Clone the repository or download the ExcelSplitter.java file.
Add Apache POI dependencies to your project:
xml
Copy code
<!-- Maven dependencies -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.0.0</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.0.0</version>
</dependency>
Usage

Instantiate the ExcelSplitter class by passing the file path and the maximum number of rows you want in each split file.

public class Main {
    public static void main(String[] args) {
        // Example: Split the Excel file into smaller chunks of 1000 rows per file.
        ExcelSplitter splitter = new ExcelSplitter("/path/to/large_file.xlsx", 1000);
    }
}
The ExcelSplitter will automatically handle the splitting and saving of the smaller files with the format fileName_0.xlsx, fileName_1.xlsx, etc.
Code Breakdown

Constructor:
public ExcelSplitter(String fileNamePath, int maxRows);

Initializes the file path and the maximum number of rows allowed per split file.
If the total number of rows exceeds the limit, it splits the workbook into smaller chunks.

splitWorkbook(XSSFWorkbook workbook)
Splits the original workbook into multiple SXSSFWorkbook instances based on the maxRows limit.
Copies rows and cell data into the new workbooks.

setValue(SXSSFCell newCell, Cell cell)
Copies the value from the original cell to the new cell, handling different cell types (numeric, string, formula, boolean, etc.).

writeWorkBooks(List<SXSSFWorkbook> wbs)
Saves each SXSSFWorkbook as a separate Excel file, with unique filenames generated based on the original file name.

extractFileName(String filepath)
Extracts the file name from a file path string, supporting both / and \\ path separators.
Example

Given a file large_file.xlsx with 5000 rows, running:

ExcelSplitter splitter = new ExcelSplitter("/path/to/large_file.xlsx", 1000);
Will generate 5 new files:

large_file_0.xlsx
large_file_1.xlsx
large_file_2.xlsx
large_file_3.xlsx
large_file_4.xlsx
Notes

This tool assumes the Excel file has only one sheet. For multi-sheet support, additional modifications would be necessary.
The row limit is adjustable by passing a different value for maxRows in the constructor.
License

This project is licensed under the MIT License - see the LICENSE file for details.


