﻿# Invoice-generator

## Overview
This script is designed to automate the process of generating invoices. It takes an Excel file and a Word document as inputs. The Excel file contains the data to be included in the invoices, and the Word document is a template for the invoices. The script generates a separate invoice for each row in the Excel file, and each invoice is saved as a PDF file with a unique name.

## Requirements

Python 3.6 or later
pandas
docxtpl
docx2pdf

## Inputs
- Excel File
 - The Excel file should contain one or more columns of data. Each row represents a separate invoice, and each column represents a different piece of data to be included in the invoices. The column names are used as placeholders in the Word document.

- Word Document
 - The Word document serves as a template for the invoices. It should contain placeholders for the data from the Excel file. The placeholders should follow Jinja2-like syntax, which means they should be enclosed in double curly braces. For example, if you have a column named 'customer_name' in your Excel file, you should use {{ customer_name }} as the placeholder in your Word document.

## Output
The output of the script is a series of Docx files, one for each row in the Excel file. The PDF files are named 'invoice_0.docx', 'invoice_1.docx', etc., where the number corresponds to the index of the row in the Excel file.
