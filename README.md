Go Excel List Merger
A simple command-line tool written in Go to merge multiple .xlsx files from a specified folder into a single file, automatically removing duplicate rows.

Description
This script is designed to solve the problem of combining many Excel files into one master list. It reads all .xlsx files in a target directory, identifies the header, collects all unique rows, and writes the clean data to a new output file named merged_output.xlsx.

Prerequisites
Go installed on your system.

The excelize package for working with Excel files.

How to Use
Clone the repository:

git clone https://github.com/StefanSolves/listMerge.git
cd listMerge

Configure the path:
Open the main.go file and change the sourceFolder variable to the path of the directory containing your .xlsx files.

// main.go
sourceFolder := "/path/to/your/excel/files/" 

Install dependencies:
Run the following command in your terminal to download the required excelize library.

go mod tidy

Run the script:
Execute the script from your terminal.

go run main.go

The script will process the files and create a merged_output.xlsx file in the same source folder.