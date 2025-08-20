package main

import (
	"fmt"
	"log"
	// "os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	// --- Configuration ---
	// 1. Set the path to the folder containing your files
	sourceFolder := "/Users/stefansolves/Downloads/bookings/"
	// 2. Set the name for the final merged and cleaned output file.
	outputFile := filepath.Join(sourceFolder, "merged_output.xlsx")

	// This map will store unique rows. We use a string representation of a row as the key
	// to automatically handle duplicates. If a row is seen more than once, it's ignored.
	uniqueRows := make(map[string][]string)
	
	// We'll store the header row separately to ensure it's at the top of the output file.
	var header []string
	isHeaderSet := false

	// --- Step 1: Find all .xlsx files in the source folder ---
	files, err := filepath.Glob(filepath.Join(sourceFolder, "*.xlsx"))
	if err != nil {
		log.Fatalf("Error finding XLSX files: %v", err)
	}

	if len(files) == 0 {
		log.Fatalf("No .xlsx files found in the directory: %s", sourceFolder)
	}

	fmt.Printf("Found %d XLSX files to merge.\n", len(files))

	// --- Step 2: Read each file and collect unique rows ---
	for _, filePath := range files {
		// We don't want to read our own output file if the script is run multiple times.
		if filepath.Base(filePath) == filepath.Base(outputFile) {
			continue
		}

		fmt.Println("  - Processing:", filepath.Base(filePath))
		
		// Open the current Excel file.
		f, err := excelize.OpenFile(filePath)
		if err != nil {
			log.Printf("Warning: Could not open file %s: %v", filePath, err)
			continue
		}

		// Get all rows from the first sheet in the workbook.
		rows, err := f.GetRows(f.GetSheetName(0))
		if err != nil {
			log.Printf("Warning: Could not get rows from file %s: %v", filePath, err)
			f.Close()
			continue
		}

		// Process each row in the file.
		for i, row := range rows {
			// Assume the first row of the first file is the header.
			if i == 0 && !isHeaderSet {
				header = row
				isHeaderSet = true
				continue // Skip adding the header to the data rows
			}
			
			// To check for duplicates, we create a simple, unique key from the row's content.
			rowKey := strings.Join(row, "|")

			// If we haven't seen this row before, add it to our map of unique rows.
			if _, exists := uniqueRows[rowKey]; !exists {
				uniqueRows[rowKey] = row
			}
		}
		f.Close()
	}

	// --- Step 3: Write the unique data to a new Excel file ---
	fmt.Println("\nWriting unique data to output file...")
	
	newFile := excelize.NewFile()
	sheetName := "Merged Data"
	newFile.SetSheetName("Sheet1", sheetName)

	// Write the header row first.
	if isHeaderSet {
		err := newFile.SetSheetRow(sheetName, "A1", &header)
		if err != nil {
			log.Fatalf("Failed to write header: %v", err)
		}
	}

	// Write all the unique data rows.
	rowIndex := 2 // Start writing data from the second row (A2)
	for _, rowData := range uniqueRows {
		cell, _ := excelize.CoordinatesToCellName(1, rowIndex)
		newFile.SetSheetRow(sheetName, cell, &rowData)
		rowIndex++
	}

	// Save the newly created Excel file.
	if err := newFile.SaveAs(outputFile); err != nil {
		log.Fatalf("Failed to save output file: %v", err)
	}

	fmt.Printf("\n✅ Success! Merged %d unique rows into %s\n", len(uniqueRows), outputFile)
}