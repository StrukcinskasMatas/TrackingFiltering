package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/StrukcinskasMatas/TrackingFiltering/config"
)

func main() {
	fmt.Printf("Starting to read data from %s...\n", config.DATA_FILE)
	dataToFilter, err := readDataToFilter(config.DATA_FILE)
	if err != nil {
		fmt.Println("Failure reading data!")
		log.Fatal(err)
	}
	fmt.Println("Finished reading data!")

	fmt.Printf("Starting to filter %s...\n", config.EXCEL_FILE)
	start := time.Now()
	err = filterExcelFile(config.EXCEL_FILE, config.SHEET_NAME, dataToFilter, false)
	if err != nil {
		fmt.Println("\nFailure filtering data!")
		log.Fatal(err)
	}
	fmt.Printf("Finished filtering! It took %s", time.Now().Sub(start))
}

// filterExcelFile function filters an excel file by given 'valuesToRemove' data set
// removeValues:
// true - remove values that match valuesToRemove
// false - remove values that do not match valuesToRemove
func filterExcelFile(fileName string, sheetName string, valuesToRemove []string, removeValues bool) error {
	excelFile, err := excelize.OpenFile(fileName)
	if err != nil {
		return err
	}

	valuesProcessed := 0

	for i := 2; i < 30000; i++ {
		cellValue, err := excelFile.GetCellValue(sheetName, fmt.Sprintf("%s%d", config.TRACING_NUMBER_COLUMN, i))
		if err != nil {
			return err
		}

		if removeValues {
			if contains(&valuesToRemove, cellValue) {
				err := excelFile.RemoveRow(sheetName, i)
				if err != nil {
					return err
				}
				i--
			}
		} else {
			if !contains(&valuesToRemove, cellValue) && cellValue != "" {
				err := excelFile.RemoveRow(sheetName, i)
				if err != nil {
					return err
				}
				i--
			}
		}

		valuesProcessed++
		fmt.Printf("\r-- Processed %d Rows --", valuesProcessed)
	}

	err = excelFile.SaveAs(config.OUTPUT_FILE)
	if err != nil {
		return err
	}

	fmt.Println()
	return nil
}

func readDataToFilter(fileName string) ([]string, error) {
	data := []string{}

	f, err := os.Open(fileName)
	if err != nil {
		return data, err
	}
	defer f.Close()

	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		data = append(data, scanner.Text())
	}

	return data, nil
}

func contains(stringSlice *[]string, stringToFind string) bool {
	for _, value := range *stringSlice {
		if value == stringToFind {
			return true
		}
	}
	return false
}
