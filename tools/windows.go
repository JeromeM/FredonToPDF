//go:build windows
// +build windows

package tools

import (
	"fmt"
	"path/filepath"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func ProcessFilesWindows(inputDir, outputDir string) {
	files, err := filepath.Glob(filepath.Join(inputDir, "*.xls"))
	if err != nil {
		fmt.Printf("Erreur lors de la lecture des fichiers : %v\n", err)
		return
	}

	// Initialiser COM
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	// Ouvrir Excel
	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		fmt.Printf("Erreur : impossible de démarrer Excel : %v\n", err)
		return
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Printf("Erreur : impossible d'interfacer Excel : %v\n", err)
		return
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "Visible", false)

	// Traiter chaque fichier
	for _, file := range files {
		fmt.Printf("Traitement du fichier : %s\n", file)

		workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
		defer workbooks.Release()

		// Ouvrir le fichier Excel
		workbook := oleutil.MustCallMethod(workbooks, "Open", file).ToIDispatch()
		defer workbook.Release()

		// Exporter en PDF
		outputFile := filepath.Join(outputDir, filepath.Base(file)+".pdf")
		oleutil.MustCallMethod(workbook, "ExportAsFixedFormat", 0, outputFile)

		// Fermer le fichier
		oleutil.MustCallMethod(workbook, "Close", false)
	}

	fmt.Println("Conversion terminée pour Windows.")
}
