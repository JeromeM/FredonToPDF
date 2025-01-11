package tools

import (
	"fredon_to_pdf/helper"
	"os"
	"path/filepath"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Struct pour WindowsFileProcessor
type WindowsFileProcessor struct{}

// Implémentation de la méthode ProcessFile pour Windows
func (w *WindowsFileProcessor) ProcessFile(inputFile, outputDir string) {

	// Vérifier l'accès au fichier
	if _, err := os.Stat(inputFile); err != nil {
		helper.GFatalLn("Erreur : Impossible d'accéder au fichier %s : %v\n", inputFile, err)
		return
	}

	// Vérifier le dossier de sortie
	if _, err := os.Stat(outputDir); os.IsNotExist(err) {
		helper.GFatalLn("Erreur : Le dossier de sortie %s n'existe pas.\n", outputDir)
		return
	}

	// Initialiser COM
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		helper.GFatalLn("Erreur : impossible d'initialiser COM : %v\n", err)
		return
	}
	defer ole.CoUninitialize()

	// Démarrer Excel
	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		helper.GFatalLn("Erreur : impossible de démarrer Excel : %v\n", err)
		return
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		helper.GFatalLn("Erreur : impossible d'interfacer Excel : %v\n", err)
		return
	}
	defer excel.Release()

	oleutil.PutProperty(excel, "Visible", false)

	// Charger les Workbooks
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	defer workbooks.Release()

	// Ouvrir le fichier Excel
	workbook := oleutil.MustCallMethod(workbooks, "Open", inputFile, false, true).ToIDispatch()
	if workbook == nil {
		helper.GFatalLn("Erreur : Impossible d'ouvrir le fichier %s. Vérifiez qu'il n'est pas ouvert ailleurs.\n", inputFile)
		return
	}
	defer workbook.Release()

	// Exporter en PDF
	outputFile, _ := filepath.Abs(filepath.Join(outputDir, filepath.Base(inputFile)+".pdf"))
	result, _ := oleutil.CallMethod(workbook, "ExportAsFixedFormat", 0, outputFile)
	if result.Val != 0 { // Vérifie si une erreur est signalée
		helper.GFatalLn("Erreur lors de l'exportation en PDF : %v\n", result.Val)
		return
	}
	if err != nil {
		helper.GFatalLn("Erreur lors de l'exportation en PDF pour %s : %v\n", inputFile, err)
		return
	}

	// Fermer le fichier
	oleutil.MustCallMethod(workbook, "Close", false)
}
