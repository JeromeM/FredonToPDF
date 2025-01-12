package tools

import (
	"fmt"
	"fredon_to_pdf/helper"
	"os"
	"path/filepath"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Struct pour WindowsFileProcessor
type WindowsFileProcessor struct{}

func initializeCOMWithRetry(retries int, delay time.Duration) (string, error) {
	var err error
	for i := 0; i < retries; i++ {
		err = ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
		if err == nil {
			return "", nil // Succès
		}

		fmt.Println(err.Error())

		if err.Error() == "OLE error 0x80010106" { // Erreur COM déjà initialisé
			return "", nil
		}

		// Vérifiez si l'erreur est critique ou récupérable
		if err.Error() != "OLE error 0x80010106" { // 0x80010106 signifie "COM déjà initialisé"
			return err.Error(), err // Ne pas réessayer pour une erreur critique
		}

		// Attendre avant de réessayer
		helper.GWarningLn("Échec de l'initialisation de COM. Réessai dans %v...", delay)
		ole.CoUninitialize()
		time.Sleep(delay)
	}

	// Si toutes les tentatives échouent, renvoyer la dernière erreur
	return err.Error(), fmt.Errorf("echec de l'initialisation de COM après %d tentatives : %v", retries, err)
}

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
	if errStr, err := initializeCOMWithRetry(5, 2*time.Second); err != nil {
		helper.GFatalLn("Impossible d'initialiser COM : %v - %v\n", err, errStr)
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
