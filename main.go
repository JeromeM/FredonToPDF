package main

import (
	"fmt"
	"runtime"

	"fredon_to_pdf/tools"
)

const (
	excelDir  = "./excel_files"
	outputDir = "./pdf_files"
)

func main() {
	// Vérifiez et créez les répertoires nécessaires
	if err := ensureDirExists(excelDir); err != nil {
		fmt.Printf("Erreur : impossible de créer le dossier source : %v\n", err)
		return
	}
	if err := ensureDirExists(outputDir); err != nil {
		fmt.Printf("Erreur : impossible de créer le dossier de sortie : %v\n", err)
		return
	}

	// Détection du système d'exploitation
	switch runtime.GOOS {
	case "windows":
		fmt.Println("Système détecté : Windows")
		tools.ProcessFilesWindows(excelDir, outputDir)
	case "linux":
		fmt.Println("Système détecté : Linux")
		tools.ProcessFilesLinux(excelDir, outputDir)
	default:
		fmt.Println("Système non pris en charge")
	}
}
