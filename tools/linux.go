//go:build linux
// +build linux

package tools

import (
	"fmt"
	"os/exec"
	"path/filepath"
)

func ProcessFilesLinux(inputDir, outputDir string) {
	files, err := filepath.Glob(filepath.Join(inputDir, "*.xls"))
	if err != nil {
		fmt.Printf("Erreur lors de la lecture des fichiers : %v\n", err)
		return
	}

	for _, file := range files {
		fmt.Printf("Traitement du fichier : %s\n", file)

		// Construire la commande LibreOffice
		outputFile := filepath.Join(outputDir, filepath.Base(file)+".pdf")
		cmd := exec.Command("soffice", "--headless", "--convert-to", "pdf", "--outdir", outputDir, file)

		// Exécuter la commande
		err := cmd.Run()
		if err != nil {
			fmt.Printf("Erreur lors de la conversion avec LibreOffice : %v\n", err)
		} else {
			fmt.Printf("PDF généré : %s\n", outputFile)
		}
	}

	fmt.Println("Conversion terminée pour Linux.")
}
