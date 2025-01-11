package tools

import (
	"fredon_to_pdf/helper"
	"os/exec"
)

// Struct pour LinuxFileProcessor
type LinuxFileProcessor struct{}

// Implémentation de la méthode ProcessFile pour Linux
func (l *LinuxFileProcessor) ProcessFile(inputFile, outputDir string) {

	// Construire la commande LibreOffice
	cmd := exec.Command("soffice", "--headless", "--convert-to", "pdf", "--outdir", outputDir, inputFile)

	// Exécuter la commande
	cmd.Stdout = nil
	cmd.Stderr = nil
	err := cmd.Run()
	if err != nil {
		helper.GFatalLn("Erreur lors de la conversion avec LibreOffice : %v\n", err)
	}

}
