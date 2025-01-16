package tools

import (
	"fmt"
	"runtime"
)

// Interface FileProcessor qui définit la méthode ProcessFile
type FileProcessor interface {
	ProcessFile(inputFile, outputDir string) error
}

// Fonction pour obtenir le bon FileProcessor selon l'OS
func GetFileProcessor() (FileProcessor, error) {
	switch runtime.GOOS {
	case "windows":
		return &WindowsFileProcessor{}, nil
	default:
		return nil, fmt.Errorf("système d'exploitation non supporté: %s", runtime.GOOS)
	}
}
