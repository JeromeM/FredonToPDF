package tools

import (
	"runtime"
)

// Interface FileProcessor qui définit la méthode ProcessFile
type FileProcessor interface {
	ProcessFile(inputFile, outputDir string)
}

// Fonction pour obtenir le bon FileProcessor selon l'OS
func GetFileProcessor() FileProcessor {
	switch runtime.GOOS {
	case "windows":
		return &WindowsFileProcessor{}
	case "linux":
		return &LinuxFileProcessor{}
	default:
		return nil
	}
}
