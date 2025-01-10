package main

import (
	"os"
)

// Vérifie si un dossier existe, sinon le crée
func ensureDirExists(dir string) error {
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		return os.MkdirAll(dir, 0755)
	}
	return nil
}
