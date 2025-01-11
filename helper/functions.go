package helper

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
)

// Vérifie si un dossier existe, sinon le crée
func EnsureDirExists(dir string) error {
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		return os.MkdirAll(dir, 0755)
	}
	return nil
}

func SanitizeFilename(filename string) string {
	invalidChars := regexp.MustCompile(`[<>:"/\\|?*]`)
	return invalidChars.ReplaceAllString(filename, "_")
}

// Fonction pour créer un fichier zip à partir des fichiers PDF
func CreateZipFile(zipFileName string, files []string) error {
	zipFile, err := os.Create(zipFileName)
	if err != nil {
		return fmt.Errorf("impossible de créer le fichier zip : %v", err)
	}
	defer zipFile.Close()

	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	for _, file := range files {
		err := AddFileToZip(zipWriter, file)
		if err != nil {
			return err
		}
	}

	return nil
}

// Fonction pour ajouter un fichier au fichier zip
func AddFileToZip(zipWriter *zip.Writer, file string) error {
	// Ouvrir le fichier à ajouter
	fileToZip, err := os.Open(file)
	if err != nil {
		return fmt.Errorf("impossible d'ouvrir le fichier pour l'ajouter au zip : %v", err)
	}
	defer fileToZip.Close()

	// Créer une entrée dans le fichier zip
	zipEntry, err := zipWriter.Create(filepath.Base(file))
	if err != nil {
		return fmt.Errorf("impossible de créer l'entrée zip : %v", err)
	}

	// Copier le contenu du fichier dans l'entrée zip
	_, err = io.Copy(zipEntry, fileToZip)
	if err != nil {
		return fmt.Errorf("impossible de copier le contenu du fichier dans le zip : %v", err)
	}

	return nil
}
