package helper

import (
	"archive/zip"
	"fmt"
	"fredon_to_pdf/types"
	"io"
	"os"
	"path/filepath"
	"regexp"

	"github.com/schollz/progressbar/v3"
)

// EnsureDirExists vérifie si un dossier existe et le crée si nécessaire
func EnsureDirExists(dir string) error {
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		if err := os.MkdirAll(dir, 0755); err != nil {
			return fmt.Errorf("impossible de créer le dossier %s : %v", dir, err)
		}
	}
	return nil
}

// SanitizeFilename nettoie un nom de fichier en remplaçant les caractères invalides par des underscores
func SanitizeFilename(filename string) string {
	invalidChars := regexp.MustCompile(`[<>:"/\\|?*]`)
	return invalidChars.ReplaceAllString(filename, "_")
}

// CreateZipFile crée un fichier ZIP contenant les fichiers spécifiés
func CreateZipFile(zipPath string, results []types.ProcessResult, bar *progressbar.ProgressBar) error {
	// Création du fichier ZIP
	zipFile, err := os.Create(zipPath)
	if err != nil {
		return fmt.Errorf("impossible de créer le fichier ZIP : %v", err)
	}
	defer zipFile.Close()

	// Création du writer ZIP
	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	// Ajout des fichiers au ZIP
	for _, result := range results {
		if result.Err != nil {
			continue
		}

		// Ouverture du fichier source
		file, err := os.Open(result.PdfPath)
		if err != nil {
			GWarningLn("Impossible d'ouvrir le fichier %s : %v", result.FileName, err)
			continue
		}

		// Création de l'entrée ZIP
		writer, err := zipWriter.Create(filepath.Base(result.PdfPath))
		if err != nil {
			file.Close()
			GWarningLn("Impossible de créer l'entrée ZIP pour %s : %v", result.FileName, err)
			continue
		}

		// Copie du contenu
		if _, err := io.Copy(writer, file); err != nil {
			file.Close()
			GWarningLn("Impossible de copier le fichier %s dans le ZIP : %v", result.FileName, err)
			continue
		}

		file.Close()
		bar.Add(1)
	}

	return nil
}
