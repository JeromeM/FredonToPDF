package main

import (
	"fmt"
	"fredon_to_pdf/config"
	"fredon_to_pdf/helper"
	"fredon_to_pdf/tools" // Assurez-vous que l'importation est correcte
	"os"
	"path/filepath"
	"strings"

	"github.com/schollz/progressbar/v3"
)

func main() {

	helper.GInfoLn("VINCE'S CUSTOM EXCEL TO PDF CONVERTER")
	helper.GInfoLn("---------------------------------------")
	helper.GBlank()

	var cfg *config.Config = config.NewConfig()

	// Vérifiez et créez les répertoires nécessaires
	if err := helper.EnsureDirExists(cfg.ExcelDir); err != nil {
		helper.GFatalLn("Erreur : impossible de créer le dossier source : %v\n", err)
		return
	}
	if err := helper.EnsureDirExists(cfg.OutputDir); err != nil {
		helper.GFatalLn("Erreur : impossible de créer le dossier de sortie : %v\n", err)
		return
	}

	// Récupérer la liste des fichiers Excel à traiter
	helper.GBlank()
	helper.GInfoLn("Récuperation des fichiers Excel..")
	files, err := filepath.Glob(filepath.Join(cfg.ExcelDir, "*.xls"))
	if err != nil {
		helper.GFatalLn("Erreur lors de la lecture des fichiers : %v\n", err)
		return
	}
	helper.GInfoLn("%d fichiers trouvés", len(files))
	helper.GBlank()

	bar := progressbar.NewOptions(len(files),
		progressbar.OptionEnableColorCodes(true),
		progressbar.OptionSetWidth(50),
	)

	// Obtenir le processeur de fichiers pour l'OS actuel
	fileProcessor := tools.GetFileProcessor()
	if fileProcessor == nil {
		helper.GFatalLn("Système non pris en charge")
		return
	}

	// Boucle sur chaque fichier et appel à ProcessFile
	helper.GInfoLn("Traitement des fichiers en cours ...")
	var fullFile string = ""
	var pdfFiles []string
	for _, file := range files {

		fileInfo, err := os.Stat(file)
		if err != nil {
			helper.GFatalLn("Erreur : Impossible d'accéder au fichier %s : %v\n", file, err)
		}

		// Test pour les permissions
		if fileInfo.Mode()&0222 == 0 {
			helper.GFatalLn("Erreur : Le fichier %s est protégé en écriture.\n", file)
		}

		coloredFileName := fmt.Sprintf("\x1b[%d;1m%s\x1b[0m", 33, filepath.Base(file))

		if fullFile, err = filepath.Abs(file); err != nil {
			helper.GFatalLn("Erreur lors de la récupération du chemin absolu du fichier : %v", err)
			return
		}

		bar.Describe(fmt.Sprintf("[ %s ]", coloredFileName))
		fileProcessor.ProcessFile(fullFile, cfg.OutputDir)

		pdfFile := filepath.Join(cfg.OutputDir, strings.TrimSuffix(filepath.Base(file), filepath.Ext(file))+".pdf")
		pdfFiles = append(pdfFiles, pdfFile)

		bar.Add(1)
	}

	if strings.ToLower(cfg.CompressToZip) == "o" {
		helper.GBlank()
		helper.GBlank()
		helper.GInfoLn("Création du fichier ZIP")
		zipFileName := filepath.Join(cfg.OutputDir, "pdf_files.zip")
		if err := helper.CreateZipFile(zipFileName, pdfFiles); err != nil {
			helper.GFatalLn("Erreur lors de la création du fichier zip : %v", err)
		} else {
			// Supprimer les fichiers PDF après la création du zip
			helper.GInfoLn("Suppression des vieux fichiers PDF...")
			for _, pdfFile := range pdfFiles {
				if err := os.Remove(pdfFile); err != nil {
					helper.GWarningLn("Erreur lors de la suppression du fichier %s : %v", pdfFile, err)
				}
			}

		}
	}

	helper.GBlank()
	helper.GInfoLn("Opération terminée.")
}
