package main

import (
	"fmt"
	"fredon_to_pdf/config"
	"fredon_to_pdf/helper"
	"fredon_to_pdf/tools"
	"fredon_to_pdf/types"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/schollz/progressbar/v3"
)

var (
	Version   = "dev"
	BuildDate = "unknown"
)

func main() {
	if err := run(); err != nil {
		helper.GFatalLn("Erreur fatale : %v", err)
		os.Exit(1)
	}
}

func run() error {
	displayHeader()

	// Initialisation de la configuration
	cfg := config.NewConfig()
	if err := initializeDirs(cfg); err != nil {
		return err
	}

	// Récupération des fichiers
	files, err := getExcelFiles(cfg.ExcelDir)
	if err != nil {
		return err
	}

	if len(files) == 0 {
		helper.GWarningLn("Aucun fichier Excel trouvé dans %s", cfg.ExcelDir)
		return nil
	}

	// Traitement des fichiers
	results, err := processFiles(cfg, files)
	if err != nil {
		return err
	}

	// Filtrer les résultats avec succès
	successResults := filterSuccessResults(results)

	// Gestion du ZIP si nécessaire et s'il y a des fichiers traités avec succès
	if len(successResults) > 0 && strings.ToLower(cfg.CompressToZip) == "o" {
		if err := handleZipCreation(cfg, successResults); err != nil {
			return err
		}
	}

	// Afficher le résumé
	displaySummary(results)

	helper.GBlank()
	helper.GInfoLn("Appuyez sur une touche pour fermer...")
	fmt.Scanln()

	return nil
}

func displayHeader() {
	helper.GInfoLn("VINCE'S CUSTOM EXCEL TO PDF CONVERTER")
	helper.GInfoLn("---------------------------------------")
	helper.GBlank()
}

func initializeDirs(cfg *config.Config) error {
	if err := helper.EnsureDirExists(cfg.ExcelDir); err != nil {
		return fmt.Errorf("impossible de créer le dossier source : %v", err)
	}
	if err := helper.EnsureDirExists(cfg.OutputDir); err != nil {
		return fmt.Errorf("impossible de créer le dossier de sortie : %v", err)
	}
	return nil
}

func getExcelFiles(dir string) ([]string, error) {
	helper.GBlank()
	helper.GInfoLn("Récupération des fichiers Excel..")

	// Convert input directory to absolute path
	absDir, err := filepath.Abs(dir)
	if err != nil {
		return nil, fmt.Errorf("impossible de convertir le dossier en chemin absolu : %v", err)
	}

	var files []string
	extensions := []string{"*.xls", "*.xlsx"}

	for _, ext := range extensions {
		matches, err := filepath.Glob(filepath.Join(absDir, ext))
		if err != nil {
			return nil, fmt.Errorf("erreur lors de la lecture des fichiers %s : %v", ext, err)
		}
		files = append(files, matches...)
	}

	return files, nil
}

func processFiles(cfg *config.Config, files []string) ([]types.ProcessResult, error) {
	helper.GBlank()
	helper.GInfoLn("Traitement des fichiers Excel..")
	helper.GBlank()

	// Création de la barre de progression
	bar := progressbar.NewOptions(len(files),
		progressbar.OptionEnableColorCodes(true),
		progressbar.OptionShowCount(),
		progressbar.OptionSetWidth(40),
		progressbar.OptionSetDescription("[cyan]Conversion en cours...[reset]"),
		progressbar.OptionSetTheme(progressbar.Theme{
			Saucer:        "[green]=[reset]",
			SaucerHead:    "[green]>[reset]",
			SaucerPadding: " ",
			BarStart:      "[",
			BarEnd:        "]",
		}))

	// Nombre de workers basé sur le nombre de CPU
	numWorkers := runtime.NumCPU()
	if numWorkers > 4 {
		numWorkers = 4 // Limite à 4 workers maximum pour éviter la surcharge
	}

	// Channels pour la gestion des tâches
	jobs := make(chan string, len(files))
	results := make(chan types.ProcessResult, len(files))
	var wg sync.WaitGroup

	// Démarrage des workers
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()

			// Création d'un nouveau processeur pour chaque goroutine
			processor, err := tools.NewWindowsFileProcessor()
			if err != nil {
				results <- types.ProcessResult{
					FileName: "initialization",
					Err:      fmt.Errorf("erreur d'initialisation du processeur : %v", err),
				}
				return
			}

			for file := range jobs {
				result := processFile(file, cfg.OutputDir, processor)
				results <- result
				bar.Add(1)
				time.Sleep(100 * time.Millisecond) // Petit délai pour éviter la surcharge
			}
		}()
	}

	// Envoi des fichiers aux workers
	for _, file := range files {
		jobs <- file
	}
	close(jobs)

	// Attente de la fin du traitement
	go func() {
		wg.Wait()
		close(results)
	}()

	// Collecte des résultats
	var processResults []types.ProcessResult
	for result := range results {
		processResults = append(processResults, result)
	}

	helper.GBlank()
	return processResults, nil
}

func processFile(file, outputDir string, processor tools.FileProcessor) types.ProcessResult {
	result := types.ProcessResult{
		FileName: filepath.Base(file),
	}

	if err := checkFilePermissions(file); err != nil {
		result.Err = fmt.Errorf("erreur de permissions : %v", err)
		return result
	}

	if err := processor.ProcessFile(file, outputDir); err != nil {
		result.Err = err
		return result
	}

	// Construction du chemin PDF
	pdfName := strings.TrimSuffix(filepath.Base(file), filepath.Ext(file)) + ".pdf"
	result.PdfPath = filepath.Join(outputDir, pdfName)

	return result
}

func checkFilePermissions(file string) error {
	// Vérification des permissions en lecture
	f, err := os.OpenFile(file, os.O_RDONLY, 0)
	if err != nil {
		return fmt.Errorf("impossible d'ouvrir le fichier en lecture : %v", err)
	}
	f.Close()
	return nil
}

func filterSuccessResults(results []types.ProcessResult) []types.ProcessResult {
	var successResults []types.ProcessResult
	for _, result := range results {
		if result.Err == nil {
			successResults = append(successResults, result)
		}
	}
	return successResults
}

func handleZipCreation(cfg *config.Config, results []types.ProcessResult) error {
	helper.GBlank()
	helper.GInfoLn("Création du fichier ZIP..")

	zipPath := filepath.Join(cfg.OutputDir, "pdfs.zip")
	bar := progressbar.NewOptions(len(results),
		progressbar.OptionEnableColorCodes(true),
		progressbar.OptionShowCount(),
		progressbar.OptionSetWidth(40),
		progressbar.OptionSetDescription("[cyan]Compression en cours...[reset]"),
		progressbar.OptionSetTheme(progressbar.Theme{
			Saucer:        "[green]=[reset]",
			SaucerHead:    "[green]>[reset]",
			SaucerPadding: " ",
			BarStart:      "[",
			BarEnd:        "]",
		}))

	if err := helper.CreateZipFile(zipPath, results, bar); err != nil {
		return fmt.Errorf("erreur lors de la création du ZIP : %v", err)
	}

	helper.GBlank()
	helper.GInfoLn("Fichier ZIP créé avec succès : %s", zipPath)
	return nil
}

func displaySummary(results []types.ProcessResult) {
	helper.GBlank()
	helper.GInfoLn("Résumé de la conversion :")
	helper.GBlank()

	success := 0
	failed := 0
	for _, result := range results {
		if result.Err == nil {
			success++
		} else {
			failed++
			helper.GFatalLn("Échec pour %s : %v", result.FileName, result.Err)
		}
	}

	helper.GInfoLn("Fichiers traités avec succès : %d", success)
	helper.GInfoLn("Fichiers en échec : %d", failed)
}
