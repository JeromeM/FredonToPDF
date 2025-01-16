package tools

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	maxRetries       = 3
	retryDelay       = 2 * time.Second
	operationTimeout = 30 * time.Second
)

type WindowsFileProcessor struct {
	initialized bool
}

func NewWindowsFileProcessor() (*WindowsFileProcessor, error) {
	processor := &WindowsFileProcessor{}
	if err := processor.initializeCOM(); err != nil {
		return nil, fmt.Errorf("erreur d'initialisation COM : %v", err)
	}
	return processor, nil
}

func (p *WindowsFileProcessor) ProcessFile(inputFile, outputDir string) error {
	// Vérification des chemins
	if err := p.validatePaths(inputFile, outputDir); err != nil {
		return fmt.Errorf("erreur de validation des chemins : %v", err)
	}

	// Création du contexte avec timeout
	ctx, cancel := context.WithTimeout(context.Background(), operationTimeout)
	defer cancel()

	// Création de l'application Excel avec retries
	excel, err := p.createExcelApp(ctx)
	if err != nil {
		return fmt.Errorf("erreur de création de l'application Excel : %v", err)
	}
	defer safeReleaseWithRetry(excel)

	// Configuration de l'application Excel
	if err := p.configureExcel(excel); err != nil {
		return fmt.Errorf("erreur de configuration d'Excel : %v", err)
	}

	// Ouverture du classeur
	workbook, err := p.openWorkbook(excel, inputFile)
	if err != nil {
		return fmt.Errorf("erreur d'ouverture du classeur : %v", err)
	}
	defer safeReleaseWithRetry(workbook)

	// Export en PDF
	if err := p.exportToPDF(workbook, inputFile, outputDir); err != nil {
		return fmt.Errorf("erreur d'export en PDF : %v", err)
	}

	// Fermeture du classeur
	if err := p.closeWorkbook(workbook, excel); err != nil {
		return fmt.Errorf("erreur de fermeture du classeur : %v", err)
	}

	return nil
}

func (p *WindowsFileProcessor) validatePaths(inputFile, outputDir string) error {
	// Vérification du fichier d'entrée
	if _, err := os.Stat(inputFile); err != nil {
		return fmt.Errorf("le fichier d'entrée n'existe pas : %v", err)
	}

	// Vérification du dossier de sortie
	if _, err := os.Stat(outputDir); err != nil {
		if os.IsNotExist(err) {
			if err := os.MkdirAll(outputDir, 0755); err != nil {
				return fmt.Errorf("impossible de créer le dossier de sortie : %v", err)
			}
		} else {
			return fmt.Errorf("erreur lors de la vérification du dossier de sortie : %v", err)
		}
	}

	return nil
}

func (p *WindowsFileProcessor) initializeCOM() error {
	if p.initialized {
		return nil
	}

	var lastErr error
	for i := 0; i < maxRetries; i++ {
		if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
			// Si COM est déjà initialisé, on considère que c'est un succès
			if err.Error() == "CoInitialize has not been called" {
				p.initialized = true
				return nil
			}
			lastErr = err
			time.Sleep(retryDelay)
			continue
		}
		p.initialized = true
		return nil
	}

	return fmt.Errorf("échec de l'initialisation COM après %d tentatives : %v", maxRetries, lastErr)
}

func (p *WindowsFileProcessor) createExcelApp(ctx context.Context) (*ole.IDispatch, error) {
	var excel *ole.IDispatch
	var lastErr error

	for i := 0; i < maxRetries; i++ {
		select {
		case <-ctx.Done():
			return nil, fmt.Errorf("timeout lors de la création de l'application Excel")
		default:
			unknown, err := oleutil.CreateObject("Excel.Application")
			if err != nil {
				lastErr = err
				time.Sleep(retryDelay)
				continue
			}

			excel, err = unknown.QueryInterface(ole.IID_IDispatch)
			if err != nil {
				unknown.Release()
				lastErr = err
				time.Sleep(retryDelay)
				continue
			}

			return excel, nil
		}
	}

	return nil, fmt.Errorf("échec de la création de l'application Excel après %d tentatives : %v", maxRetries, lastErr)
}

func (p *WindowsFileProcessor) configureExcel(excel *ole.IDispatch) error {
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("Récupération d'une panique lors de la configuration d'Excel : %v\n", r)
		}
	}()

	// Configuration silencieuse d'Excel
	oleutil.MustPutProperty(excel, "Visible", false)
	oleutil.MustPutProperty(excel, "DisplayAlerts", false)

	return nil
}

func (p *WindowsFileProcessor) openWorkbook(excel *ole.IDispatch, inputFile string) (*ole.IDispatch, error) {
	// Convert to absolute path if it's not already
	absPath, err := filepath.Abs(inputFile)
	if err != nil {
		return nil, fmt.Errorf("impossible de convertir en chemin absolu : %v", err)
	}

	// Convert to UNC path if it's a network drive
	if strings.HasPrefix(absPath, `J:\`) {
		absPath = strings.Replace(absPath, `J:\`, `\\localhost\J$\`, 1)
	}

	// Replace backslashes with forward slashes
	absPath = strings.ReplaceAll(absPath, `\`, `/`)

	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	defer safeReleaseWithRetry(workbooks)

	// Try to open with minimal parameters first
	workbook, err := oleutil.CallMethod(workbooks, "Open", absPath)
	if err != nil {
		// If that fails, try with the full path escaped
		absPath = strings.ReplaceAll(absPath, `/`, `\`)
		workbook, err = oleutil.CallMethod(workbooks, "Open", absPath)
		if err != nil {
			return nil, fmt.Errorf("impossible d'ouvrir le classeur : %v", err)
		}
	}

	return workbook.ToIDispatch(), nil
}

func (p *WindowsFileProcessor) exportToPDF(workbook *ole.IDispatch, inputFile, outputDir string) error {
	var lastErr error
	for i := 0; i < maxRetries; i++ {
		pdfPath := filepath.Join(outputDir, strings.TrimSuffix(filepath.Base(inputFile), filepath.Ext(inputFile))+".pdf")
		if _, err := oleutil.CallMethod(workbook, "ExportAsFixedFormat", 0, pdfPath); err != nil {
			lastErr = err
			time.Sleep(retryDelay)
			continue
		}
		return nil
	}

	return fmt.Errorf("échec de l'export PDF après %d tentatives : %v", maxRetries, lastErr)
}

func (p *WindowsFileProcessor) closeWorkbook(workbook, excel *ole.IDispatch) error {
	if _, err := oleutil.CallMethod(workbook, "Close", false); err != nil {
		return fmt.Errorf("impossible de fermer le classeur : %v", err)
	}

	if _, err := oleutil.CallMethod(excel, "Quit"); err != nil {
		return fmt.Errorf("impossible de quitter Excel : %v", err)
	}

	return nil
}

func safeReleaseWithRetry(dispatch *ole.IDispatch) {
	if dispatch == nil {
		return
	}

	for i := 0; i < maxRetries; i++ {
		refCount := dispatch.Release()
		if refCount >= 0 { // A non-negative return value indicates success
			break
		}
		time.Sleep(100 * time.Millisecond)
	}
}
