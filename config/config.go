package config

import (
	"encoding/json"
	"fmt"
	"fredon_to_pdf/helper"
	"os"
)

const (
	defaultExcelDir      = "./excel_files"
	defaultOutputDir     = "./pdf_files"
	defaultCompressToZip = "O"
	configFilePath       = "./config.json"
)

type Config struct {
	ExcelDir      string `json:"excel_dir"`
	OutputDir     string `json:"output_dir"`
	CompressToZip string `json:"compress_to_zip"`
}

func NewConfig() *Config {
	cfg := &Config{}
	return cfg.configure(configFilePath)
}

func (cfg *Config) configure(configFilePath string) *Config {
	if _, err := os.Stat(configFilePath); err == nil {
		cfg, err = loadConfig(configFilePath)
		if err != nil {
			helper.GFatalLn("Erreur lors du chargement de la configuration : %v", err)
		}
	}

	// Demander à l'utilisateur le dossier source (Excel files)
	if cfg.ExcelDir == "" {
		helper.GInfo("Veuillez entrer le chemin du dossier des fichiers Excel (ou appuyez sur Entrée pour utiliser '%s') : ", defaultExcelDir)
		fmt.Scanln(&cfg.ExcelDir)
		if cfg.ExcelDir == "" {
			cfg.ExcelDir = defaultExcelDir
		}
	}

	// Demander à l'utilisateur le dossier de sortie (PDF files)
	if cfg.OutputDir == "" {
		helper.GInfo("Veuillez entrer le chemin du dossier de sortie des fichiers PDF (ou appuyez sur Entrée pour utiliser '%s') : ", defaultOutputDir)
		fmt.Scanln(&cfg.OutputDir)
		if cfg.OutputDir == "" {
			cfg.OutputDir = defaultOutputDir
		}
	}

	// Demander si il veut créer un ZIP
	if cfg.CompressToZip == "" {
		helper.GInfo("Voulez-vous créer un fichier ZIP avec les fichiers PDF générés ? (O/n) : ")
		fmt.Scanln(&cfg.CompressToZip)
		if cfg.CompressToZip == "" {
			cfg.CompressToZip = defaultCompressToZip
		}
	}

	if err := saveConfig(cfg, configFilePath); err != nil {
		helper.GFatalLn("Erreur lors de la sauvegarde de la configuration : %v", err)
	}

	return cfg
}

func saveConfig(config *Config, filePath string) error {
	file, err := os.Create(filePath)
	if err != nil {
		return fmt.Errorf("impossible de créer le fichier de configuration : %v", err)
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(config); err != nil {
		return fmt.Errorf("impossible d'écrire dans le fichier de configuration : %v", err)
	}
	return nil
}

// Charge les paramètres depuis un fichier de configuration
func loadConfig(filePath string) (*Config, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("impossible d'ouvrir le fichier de configuration : %v", err)
	}
	defer file.Close()

	var config Config
	decoder := json.NewDecoder(file)
	if err := decoder.Decode(&config); err != nil {
		return nil, fmt.Errorf("impossible de lire le fichier de configuration : %v", err)
	}
	return &config, nil
}
