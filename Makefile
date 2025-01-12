# Version à définir ici
VERSION ?= 1.0.0

# Date de compilation
BUILD_DATE := $(shell date -u +"%Y-%m-%d")

# Chemins pour les outils
WINDRES = x86_64-w64-mingw32-windres

# Fichiers sources
SRC_DIR = .
OUTPUT_DIR = ./output
EXE_NAME = FredonToPDF.exe

# Fichier RC et ICÔNE
RC_FILE = resources.rc
ICON_FILE = res/icon.ico

# Cibles par défaut
all: build

# Cible pour créer l'exécutable pour Windows
build: resources.syso $(SRC_DIR)/*.go
	@echo "Compilation de l'application Go pour Windows..."
	GOOS=windows GOARCH=amd64 go build -o $(OUTPUT_DIR)/$(EXE_NAME) \
		-ldflags "-X main.version=$(VERSION) -X main.BuildDate=$(BUILD_DATE)" .

# Cible pour générer le fichier .syso à partir du fichier .rc (ressources)
resources.syso: $(RC_FILE)
	@echo "Génération du fichier de ressources..."
	$(WINDRES) -o resources.syso $(RC_FILE)

# Cible pour nettoyer les fichiers générés
clean:
	@echo "Nettoyage des fichiers générés..."
	rm -f $(OUTPUT_DIR)/$(EXE_NAME) resources.syso

# Cible pour afficher la version
version:
	@echo "Version : $(VERSION)"
