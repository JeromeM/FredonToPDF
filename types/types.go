package types

// ProcessResult représente le résultat du traitement d'un fichier
type ProcessResult struct {
	FileName string
	PdfPath  string
	Err      error
}
