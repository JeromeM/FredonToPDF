package helper

import (
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/gookit/color"
)

const logFormat string = "2006-01-02 15:04:05"

var colors = map[string]func(a ...any) string{
	"info":    color.FgGreen.Render,
	"warning": color.FgYellow.Render,
	"fatal":   color.FgRed.Render,
}

func GLog(msg string, msgType string, newLine bool) {
	dt := time.Now()
	nl := ""
	if newLine {
		nl = "\n"
	}
	fmt.Printf("[%s] [ %s ] %s%s", dt.Format(logFormat), colors[msgType](strings.ToUpper(msgType)), msg, nl)
}

func GInfo(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "info", false)
}

func GWarning(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "warning", false)
}

func GFatal(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "fatal", false)
	GInfoLn("Appuyez sur une touche pour fermer...")
	fmt.Scanln()
	os.Exit(1)
}

func GInfoLn(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "info", true)
}

func GWarningLn(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "warning", true)
}

func GFatalLn(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "fatal", true)
	GInfoLn("Appuyez sur une touche pour fermer...")
	fmt.Scanln()
	os.Exit(1)
}

func GBlank() {
	fmt.Println("")
}
