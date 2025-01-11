package helper

import (
	"fmt"
	"os"
	"strings"
	"time"
)

const logFormat string = "2006-01-02 15:04:05"

var colors = map[string]int8{
	"info":    32,
	"warning": 33,
	"fatal":   31,
}

func GLog(msg string, msgType string, newLine bool) {
	dt := time.Now()
	nl := ""
	if newLine {
		nl = "\n"
	}
	fmt.Printf("[%s] [ \x1b[%d;1m%s\x1b[0m ] %s%s", dt.Format(logFormat), colors[msgType], strings.ToUpper(msgType), msg, nl)
}

func GInfo(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "info", false)
}

func GWarning(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "warning", false)
}

func GFatal(format string, args ...interface{}) {
	GLog(fmt.Sprintf(format, args...), "fatal", false)
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
	os.Exit(1)
}

func GBlank() {
	fmt.Println("")
}
