package main

import (
	"fmt"
	"os"
	"io/ioutil"
	"path/filepath"
)


func main() {

	configFileName := "config.yaml"
	if len(os.Args) > 1 {
		configFileName = os.Args[1]
	}

	var conf config
	conf.load(configFileName)

	const EXCEL_EXT = ".xlsx"
	files, err := ioutil.ReadDir("./")
    if err != nil {
        panic(err.Error())
	}
	
	for _, f := range files {
		if f.IsDir() || filepath.Ext(f.Name()) != EXCEL_EXT {
			continue
		}

		println(fmt.Sprintf("--------- // %s // ----------", f.Name()))

		create_tealeg_report(&conf, f.Name())

		println("------------------------------")
	}
}


