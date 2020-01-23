package main

import (
	"fmt"
	"io/ioutil"
	"gopkg.in/yaml.v2"
)

const (
	DATA_SHEET = "Sheet"
	REPORT_SHEET = "Report"
)

type config struct {
	Score string
	ScoreWidth float64 `yaml:"score_width"`
	StartPoll string `yaml:"start_poll"`
	Colls []struct {
		Source string
		Target string
		Width float64
	}
	//dateCreated int
	//date_updated int

}

func (conf *config) load(file string)  {
	cf, err := ioutil.ReadFile(file)
    if err != nil {
        panic(fmt.Sprintf("Error on reading config.yaml file: %v",err))
	}
	
    err = yaml.Unmarshal(cf, &conf)
    if err != nil {
		panic(fmt.Sprintf("Error while processing config file: %v", err))
    }
}