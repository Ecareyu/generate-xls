package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"path/filepath"

	// "github.com/go-redis/redis"
	"github.com/360EntSecGroup-Skylar/excelize"
	yaml "gopkg.in/yaml.v2"
)

// 队列名称
const QUEUE_KEY string = "event:generate-xlsx:go"

// Config配置文件
type Config struct {
	// Redis配置
	Redis struct {
		Host     string
		Port     string
		Password string
	}
}

type Queue struct {
	// 表头
	Titles []string

	// 内容url
	Resource string
}

func main() {
	// 载入配置项
	var c Config
	c.loadConf()

	// writeExcel()

	fmt.Println(c.Redis.Port)
}

// 自动载入配置
func (c *Config) loadConf() *Config {
	content, err := ioutil.ReadFile(filepath.Base("./conf.yaml"))
	if err != nil {
		log.Fatal(err)
	}

	yaml.Unmarshal(content, &c)

	return c
}

func nextJob() {

}

// 写excel
func writeExcel() {
	// 创建excel，然后去配置
	xlsx := excelize.NewFile()
	index := xlsx.NewSheet("Sheet2")

	xlsx.SetCellValue("Sheet2", "A2", "Hello World")
	xlsx.SetCellValue("Sheet1", "B2", 100)
	xlsx.SetActiveSheet(index)

	err := xlsx.SaveAs("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
