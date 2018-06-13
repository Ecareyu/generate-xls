package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/go-redis/redis"
	yaml "gopkg.in/yaml.v2"
)

// 队列名称
const JobKey string = "go:event:generate-xlsx"

// 配置文件结构
type Config struct {
	// Redis配置
	Redis RedisConf
}

type RedisConf struct {
	Host     string
	Port     string
	Password string
}

type Job struct {
	// 表头
	Head struct {
		Cells []struct {
			K string
			V string
		}
		Merges []struct {
			H string
			V string
		}
	}

	Data struct {
		Url   string
		Total uint
		Take  uint
	}
}

func main() {
	// 载入配置项
	var c Config
	c.loadConf()

	// 连接到redis客户端
	// rc := redisClient(c.Redis)

	// rc.LPush(JobKey, 12312312)

	// value := rc.LPop(JobKey)
	// fmt.Println(value)
	// 读取消息任务信息
	var job Job
	job.parseJob()

	processJob(&job)

	// fmt.Println(job)

	// writeExcel()

	// fmt.Println(c.Redis.Port)
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

// 解析任务数据
func (j *Job) parseJob() *Job {
	jsonStream := `{"head":{"cells":[{"k":"A1","v":"\u5e8f\u53f7"},{"k":"B1","v":"\u7701\u4efd"},{"k":"C1","v":"\u57ce\u5e02"},{"k":"D1","v":"\u95e8\u5e97\u5c5e\u6027"},{"k":"E1","v":"\u95e8\u5e97\u540d\u79f0"},{"k":"F1","v":"\u672c\u5468\u6267\u884c\u7edf\u8ba1"},{"k":"K1","v":"\u5b9e\u9645\u6267\u884c\u603b\u8ba1"},{"k":"F2","v":"\u6267\u884c\u573a\u5730"},{"k":"G2","v":"\u9884\u7ea6\u4eba\u6570"},{"k":"H2","v":"\u9884\u7ea6\u65f6\u95f4"},{"k":"I2","v":"\u7b2c\u4e09\u65b9\u5230\u573a\u524d\u8054\u7cfb\u7535\u8bdd"},{"k":"J2","v":"\u573a\u6b21\u7b49\u7ea7"},{"k":"K2","v":"\u662f\u5426\u6b63\u5e38\u6267\u884c"},{"k":"L2","v":"\u5b9e\u53d1\u4eba\u6570"}],"merges":[{"h":"A1","v":"A2"},{"h":"B1","v":"B2"},{"h":"C1","v":"C2"},{"h":"D1","v":"D2"},{"h":"E1","v":"E2"},{"h":"F1","v":"J1"},{"h":"K1","v":"L1"}]},"data":{"url":"http:\/\/foo.com\/getData","total":300,"take":1000}}`

	dec := json.NewDecoder(strings.NewReader(jsonStream))

	err := dec.Decode(&j)
	if err != nil {
		fmt.Println(err)
	}

	return j
}

// 写excel
func processJob(job *Job) {
	sheet := "sheet1"

	// 创建excel，然后去配置
	xlsx := excelize.NewFile()
	index := xlsx.NewSheet(sheet)

	// 表格头的值
	headCells := job.Head.Cells

	// 需要合并的单元格值
	mergeCells := job.Head.Merges

	// 设置表格头内容
	for _, cell := range headCells {
		xlsx.SetCellValue(sheet, cell.K, cell.V)
	}

	// 合并表格头
	for _, cell := range mergeCells {
		xlsx.MergeCell(sheet, cell.H, cell.V)
	}

	xlsx.SetActiveSheet(index)

	err := xlsx.SaveAs("./Book1.xlsx")
	if err != nil {
		fmt.Println(err)
	}

}

func requestData(jd *Job.Data) {
	jd.Url
}

// redis客户端
func redisClient(rc RedisConf) *redis.Client {
	client := redis.NewClient(&redis.Options{
		Addr:     rc.Host + ":" + rc.Port,
		Password: rc.Password,
		DB:       0,
	})

	return client
}
