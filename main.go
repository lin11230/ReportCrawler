// Package Main implements OVID report crawler
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"os/exec"
	"regexp"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

var (
	lastyear bool
	chtMap   = map[string]string{}
	m        = time.Now().Month()
)

func flags() {

	flag.BoolVar(&lastyear, "lastyear", false, `Flag for Last Year, Default is: false`)

}

func init() {

	flags()
	flag.Parse()

	log.SetOutput(os.Stdout)
	log.SetFlags(log.LstdFlags | log.Lshortfile)

	// 指令範例 LASTYEAR=true ./ReportCrawler "網址"
	// $ LASTYEAR=true ./ReportCrawler "https://ovidspstats.ovid.com/scripts/osp.wsc/osp_getReport2.p?Metho=2014801.html&WhichReport=leOaijbicaaMhjdb"
	if os.Getenv("LASTYEAR") != "" {
		lastyear = true
	}

	if lastyear {
		fmt.Println("Running in LAST YEAR mode")
	}
}

// Func main 可傳入網址做為參數
func main() {
	//argUrl := "https://ovidspstats.ovid.com/scripts/osp.wsc/osp_getReport2.p?Method=1510332.html&WhichReport=nfnjakmlRphijTrl"

	argUrl := os.Args
	if len(argUrl) <= 1 {
		fmt.Println("請輸入參數網址")
		os.Exit(1)
	}
	//fmt.Println("arg0: ", argUrl[0])
	//fmt.Println("arg1: ", argUrl[1])
	//fmt.Println("")

	prepareChtMapFile()

	mbrMap := make(map[string]string)
	mbrMap = getMemberList(argUrl[1])

	usageMap := make(map[string]string)
	usageMap = getMemberUsageList(mbrMap)

	genExcelFile(usageMap)

}
func prepareChtMapFile() {
	var c interface{}
	csbytes, err := ioutil.ReadFile("clients.list")
	if err != nil {
		fmt.Println("clients.list 檔案開啟失敗. ", err)
	}

	err = json.Unmarshal(csbytes, &c)
	if err != nil {
		fmt.Println("json 格式錯誤，解析失敗. ", err)
	}

	cs := c.(map[string]interface{})
	for k, v := range cs {
		//fmt.Println("k is ", k, "; v is ", v)
		chtMap[k] = v.(string)
	}
}
func getMemberList(strURL string) map[string]string {
	r, _ := regexp.Compile("<OPTION VALUE=\"(.*)\" >(.*)</OPTION>")
	afmap := make(map[string]string)

	client := &http.Client{}

	req, err := http.NewRequest("GET", strURL, nil)
	req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:11.0) Gecko/20100101 Firefox/11.0")
	req.Header.Set("http.accept_language", "zh-tw,en-us;q=0.7,en;q=0.3")
	req.Header.Set("http.accept_encoding", "gzip, deflate")
	req.Header.Set("http.accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
	req.Header.Add("Referer", strURL)

	res, err := client.Do(req)
	if err != nil {
		log.Fatal(err)
	}

	robots, err := ioutil.ReadAll(res.Body)
	res.Body.Close()

	if err != nil {
		log.Fatal(err)
	}
	content := string(robots)
	content = strings.Replace(content, "selected", "", -1)
	content = strings.Replace(content, "&amp;", "&", -1)
	content = strings.Replace(content, "&#39;", "'", -1)

	foundmb := r.FindAllStringSubmatch(content, -1)
	for _, v := range foundmb {
		afmap[v[2]] = v[1]
	}

	return afmap
}

func getMemberUsageList(m map[string]string) map[string]string {
	usageMap := make(map[string]string)
	for k, v := range m {
		fmt.Println("")
		fmt.Println("processing [", k, "] data....")
		//fmt.Println("data v is ", v)
		usage := getUsage(v)
		usageMap[k] = usage
	}
	return usageMap
}

func getUsage(para string) string {
	comboLink := "https://ovidspstats.ovid.com/scripts/osp_mail.wsc/osp_getReport2.p?" + para
	//r, _ := regexp.Compile("<TD ALIGN=\"RIGHT\" NOWRAP>(.*)</TD>")
	r, err := regexp.Compile("(?i)<TD ALIGN=\"RIGHT\" NOWRAP>(.*)</TD>")
	if err != nil {
		fmt.Println("getUsage regexp fail! ", err)
	}

	client := &http.Client{}

	req, err := http.NewRequest("GET", comboLink, nil)
	req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:11.0) Gecko/20100101 Firefox/11.0")
	req.Header.Set("http.accept_language", "zh-tw,en-us;q=0.7,en;q=0.3")
	req.Header.Set("http.accept_encoding", "gzip, deflate")
	req.Header.Set("http.accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
	req.Header.Add("Referer", comboLink)

	res, err := client.Do(req)
	if err != nil {
		log.Fatal(err)
	}

	robots, err := ioutil.ReadAll(res.Body)
	res.Body.Close()

	if err != nil {
		log.Fatal(err)
	}
	content := string(robots)
	//fmt.Println("content is", content)
	content = strings.Replace(content, "selected", "", -1)
	content = strings.Replace(content, "&amp;", "&", -1)
	content = strings.Replace(content, "&#39;", "'", -1)
	content = strings.ToUpper(content)

	//m := time.Now().Month()
	var foundUsage [][]string
	if lastyear {
		foundUsage = r.FindAllStringSubmatch(content, 13)
	} else {
		foundUsage = r.FindAllStringSubmatch(content, int(m))
	}

	strUsage := ""
	for _, v := range foundUsage {
		//fmt.Println("value in cell", v[1])
		strUsage += strings.Replace(v[1], "&NBSP;", "0", -1) + ","
	}
	//log.Println("strUsage is:", strUsage)

	return strUsage

}
func genExcelFile(umap map[string]string) {

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("sheet1")
	if err != nil {
		fmt.Println("Add Sheet error!: ", err)
	}

	type interfaceA []interface{}
	title := interfaceA{"英文名", "中文名", "總計", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"}
	row = sheet.AddRow()
	row.WriteSlice(&title, -1)

	for k, v := range umap {
		//fmt.Println(k, ", ", chtMap[k], ",", v)
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = k
		cell = row.AddCell()
		cell.Value = chtMap[k]
		arrv := strings.Split(v, ",")
		for _, vv := range arrv {
			cell = row.AddCell()
			cell.Value = vv
		}
	}

	year, month, day := time.Now().Date()
	filename := fmt.Sprintf("/tmp/OVID_BR1_%d%02d%02d.xlsx", year, month, day)
	fmt.Println("寫入檔案至路徑：", filename)
	err = file.Save(filename)
	if err != nil {
		fmt.Println("檔案寫入失敗！  ", err)
	}

	fmt.Println("開啟 excel 檔案...")

	cmd := exec.Command("open", filename)
	err = cmd.Run()
	if err != nil {
		fmt.Println("檔案開啟失敗！ ", err)
	}
}
