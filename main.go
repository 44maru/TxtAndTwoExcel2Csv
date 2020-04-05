package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"regexp"

	"github.com/tealeg/xlsx"
)

const ValidArgsNum = 4

type Account struct {
	Id        int
	AccountId string
	LoginId   string
	Password  string
	Retailer  string
}

func main() {
	if len(os.Args) != ValidArgsNum {
		failOnError("main.exeにTXTファイル、アドレスExcelファイル、weeek Excelファイルを同時にドラッグ&ドロップしてください", nil)
	}

	profileItemMap := getProfileItemMapFromTxt(findTxtFilePathFromArgs(os.Args))
	//profileItemMap := getProfileItemMapFromTxt("./testdata/7ec86e09b001b481.txt")
	emailItemListMap := getMailFromAddressExcel(findAddressExcelFilePathFromArgs(os.Args), profileItemMap)
	//emailItemListMap := getMailFromAddressExcel("./testdata/address.xlsx", profileItemMap)
	discordIdItemListMap := getDiscordIdItemListMap(findWeekExcelFilePathFromArgs(os.Args), emailItemListMap)
	//discordIdItemListMap := getDiscordIdItemListMap("./testdata/20SS_week6.xlsx", emailItemListMap)
	dumpOutputCsv(discordIdItemListMap)
	//convertCsv2Json(os.Args[1])
	waitEnter()
}

func findTxtFilePathFromArgs(args []string) string {
	for i := 1; i < ValidArgsNum; i++ {
		ext := filepath.Ext(args[i])
		if ext == ".txt" {
			return args[i]
		}
	}
	failOnError("入力ファイルからTXTファイルが見つかりませんでした。", nil)

	return ""
}

func findAddressExcelFilePathFromArgs(args []string) string {
	for i := 1; i < ValidArgsNum; i++ {
		fileName := filepath.Base(args[i])
		if fileName == "address.xls" || fileName == "address.xlsx" || fileName == "address.xlsm" {
			return args[i]
		}
	}
	failOnError("入力ファイルからaddress.xlsxが見つかりませんでした。", nil)

	return ""
}

func findWeekExcelFilePathFromArgs(args []string) string {
	for i := 1; i < ValidArgsNum; i++ {
		ext := filepath.Ext(args[i])
		if ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" {
			continue
		}

		fileName := filepath.Base(args[i])
		if fileName != "address.xls" && fileName != "address.xlsx" && fileName != "address.xlsm" {
			return args[i]
		}
	}
	failOnError("入力ファイルからweekのExcelファイルが見つかりませんでした。", nil)

	return ""
}

func failOnError(errMsg string, err error) {
	//errs := errors.WithStack(err)
	fmt.Println(errMsg)
	if err != nil {
		//fmt.Printf("%+v\n", errs) Stack trace
		fmt.Printf("%s\n", err.Error())
	}
	waitEnter()
	os.Exit(1)
}

func waitEnter() {
	fmt.Println("エンターを押すと処理を終了します。")
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
}

func dumpOutputCsv(discordIdItemListMap map[string][]string) {
	exe, err := os.Executable()
	if err != nil {
		failOnError("exeファイル実行パス取得失敗", err)
	}

	outputDirPath := filepath.Dir(exe)
	outputCsv, err := os.OpenFile(outputDirPath+"/output.csv", os.O_WRONLY|os.O_CREATE, 0600)
	if err != nil {
		failOnError("Account.jsonのオープンに失敗しました", err)
	}
	defer outputCsv.Close()

	err = outputCsv.Truncate(0) // ファイルを空っぽにする(実行2回目以降用)
	if err != nil {
		failOnError("output.csvの初期化に失敗しました", err)
	}

	for discordId, itemList := range discordIdItemListMap {
		for i, item := range itemList {
			if i == 0 {
				outputCsv.WriteString(fmt.Sprintf("%s,%s\n", discordId, item))
			} else {
				outputCsv.WriteString(fmt.Sprintf(",%s\n", item))
			}
		}
	}
	fmt.Println(outputDirPath + "\\output.csvを出力しました")
}

func getDiscordIdItemListMap(excelFilePath string, emailItemListMap map[string][]string) map[string][]string {
	discordIdItemListMap := make(map[string][]string)
	excel, err := xlsx.OpenFile(excelFilePath)
	if err != nil {
		failOnError(fmt.Sprintf("%sのオープンに失敗", excelFilePath), err)
	}

	alreadyFoundEmail := make(map[string]bool)
	sheet := excel.Sheets[0]
	for _, row := range sheet.Rows {
		email := row.Cells[8].String()

		if _, doesExist := alreadyFoundEmail[email]; doesExist {
			continue
		}
		alreadyFoundEmail[email] = true

		emailItemList, doesExist := emailItemListMap[email]
		if !doesExist {
			continue
		}

		discordId := row.Cells[20].String()
		discordIdItemList, doesExist := discordIdItemListMap[discordId]
		if !doesExist {
			discordIdItemList = []string{}
		}
		discordIdItemList = append(discordIdItemList, emailItemList...)
		discordIdItemListMap[discordId] = discordIdItemList
	}

	return discordIdItemListMap
}

func getMailFromAddressExcel(excelFilePath string, profileItemMap map[string]string) map[string][]string {
	emailItemListMap := make(map[string][]string)
	excel, err := xlsx.OpenFile(excelFilePath)
	if err != nil {
		failOnError(fmt.Sprintf("%sのオープンに失敗", excelFilePath), err)
	}

	sheet := excel.Sheets[0]
	for _, row := range sheet.Rows {
		profile := row.Cells[18].String()
		item, doesExist := profileItemMap[profile]
		if !doesExist {
			continue
		}

		email := row.Cells[24].String()
		itemList, doesExist := emailItemListMap[email]
		if !doesExist {
			itemList = []string{}
		}

		itemList = append(itemList, item)
		emailItemListMap[email] = itemList
	}

	return emailItemListMap
}

func getProfileItemMapFromTxt(txtFilePath string) map[string]string {
	profileItemMap := make(map[string]string)
	re := regexp.MustCompile("^Profile[0-9]+$")
	txtFile, err := os.Open(txtFilePath)
	if err != nil {
		failOnError("txtファイルオープンエラー", err)
	}
	defer txtFile.Close()

	scanner := bufio.NewScanner(txtFile)
	isItemLine := false
	profile := ""
	for scanner.Scan() {
		line := scanner.Text()
		if isItemLine {
			profileItemMap[profile] = line
			isItemLine = false

		} else if re.MatchString(line) {
			profile = line
			isItemLine = true

		} else {
			isItemLine = false
		}
	}

	if err := scanner.Err(); err != nil {
		failOnError("txtファイル読み込みエラー", err)
	}

	return profileItemMap
}
