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

type ItemOrderInfo struct {
	DiscordId  string
	Email      string
	FamilyName string
	FirstName  string
	ItemName   string
	Size       string
	Color      string
	Profile    string
}

func main() {
	if len(os.Args) != ValidArgsNum {
		exe, err := os.Executable()
		if err != nil {
			failOnError("exeファイル実行パス取得失敗", err)
		}
		exeName := filepath.Base(exe)
		failOnError(
			fmt.Sprintf(
				"%sにTXTファイル、アドレスExcelファイル、weeek Excelファイルを同時にドラッグ&ドロップしてください",
				exeName),
			nil)
	}

	profileOrderItemMap := getProfileItemMapFromTxt(findTxtFilePathFromArgs(os.Args))
	//profileOrderItemMap := getProfileOrderItemMapFromTxt("./testdata/7ec86e09b001b481.txt")
	emailOrderItemListMap := getMailOrderItemMapFromAddressExcel(findAddressExcelFilePathFromArgs(os.Args), profileOrderItemMap)
	//emailOrderItemListMap := getMailOrderItemMapFromAddressExcel("./testdata/address.xlsx", profileOrderItemMap)
	discordIdOrderItemListMap := getDiscordIdOrderItemListMap(findWeekExcelFilePathFromArgs(os.Args), emailOrderItemListMap)
	//discordIdOrderItemListMap := getDiscordIdOrderItemListMap("./testdata/20SS_week6.xlsx", emailOrderItemListMap)
	dumpOutputCsv(discordIdOrderItemListMap)
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

func dumpOutputCsv(discordIdItemListMap map[string][]*ItemOrderInfo) {
	exe, err := os.Executable()
	if err != nil {
		failOnError("exeファイル実行パス取得失敗", err)
	}

	outputDirPath := filepath.Dir(exe)
	outputCsv, err := os.OpenFile(outputDirPath+"/output.csv", os.O_WRONLY|os.O_CREATE, 0600)
	if err != nil {
		failOnError("output.csvのオープンに失敗しました", err)
	}
	defer outputCsv.Close()

	err = outputCsv.Truncate(0) // ファイルを空っぽにする(実行2回目以降用)
	if err != nil {
		failOnError("output.csvの初期化に失敗しました", err)
	}

	for discordId, itemList := range discordIdItemListMap {
		for i, orderItem := range itemList {
			if i == 0 {
				outputCsv.WriteString(discordId)
			}
			outputCsv.WriteString(fmt.Sprintf(",%s,%s,%s,%s,%s,%s,%s\n", orderItem.ItemName, orderItem.Size, orderItem.Color, orderItem.Profile, orderItem.Email, orderItem.FamilyName, orderItem.FirstName))
		}
	}
	fmt.Println(outputDirPath + "\\output.csvを出力しました")
}

func getDiscordIdOrderItemListMap(excelFilePath string, emailOrderItemListMap map[string][]*ItemOrderInfo) map[string][]*ItemOrderInfo {
	discordIdOrderItemListMap := make(map[string][]*ItemOrderInfo)
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

		orderItemInfoList, doesExist := emailOrderItemListMap[email]
		if !doesExist {
			continue
		}

		familyName := row.Cells[1].String()
		firstName := row.Cells[2].String()
		for _, orderItemInfo := range orderItemInfoList {
			orderItemInfo.FamilyName = familyName
			orderItemInfo.FirstName = firstName
		}

		discordId := row.Cells[20].String()
		discordIdOrderItemList, doesExist := discordIdOrderItemListMap[discordId]
		if !doesExist {
			discordIdOrderItemList = []*ItemOrderInfo{}
		}
		discordIdOrderItemList = append(discordIdOrderItemList, orderItemInfoList...)
		discordIdOrderItemListMap[discordId] = discordIdOrderItemList
	}

	return discordIdOrderItemListMap
}

func getMailOrderItemMapFromAddressExcel(excelFilePath string, profileItemMap map[string]*ItemOrderInfo) map[string][]*ItemOrderInfo {
	emailOrderItemListMap := make(map[string][]*ItemOrderInfo)
	excel, err := xlsx.OpenFile(excelFilePath)
	if err != nil {
		failOnError(fmt.Sprintf("%sのオープンに失敗", excelFilePath), err)
	}

	sheet := excel.Sheets[0]
	for _, row := range sheet.Rows {
		profile := row.Cells[18].String()
		orderItem, doesExist := profileItemMap[profile]
		if !doesExist {
			continue
		}

		email := row.Cells[24].String()
		orderItem.Email = email
		orderItemList, doesExist := emailOrderItemListMap[email]
		if !doesExist {
			orderItemList = []*ItemOrderInfo{}
		}

		orderItemList = append(orderItemList, orderItem)
		emailOrderItemListMap[email] = orderItemList
	}

	return emailOrderItemListMap
}

func getProfileItemMapFromTxt(txtFilePath string) map[string]*ItemOrderInfo {
	profileItemMap := make(map[string]*ItemOrderInfo)
	re := regexp.MustCompile("^Profile[0-9]+$")
	txtFile, err := os.Open(txtFilePath)
	if err != nil {
		failOnError("txtファイルオープンエラー", err)
	}
	defer txtFile.Close()

	scanner := bufio.NewScanner(txtFile)
	isItemLine := false
	isColorLine := false
	isSizeLine := false
	var itemInfo *ItemOrderInfo

	for scanner.Scan() {
		line := scanner.Text()
		if isItemLine {
			itemInfo.ItemName = line
			isItemLine = false
			isColorLine = true

		} else if isColorLine {
			itemInfo.Color = line
			isColorLine = false
			isSizeLine = true

		} else if isSizeLine {
			itemInfo.Size = line
			profileItemMap[itemInfo.Profile] = itemInfo
			isSizeLine = false

		} else if re.MatchString(line) {
			itemInfo = &ItemOrderInfo{}
			itemInfo.Profile = line
			isItemLine = true

		}
	}

	if err := scanner.Err(); err != nil {
		failOnError("txtファイル読み込みエラー", err)
	}

	return profileItemMap
}
