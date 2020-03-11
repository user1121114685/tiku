package main

import (
	"bufio"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gogf/gf/text/gregex"
)

func main() {
	// 创建新的题库excel
	newExcel := excelize.NewFile()
	newExcel.SetCellValue("Sheet1", "A1", "Topic")      // 创建首行db关键字 题目
	newExcel.SetCellValue("Sheet1", "B1", "OptionList") // 创建首行db关键字 选项
	newExcel.SetCellValue("Sheet1", "C1", "Result")     // 创建首行db关键字 答案
	newExcel.SetCellValue("Sheet1", "D1", "Explain")    // 创建首行db关键字 解析
	// 逐行扫描txt，将他们写入excel
	file, err := os.Open("./导入文本.txt")
	if err != nil {
		println(err.Error())
	}
	defer file.Close()

	LineExcel := 2 // 给excel 赋值正在使用的行号赋值
	Scanner := bufio.NewScanner(file)
	for Scanner.Scan() {

		var Txts string
		Txts = Scanner.Text()

		for strings.Contains(Txts, "[填空题]") { // 如果字符串里面包含了[填空题]，就进行下面的

			for Scanner.Scan() {

				Txts = Scanner.Text()

				if strings.Contains(Txts, "[选择题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				if strings.Contains(Txts, "[判断题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[简答题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				LineExcelString := strconv.Itoa(LineExcel)
				newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts)
				LineExcel++

				println(Txts)
			}
		}

		var TxtsA string // 多选 单选 选项A
		var TxtsB string
		var TxtsC string
		var TxtsD string

		for strings.Contains(Txts, "[选择题]") { // 如果字符串里面包含了[填空题]，就进行下面的
			for Scanner.Scan() {

				Txts = Scanner.Text()

				LineExcelString := strconv.Itoa(LineExcel)
				// var TxtsA string
				// var TxtsB string
				// var TxtsC string
				// var TxtsD string
				if strings.Contains(Txts, "[填空题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[判断题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[简答题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				if strings.Contains(Txts, "A.") && strings.Contains(Txts, "B.") && strings.Contains(Txts, "C.") && strings.Contains(Txts, "D.") { // 如果字段包含A B C D 就当成是选项
					if strings.Contains(Txts, "ABCD") {
						newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts)
						LineExcel--
					}
					Txts = strings.Replace(Txts, " B.", ";;B.", -1)
					Txts = strings.Replace(Txts, " C.", ";;C.", -1)
					Txts = strings.Replace(Txts, " D.", ";;D.", -1) // 替换字符串
					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Txts)
					LineExcel++
					println(Txts)
				} else if strings.Contains(Txts, "A.") { // 如果包含A.

					TxtsA = Txts // 将结果存入txtsA中

					Txts = Scanner.Text()
					println(Txts)
				} else if strings.Contains(Txts, "B.") { // 如果包含B.
					TxtsB = Txts

					Txts = Scanner.Text()
					println(Txts)
				} else if strings.Contains(Txts, "C.") { // 如果包含B.
					TxtsC = Txts

					Txts = Scanner.Text()
					println(Txts)
				} else if strings.Contains(Txts, "D.") { // 如果包含B.
					TxtsD = Txts
					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, TxtsA+";;"+TxtsB+";;"+TxtsC+";;"+TxtsD)
					LineExcel++
				} else {

					// var ResultReg *regexp.Regexp
					match, _ := gregex.MatchString(`\(.*[A-D]+.*\)`, Txts) // 正则表达式中的* 表示前一个字符出现任意次，与我们所谓的*匹配任意字符不同  参见 https://zh.wikipedia.org/wiki/%E6%AD%A3%E5%88%99%E8%A1%A8%E8%BE%BE%E5%BC%8F
					if match != nil {
						result := match[len(match)-1]

						// Result := reg.FindAllString(Txts, -1)
						// ResultReg = regexp.MustCompile(`(\(|\ |\))`) // 查找 括号 空格 反括弧
						// Result = ResultReg.ReplaceAllString(result_1, "")
						resultReg, _ := gregex.MatchString(`[A-D]+`, result)
						newExcel.SetCellValue("Sheet1", "C"+LineExcelString, resultReg[0]) // 写入选择题答案
						Txts, err = gregex.ReplaceString(result, " ", Txts)                // 参数含义 需要替换的字符  替换后的字符 目标处理文件
						if err != nil {
							println(err.Error())
						}
					}
					newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts) // 题目不换行

				}
			}

		}

		for strings.Contains(Txts, "[判断题]") { // 如果字符串里面包含了[填空题]，就进行下面的

			for Scanner.Scan() {

				Txts = Scanner.Text()
				if strings.Contains(Txts, "[选择题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				if strings.Contains(Txts, "[判断题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[简答题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				LineExcelString := strconv.Itoa(LineExcel)
				if strings.Contains(Txts, "√") {
					newExcel.SetCellValue("Sheet1", "C"+LineExcelString, "Y")
				}
				if strings.Contains(Txts, "×") {
					newExcel.SetCellValue("Sheet1", "C"+LineExcelString, "N")
				}
				Txts, err = gregex.ReplaceString("√|×", " ", Txts) // 参数含义 需要替换的字符  替换后的字符 目标处理文件
				if err != nil {
					println(err.Error())
				}
				newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts)
				LineExcel++

				println(Txts)
			}
		}
		//println(Scanner.Text())
		// fmt.Printf("值类型为 :%T\n", Scanner.Text())
		// println(Scanner.Text())
		//	lines = append(lines, Scanner.Text()) // append函数是用来在slice末尾追加一个或者多个元素。

	}

	// 获取当前时间
	t := time.Now() //2019-07-31 13:55:21.3410012 +0800 CST m=+0.006015601

	if err := newExcel.SaveAs("联盟少侠" + t.Format("20060102150405") + ".xlsx"); err != nil {
		println(err.Error())
	}
}
