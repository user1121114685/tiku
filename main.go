package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

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
	newExcel.SetCellValue("Sheet1", "E1", "Type")       // 创建首行db关键字 题型 pd 判断  dx 单选 dd 多选 tk 填空 jd 简答
	// 逐行扫描txt，将他们写入excel
	file, err := os.Open("./导入文本.txt")
	if err != nil {
		println(err.Error())
	}
	defer file.Close()

	LineExcel := 2 // 给excel 赋值正在使用的行号赋值
	Answer := ""   // 简答题/选择题 答案 变量

	Scanner := bufio.NewScanner(file)
	for Scanner.Scan() {

		var Txts string
		Txts = Scanner.Text()

		for strings.Contains(Txts, "[填空题]") { // 如果字符串里面包含了[填空题]，就进行下面的
			Answer = "" // 将答案清空，避免其他问题导致答案错误

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
				if strings.Contains(Txts, "[填空题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				LineExcelString := strconv.Itoa(LineExcel) // 数字类型转字符串类型
				// 录入解析
				if strings.Contains(Txts, "[解析]") { // 包含关键字 退出循环
					LineExcel-- // 切换到上一行
					LineExcelString = strconv.Itoa(LineExcel)
					newExcel.SetCellValue("Sheet1", "D"+LineExcelString, Txts)
					println("解析  在ECXEL的：" + LineExcelString + " 行 " + Txts)
					LineExcel++
				}

				// 提取题目中的答案
				match, _ := gregex.MatchString(`\(\(.*\)\)`, Txts) // 正则表达式中的* 表示前一个字符出现任意次，与我们所谓的*匹配任意字符不同  参见 https://zh.wikipedia.org/wiki/%E6%AD%A3%E5%88%99%E8%A1%A8%E8%BE%BE%E5%BC%8F
				if match != nil {
					result := match[len(match)-1]
					// 排除((  )) 只提取里面的内容
					//resultReg, err := gregex.MatchString(`[^\(\)]`, result) // 排除型字符集合（negated character classes）。匹配未列出的任意字符。例如，“[^abc]”可以匹配“plain”中的“plin”。
					Answer, err = gregex.ReplaceString(`\(|\)`, "", result) // 参数含义 需要替换的字符  替换后的字符 目标处理文件
					if err != nil {
						println(err.Error())
					}

					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer) // 写入填空题答案
					Txts, err = gregex.ReplaceString(result, "", Txts)           // 参数含义 需要替换的字符  替换后的字符 目标处理文件
					if err != nil {
						println(err.Error())
					}
					println("本题填空题 答案：" + LineExcelString + " 行 " + Answer)
				} else {
					LineExcelString := strconv.Itoa(LineExcel)
					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, "本题答案为空")

				}
				if !strings.Contains(Txts, "[解析]") {

					newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts) // 题目
					newExcel.SetCellValue("Sheet1", "E"+LineExcelString, "tk") //题目类型
					println("填空题在ECXEL的：" + LineExcelString + " 行 " + Txts)

					LineExcel++
				}

			}
		}

		for strings.Contains(Txts, "[选择题]") { // 如果字符串里面包含了[填空题]，就进行下面的
			Answer = ""
			for Scanner.Scan() {

				Txts = Scanner.Text()

				LineExcelString := strconv.Itoa(LineExcel)

				if strings.Contains(Txts, "[选择题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				if strings.Contains(Txts, "[判断题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[简答题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				if strings.Contains(Txts, "[填空题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				if strings.Contains(Txts, "A.") && strings.Contains(Txts, "B.") && strings.Contains(Txts, "C.") { // 如果字段包含A B C D 就当成是选项
					LineExcel--
					Txts = strings.Replace(Txts, " B.", ";;B.", -1)
					Txts = strings.Replace(Txts, " C.", ";;C.", -1)
					Txts = strings.Replace(Txts, " D.", ";;D.", -1) // 替换字符串
					Txts = strings.Replace(Txts, " E.", ";;E.", -1) // 替换字符串
					Txts = strings.Replace(Txts, " F.", ";;F.", -1) // 替换字符串
					Txts = strings.Replace(Txts, " G.", ";;G.", -1) // 替换字符串
					LineExcelString = strconv.Itoa(LineExcel)
					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Txts)
					println("选项   在ECXEL的：" + LineExcelString + " 行 " + Txts)
					LineExcel++

				} else if strings.Contains(Txts, "A.") { // 如果包含A.
					LineExcel--
					if Answer == "" {
						Answer = Txts
					} else {
						Answer = Answer + ";;" + Txts // 将结果存入Answer中
					}

					// Txts = Scanner.Text()
					// println("选择题在ECXEL的：" + LineExcelString + " 行 " + Txts)

				} else if strings.Contains(Txts, "B.") { // 如果包含B.
					if Answer == "" {
						Answer = Txts
					} else {
						Answer = Answer + ";;" + Txts // 将结果存入Answer中
					}

					if strings.Contains(Answer, "A.") { // 如果包含B.

						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
						LineExcel++
					}

				} else if strings.Contains(Txts, "C.") { // 如果包含B.
					if Answer == "" {
						Answer = Txts
					} else {
						Answer = Answer + ";;" + Txts // 将结果存入Answer中
					}

					if strings.Contains(Answer, "B.") { // 如果包含B.

						LineExcel-- // 切换到上一行
						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
						LineExcel++
					}

				} else if strings.Contains(Txts, "D.") { // 如果包含B.
					if Answer == "" {
						Answer = Txts
					} else {
						Answer = Answer + ";;" + Txts // 将结果存入Answer中
					}
					//newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
					// println("选项  在ECXEL的：" + LineExcelString + " 行 " + Answer)
					if strings.Contains(Answer, "C.") { // 如果包含B.

						//LineExcel = LineExcel - 2 // 切换到题目行
						LineExcel--
						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
						LineExcel++
					}

				} else if strings.Contains(Txts, "E.") { // 如果包含B.

					Answer = Answer + ";;" + Txts       // 将结果存入Answer中
					if strings.Contains(Answer, "D.") { // 如果包含B.

						LineExcel--
						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
						LineExcel++
					}

				} else if strings.Contains(Txts, "F.") { // 如果包含B.

					Answer = Answer + ";;" + Txts // 将结果存入Answer中

					if strings.Contains(Answer, "E.") { // 如果包含B.

						LineExcel--
						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)
						LineExcel++
					}

				} else if strings.Contains(Txts, "G.") { // 如果包含B.

					LineExcel--
					LineExcelString = strconv.Itoa(LineExcel)
					Answer = Answer + ";;" + Txts // 将结果存入Answer中
					newExcel.SetCellValue("Sheet1", "B"+LineExcelString, Answer)

					LineExcel++

				} else {

					// 提取题目中的答案
					match, _ := gregex.MatchString(`\(\ *[A-G]+\ *[A-G]*\ *[A-G]*\ *[A-G]*\ *[A-G]*\ *[A-G]*\ *[A-G]*\ *[A-G]*\ *[A-G]*\)`, Txts) // 正则表达式中的* 表示前一个字符出现任意次，与我们所谓的*匹配任意字符不同  参见 https://zh.wikipedia.org/wiki/%E6%AD%A3%E5%88%99%E8%A1%A8%E8%BE%BE%E5%BC%8F
					if match != nil {
						result := match[len(match)-1]
						// resultAgain, _ := gregex.MatchString(`\(*[A-G]+.*\)`, result) // 部分题目中不止包含一次，所以需要再次赛选
						// result = resultAgain[len(resultAgain)-1]
						resultReg, err := gregex.ReplaceString(`\(|\)|\ `, "", result) // 参数含义 需要替换的字符  替换后的字符 目标处理文件
						if err != nil {
							println(err.Error())
						}

						newExcel.SetCellValue("Sheet1", "C"+LineExcelString, resultReg) // 写入选择题答案
						Txts, err = gregex.ReplaceString(result, " ", Txts)             // 参数含义 需要替换的字符  替换后的字符 目标处理文件
						if err != nil {
							println(err.Error())
						}
						if utf8.RuneCountInString(resultReg) > 1 { // 答案长度超过1 则为多选
							newExcel.SetCellValue("Sheet1", "E"+LineExcelString, "dd") //题目类型  多选
						} else {

							newExcel.SetCellValue("Sheet1", "E"+LineExcelString, "dx") //题目类型  单选

						}
						println("本题选择题 答案：" + LineExcelString + " 行 " + resultReg)
					}
					// 录入解析
					if strings.Contains(Txts, "[解析]") { // 包含关键字 退出循环
						LineExcel-- // 切换到上一行
						LineExcelString = strconv.Itoa(LineExcel)
						newExcel.SetCellValue("Sheet1", "D"+LineExcelString, Txts)
						println("解析  在ECXEL的：" + LineExcelString + " 行 " + Txts)
						LineExcel++

					}
					if !strings.Contains(Txts, "[解析]") { // 取反 ，如果不包含解析 就执行下面的命令
						newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts) // 题目不换行
						Answer = ""                                                // 遇到题目就清空答案
						println("选择题在ECXEL的：" + LineExcelString + " 行 " + Txts)
						LineExcel++
					}

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
				if strings.Contains(Txts, "[填空题]") { // 包含关键字 退出循环
					break // 退出循环

				}
				LineExcelString := strconv.Itoa(LineExcel)
				// 录入解析
				if strings.Contains(Txts, "[解析]") { // 包含关键字 退出循环
					LineExcel-- // 切换到上一行
					LineExcelString = strconv.Itoa(LineExcel)
					newExcel.SetCellValue("Sheet1", "D"+LineExcelString, Txts)
					println("解析  在ECXEL的：" + LineExcelString + " 行 " + Txts)

				}

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
				if !strings.Contains(Txts, "[解析]") {
					newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts)
					newExcel.SetCellValue("Sheet1", "E"+LineExcelString, "pd") //题目类型 判断
					println("判断题在ECXEL的：" + LineExcelString + " 行 " + Txts)
				}

				LineExcel++

			}
		}

		for strings.Contains(Txts, "[简答题]") { // 如果字符串里面包含了[填空题]，就进行下面的
			Answer = ""

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
				if strings.Contains(Txts, "[填空题]") { // 包含关键字 退出循环
					break // 退出循环

				}

				LineExcelString := strconv.Itoa(LineExcel)

				if strings.Contains(Txts, "[题目]") {

					Answer = ""

					Txts, err = gregex.ReplaceString(`\[题目\]`, " ", Txts) // 参数含义 需要替换的字符  替换后的字符 目标处理文件
					if err != nil {
						println(err.Error())
					}
					newExcel.SetCellValue("Sheet1", "A"+LineExcelString, Txts)
					newExcel.SetCellValue("Sheet1", "E"+LineExcelString, "jd") //题目类型 简答
					println("简答题在ECXEL的：" + LineExcelString + " 行 " + Txts)
					LineExcel++
				} else {
					Answer = Answer + "\n" + Txts
					LineExcel--
					LineExcelString = strconv.Itoa(LineExcel)
					newExcel.SetCellValue("Sheet1", "D"+LineExcelString, Answer)
					println("答案   在ECXEL的：" + LineExcelString + " 行 " + Answer)
					LineExcel++

				}

			}
			LineExcel++
		}

	}

	// 获取当前时间
	t := time.Now() //2019-07-31 13:55:21.3410012 +0800 CST m=+0.006015601

	if err := newExcel.SaveAs("联盟少侠" + t.Format("20060102150405") + ".xlsx"); err != nil {
		println(err.Error())
	}
	// 等待用户关闭（输入）
	println("软件开源地址：https://github.com/user1121114685/tiku")
	println("执行完毕........请关闭窗口......")
	var exitScan string
	_, _ = fmt.Scan(&exitScan)

}
