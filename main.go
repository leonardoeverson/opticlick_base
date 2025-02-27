package main

import (
	"context"
	"encoding/csv"
	"fmt"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"
)

func createMap(statusFile string) map[string]string {
	rows := readFiles(statusFile)
	pedidoList := map[string]string{}

	for _, row := range rows {
		if len(row) >= 4 && strings.Contains(strings.ToLower(row[3]), "cancelado") {
			if len(row) == 6 {
				log.Printf("Pedido %s cancelado em %s", row[0], row[5])
			} else {
				log.Printf("Pedido %s cancelado. Incluído em %s", row[0], row[4])
			}

			pedidoList[row[0]] = row[3]
		}
	}

	return pedidoList
}

func readFiles(base string) [][]string {
	if filepath.Ext(base) == ".xlsx" {
		file, err := excelize.OpenFile(base)

		if err != nil {
			log.Fatalln(err)
		}

		rows, err := file.GetRows(file.GetSheetName(0))
		if err != nil {
			log.Fatalln(err)
		}

		return rows
	}

	if filepath.Ext(base) == ".csv" {
		file, err := os.Open(base)
		if err != nil {
			log.Fatalln(err)
		}

		defer func(file *os.File) {
			err := file.Close()
			if err != nil {
				log.Fatalln(err)
			}
		}(file)

		reader := csv.NewReader(file)
		reader.Comma = ';'

		rows, err := reader.ReadAll()
		if err != nil {
			log.Fatalln(err)
		}

		return rows
	}

	return [][]string{}
}

func createNewSheet() *excelize.File {
	return excelize.NewFile()
}

func getCsvWriter(fileName string) (*csv.Writer, *os.File) {
	file, err := os.Create(fileName)
	if err != nil {
		log.Fatalln(err)
	}

	return csv.NewWriter(file), file
}

func writeDataCsvFile(w *csv.Writer, row []string) {
	err := w.Write(row)
	if err != nil {
		log.Fatalln(err)
	}
}

func writeCellValue(file *excelize.File, cell string, value string) {
	err := file.SetCellValue("Sheet1", cell, value)
	if err != nil {
		log.Fatalln(err)
	}
}

func getTipo(base string, status string) string {
	if filepath.Ext(base) == ".xlsx" || filepath.Ext(status) == ".xlsx" {
		return "xlsx"
	}

	return "csv"
}

func generateFile(base string, status string, fileName string) (*excelize.File, *csv.Writer) {
	pedidoList := createMap(status)

	rows := readFiles(base)

	tipo := getTipo(base, status)
	var newSheet *excelize.File
	var newCsv *csv.Writer

	if tipo == "csv" {
		csvWriter, file := getCsvWriter(fileName)
		newCsv = csvWriter
		defer func(file *os.File) {
			err := file.Close()
			if err != nil {
				log.Fatalln(err)
			}
		}(file)
		defer newCsv.Flush()
	} else {
		newSheet = createNewSheet()
	}

	letters := []string{
		"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	}

	columns := []string{
		"PEDIDO",
		"DT_INCLUSAO",
		"DT_CALCULO",
		"TIPO_DOCUMENTO",
		"PEDIDO DE PRODUÇÃO CANCELADO?",
		"RPF",
		"NOME_LENTE",
		"MARCA",
		"MATERIAL",
		"AR",
		"HC",
		"COLORACAO",
		"CORTE",
		"COD_LAB_PRODUTOR",
		"CNPJ_LAB_PRODUTOR",
		"RAZAO_SOCIAL_LAB_PRODUTOR",
		"COD_LAB_INTERMEDIÁRIO",
		"CNPJ_LAB_INTERMEDIÁRIO",
		"RAZAO_SOCIAL_LAB_INTERMEDIÁRIO",
		"COD_OPTICA",
		"CNPJ_OPTICA",
		"RAZAO_SOCIAL_OPTICA",
	}

	for index, row := range rows {
		cancelado := "NÃO"
		corte := "SEM CORTE"

		if index == 0 {
			var record []string

			for idx, col := range columns {
				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[idx], 1), col)
				} else {
					record = append(record, col)
				}
			}

			if len(record) > 0 && tipo == "csv" {
				writeDataCsvFile(newCsv, record)
			}

			continue
		}

		if _, ok := pedidoList[row[0]]; ok {
			cancelado = "SIM"
		}

		var record []string

		for col, value := range row {
			letterIdx := col

			if col > 3 && col < 15 {
				letterIdx = col + 1
			}

			if col == 15 {
				continue
			}

			if col == 4 {
				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", "E", index+1), cancelado)
					writeCellValue(newSheet, fmt.Sprintf("%v%v", "F", index+1), value)
				} else {
					record = append(record, cancelado)
					record = append(record, value)
				}

				continue
			}

			if col == 11 {
				if value == "1" {
					corte = "COM CORTE"
				}

				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), corte)
				} else {
					record = append(record, corte)
				}

				continue
			}

			if col >= 16 && col <= 18 && row[15] == "lab" && row[12] == row[16] {
				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), "")
				} else {
					record = append(record, "")
				}

				continue
			}

			if col >= 12 && col <= 14 && row[15] == "optica" {
				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), row[col+4])
				} else {
					record = append(record, row[col+4])
				}

				continue
			}

			if col >= 16 && col <= 18 && row[15] == "optica" {
				if tipo == "xlsx" {
					writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), "")
				} else {
					record = append(record, "")
				}

				continue
			}

			if tipo == "xlsx" {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), value)
			} else {
				record = append(record, value)
			}
		}

		if len(record) > 0 && tipo == "csv" {
			writeDataCsvFile(newCsv, record)
		}
	}

	if tipo == "xlsx" {
		return newSheet, nil
	}

	return nil, newCsv
}

func writeFile(sheet *excelize.File, name string) error {
	err := sheet.SaveAs(name)

	if err != nil {
		return err
	}

	return nil
}

func getFileName(fileName string) string {
	timeStamp := time.Now().Format("2006_01_02_15_04_05")

	if filepath.Ext(fileName) == ".csv" {
		return fmt.Sprintf("%v.csv", timeStamp)
	}

	return fmt.Sprintf("%v.xlsx", timeStamp)
}

func main() {
	var arquivoBase string
	var arquivoStatus string
	var arquivoFinal string

	cmd := &cli.Command{
		Name:        "Construtor",
		Description: "Construtor de relatório",
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "base",
				Usage:       "Arquivo de base de dados",
				Aliases:     []string{"b"},
				Destination: &arquivoBase,
			},
			&cli.StringFlag{
				Name:        "status",
				Usage:       "Arquivo de status de pedido",
				Aliases:     []string{"s"},
				Destination: &arquivoStatus,
			},
			&cli.StringFlag{
				Name:        "final",
				Aliases:     []string{"f"},
				Usage:       "Nome do arquivo final",
				Destination: &arquivoFinal,
			},
		},
		Action: func(ctx context.Context, cmd *cli.Command) error {
			if cmd.String("base") == "" || cmd.String("status") == "" {
				fmt.Println("Defina os arquivos de base e status")
				return nil
			}

			if _, err := os.Stat(cmd.String("base")); err != nil {
				fmt.Println("Arquivo de base não encontrado")
				return nil
			}

			if _, err := os.Stat(cmd.String("base")); err != nil {
				fmt.Println("Arquivo de base não encontrado")
				return nil
			}

			if _, err := os.Stat(cmd.String("status")); err != nil {
				fmt.Println("Arquivo de status não encontrado")
				return nil
			}

			fileName := getFileName(arquivoBase)

			sheet, _ := generateFile(arquivoBase, arquivoStatus, fileName)

			if arquivoFinal != "" {
				fileName = arquivoFinal
			}

			if sheet != nil {
				err := writeFile(sheet, fileName)
				if err != nil {
					fmt.Println(fileName)
					log.Fatalln("Erro ao gerar a planilha final. Erro: ", err)
				}
			}

			return nil
		},
	}

	if err := cmd.Run(context.Background(), os.Args); err != nil {
		log.Fatalln(err)
	}
}
