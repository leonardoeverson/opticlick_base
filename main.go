package main

import (
	"context"
	"fmt"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
	"time"
)

func createMap(statusFile string) map[string]string {
	file, err := excelize.OpenFile(statusFile)

	if err != nil {
		log.Fatalln(err)
	}

	pedidoList := map[string]string{}

	rows, err := file.GetRows(file.GetSheetName(0))
	if err != nil {
		log.Fatalln(err)
	}

	for _, row := range rows {
		if len(row) == 4 && strings.Contains(strings.ToLower(row[3]), "cancelado") {
			pedidoList[row[0]] = row[3]
		}
	}

	return pedidoList
}

func readPedidos(base string) [][]string {
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

func createNewSheet() *excelize.File {
	return excelize.NewFile()
}

func writeCellValue(file *excelize.File, cell string, value string) {
	err := file.SetCellValue("Sheet1", cell, value)
	if err != nil {
		log.Fatalln(err)
	}
}

func generateFile(base string, status string) *excelize.File {
	pedidoList := createMap(status)

	rows := readPedidos(base)

	newSheet := createNewSheet()

	letters := []string{
		"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	}

	columns := []string{
		"PEDIDO",
		"DT_INCLUSAO",
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
			for idx, col := range columns {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[idx], 1), col)
			}

			continue
		}

		if _, ok := pedidoList[row[0]]; ok {
			cancelado = "SIM"
		}

		for col, value := range row {
			letterIdx := col

			if col > 3 && col < 14 {
				letterIdx = col + 1
			}

			if col == 14 {
				continue
			}

			if col == 3 {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", "D", index+1), cancelado)
				writeCellValue(newSheet, fmt.Sprintf("%v%v", "E", index+1), value)
				continue
			}

			if col == 10 {
				if value == "1" {
					corte = "COM CORTE"
				}

				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), corte)
				continue
			}

			if col >= 15 && col <= 17 && row[14] == "lab" && row[11] == row[15] {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), "")
				continue
			}

			if col >= 11 && col <= 13 && row[14] == "optica" {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), row[col+4])
				continue
			}

			if col >= 15 && col <= 17 && row[14] == "optica" {
				writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), "")
				continue
			}

			writeCellValue(newSheet, fmt.Sprintf("%v%v", letters[letterIdx], index+1), value)
		}
	}

	return newSheet
}

func writeFile(sheet *excelize.File, name string) error {
	err := sheet.SaveAs(name)

	if err != nil {
		return err
	}

	return nil
}

func main() {
	// cmd.Execute()
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

			sheet := generateFile(arquivoBase, arquivoStatus)
			fileName := fmt.Sprintf("%v.xlsx", time.Now().Format("2006_01_02_15_04_05"))

			if arquivoFinal != "" {
				fileName = arquivoFinal
			}

			err := writeFile(sheet, fileName)
			if err != nil {
				fmt.Println(fileName)
				log.Fatalln("Erro ao gerar a planilha final. Erro: ", err)
			}

			return nil
		},
	}

	if err := cmd.Run(context.Background(), os.Args); err != nil {
		log.Fatalln(err)
	}
}
