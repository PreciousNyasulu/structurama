package main

import (
	"fmt"
	structurama "github.com/PreciousNyasulu/structurama/reader"
)

func main() {
	readExcelFileStructure()
}

func readExcelFileStructure() {
	data, err := structurama.ReadFile("./example.xlsx", Person{}, true)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(data)

	personData, ok := data.([]Person)
	if ok {
		for _, item := range personData {
			fmt.Println(item.Id, item.FirstName, item.LastName, item.Age)
		}
	}
}

type Person struct {
	Id        string `xlsx:"0"`
	FirstName string `xlsx:"1"`
	LastName  string `xlsx:"2"`
	Age       string `xlsx:"3"`
}
