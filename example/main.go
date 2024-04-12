package main

import (
	"fmt"
	"github.com/PreciousNyasulu/structurama/reader"
)

type Person struct {
	Id        string `xlsx:"0"`
	FirstName string `xlsx:"1"`
	LastName  string `xlsx:"2"`
	Age       string `xlsx:"3"`
}

func main() {
	data, err := structurama.ReadFileDefault("./example.xlsx", Person{}, true)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(data)		

	personData, ok := data.([]Person)
	if ok {
		for _ , item := range personData  {
			fmt.Println(item.Id, item.FirstName, item.LastName, item.Age)
		}
	}
}