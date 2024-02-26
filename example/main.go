package main

import (
	"fmt"
	"github.com/PreciousNyasulu/structurama/reader"
)

func main() {
	data, err := structurama.ReadFile("./example.xlsx", Person{}, true)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(data)
}

type Person struct {
	Id        string `xlsx:"0"`
	FirstName string `xlsx:"1"`
	LastName  string `xlsx:"2"`
	Age       string `xlsx:"3"`
}
