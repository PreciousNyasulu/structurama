
# Structurama

Structurama is a lightweight package designed to simplify the process of reading data from Excel files and mapping it to custom Go structs.

## Features

- **Easy Integration**: Quickly integrate Excel reading capabilities into your Go applications.
- **Custom Struct Mapping**: Define a custom Go struct to represent the structure of the data in your Excel file.

## Installation

```bash
go get -u github.com/PreciousNyasulu/structurama
```

## Quick Start

## Custom Struct

Define a custom struct to represent the structure of the data you expect from the Excel file.

```go
type Person struct {
    Id        string `xlsx:"0"`
    FirstName string `xlsx:"1"`
    LastName  string `xlsx:"2"`
    Age       string `xlsx:"3"`
}
```

How to use:

```go
package main

import (
    "fmt"
    "github.com/PreciousNyasulu/structurama/reader"
)

func main() {
    data, err := structurama.ReadFileDefault("./example.xlsx", Person{}, true)
    if err != nil {
        fmt.Println(err)
        return
    }
    fmt.Println(data)
}
```

Here is how you can read file by Sheetname: 

```go
func main() {
    data, err := structurama.ReadFileBySheetName("./example.xlsx", Person{}, true,"Sheet1")
    if err != nil {
        fmt.Println(err)
        return
    }
    fmt.Println(data)
}
```

## Casting to the Struct

```go
func main() {
    data, err := structurama.ReadFileDefault("./example.xlsx", Person{}, true)
    if err != nil {
        fmt.Println(err)
        return
    }
    fmt.Println(data)

    peopleData, ok := data.([]Person)
    if ok {
        for _ , item := range peopleData  {
            fmt.Println(item.Id, item.FirstName, item.LastName, item.Age)
        }
    }                         
}
```

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## Inspiration 
[tealeg/xlsx](github.com/tealeg/xlsx) - Go library for reading and writing XLSX files.

## License

This package is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
