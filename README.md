
# Structurama for Go

Excel Reader for Go is a lightweight package designed to simplify the process of reading data from Excel files and mapping it to custom Go structs.

## Features

- **Easy Integration**: Quickly integrate Excel reading capabilities into your Go applications.
- **Custom Struct Mapping**: Define a custom Go struct to represent the structure of the data in your Excel file.

## Installation

```bash
go get -u github.com/PreciousNyasulu/structurama
```

## Quick Start

```go
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
```

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


## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

## License

This package is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.