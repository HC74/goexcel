# goexcel

```bash
go get github.com/HC74/goexcel@latest
```
## <u>Quick start 快速开始</u>

### 读模式
```go
package main

import (
	"fmt"
	"github.com/HC74/goexcel"
)

type TestStruct struct {
	Name string `ex:"name:¬名称"`
	Age  int    `ex:"name:年龄"`
}

func main() {
	var datas []TestStruct
	err := goexcel.Load("test.xlsx", TestStruct{}).SheetName("Sheet1").Read(&datas)
	if err != nil {
		fmt.Print(err.Error())
	}
}
```
### 写模式
