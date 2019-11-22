## MyExcel2Json

### 项目用途
解决excel表格转化为json数据的问题，能够通过定义表数据的类型以及字段名字

### 使用规则
传递进入的参数：
参数1.需要存储的json文件位置，
参数2.需要转化的excel文件路径

### excel文件规则
按sample规则配表，第一行作为json化的key，第二行表示数据类型，第三行作为描述，第四行为中文解释

支持的数据类型有：
```
	int32:      golang中int类型
	int64:      golang中int64类型
	string:     golang中string类型
	float32:    golang中float32类型
	float64:    golang中float64类型
	mapInterface:   string，字典格式
	percentage: 自定义百分号类型格式为  80% 50.23%等均可
	sheetInfo:      实例：SheetName:Sample|KeyID:ID|FindID:102 ， 使用SheetName:【本文件sheet名字】  KeyID:【第一行keyID】， FindID:【该sheet下的id】
    []float32:  golang中[]float32类型
	[]int32:    golang中[]int类型
    []int64:    golang中[]int64类型
    []float64:  golang中[]float64类型
	[]percentage:自定义百分号类型格式List  80%|50.23%, 用"|"作为分隔符
	[]string:       []string
	[]sheetInfo:    会将该sheet下的所有FindID值加入结果

```

### 一般规则
```
sheet名字中"@"开头为过滤符，不会处理
sheet中KeyID中"@"开头不作处理
数字类型的List数据一般为","作为分隔符
带有文字类型的数据一般使用"|"作为分隔符，如：[]string,[]percentage

```
### todo
```
目前实现的是代码中自定义一个结构体，能够读入至结构体中。
todo1：将读取的excel数据能够传入自定义的struct中，目前可以重新读取json文件在读入自定义结构体中
todo2：完成desc部分，能够通过desc获取更多的信息，进行灵活配置List的区分以及map的读取
todo3：通过一定规则区分json表格为list格式还是map格式，目前会自动判断如果表格中只有一个数据则自动将数据转化为map存储
todo4：在一个表格数据中的数据进行判断是否为sheetinfo类型，并进行递归查找
todo5：对一个文件夹进行读取转换，自动遍历文件夹中的excel文件进行读取转化
```