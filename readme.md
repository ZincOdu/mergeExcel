一个小巧的excel合并工具

读取每个excel第一个sheet合并到一个新的excel文件中

可以按行合并也可以按sheet合并

界面：

![image](./interface.jpg)

打包：

```
pyinstaller -w main.py -i icons8-table-32.ico --name=mergeExcel --onefile
```