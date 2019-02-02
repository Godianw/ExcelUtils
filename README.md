#ExcelUtils

---
***基于poi的解析和导出通用工具类***

* **Maven依赖**

*poi*：

    <dependency>
	    <groupId>org.apache.poi</groupId>
	    <artifactId>poi</artifactId>
		<version>4.0.1</version>
	</dependency> 
	<dependency>
		<groupId>org.apache.poi</groupId>
		<artifactId>poi-ooxml</artifactId>
		<version>4.0.1</version>
	</dependency>

* **用法**

*解析excel*：
![](https://i.imgur.com/EpqN8Dd.png)
<br>
*创建一个表格列名与模型字段名的映射*
<br><br>

![](https://i.imgur.com/XVHYo6A.png)
*读取excel表格内容到集合中*
<br><br>

*导出excel*：
![](https://i.imgur.com/aZWPDdP.png)
<br>
*创建一个模型字段名与表格列名的映射*
<br><br>

![](https://i.imgur.com/qgTJJ6a.png)
*将集合写入到excel文件中*