#### 页面级导出



性能取决于客户端电脑， 数据量取决于客户端内存。一般几十万的数据量是没有问题，导出时间在秒级别，一般数据少的话就是毫秒级



导出分为两步 

 	1.生成
		1.  字符串拼接
		2.  按照规范生成
		3.  使用ActiveX 控件（ie 系列） 
 	2.导出
		1.  使用base64 生成url地址，然后触发 url  (不支持ie 8)
		2.  使用FileSaver.js ，(不支持ie8)
		3.  使用flash (主要是支持ie8，使用比较麻烦，用户体验不好)
		4.  使用ActiveX 控件（ie 系列） 



第一步的问题就是一些样式的支持和表格合并的支持

第二步的问题就是浏览器的支持，ie9一下只有flash 和ActiveX 控件，并且巨难用。



#### demo 介绍



1. 文件夹：

   DataTables ： Datatables是一款jquery表格插件。它是一个高度灵活的工具，可以将任何HTML表格添加高级的交互功能。

   excel.plugin ：  jquery 的一个excel 导出的插件

   LayUI ： LayUI里的表格导出插件，功能比较全，缺点是不支持ie8，表格需要转成数据才能导出

   HtmlToExcel ：网上比较流行的字符串拼接的方式导出。 简单但是各种小问题需要自己捣鼓

   SheetJs ： 一个比较简单实用的表格导出插件，各种支持，pro 版的才支持样式和性能

   Sheetjs2 ： 同上，扒下来了如何支持ie8 的文件

   

2. 使用方式：

   1 `git clone `

   2 安装python3

   3 `python -m http.server 9000`