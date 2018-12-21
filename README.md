# Translation_Python
python连接百度、搜狗等翻译平台进行字段翻译代码

-------------------------------------------------------------------------------------------------------------------
说明：
	运行该脚本文件首先需安装python，建议使用python3.6

	安装外部库：openpyxl、json

	安装指令： 
		pip install openpyxl
		pip install json

-------------------------------------------------------------------------------------------------------------------
	<<< 表格自动翻译工具 >>>

    使用说明:

        1、将需要翻译的表格文件防止脚本同级目录下，文件格式为 xlsx
        2、windows用户可运行目录下的run.bat文件,需用户自主配置python环境
        3、输入文件名称，不需要输入文件格式
        4、输入需要翻译的表格单元格列名，例如： A，B
        5、输入翻译保存表格单元格
        6、直到运行完成即可(提示' 翻译完成 '字段信息)
    
    注意事项：
        1、如果需要百度翻译平台，则需要修改config.json文件中的appid和key(目前软件只支持百度翻译)
           appid和key的获取可通过注册百度开发者翻译平台开发者获取
           
    附：百度翻译平台  http://api.fanyi.baidu.com/api/trans/product/index

--------------------------------------------------------------------------------------------------------------------
作者： Awesome QQ：2281280195
