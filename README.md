# Tianyancha_guanxi
#### 运行环境及下载地址
- python 地址：https://www.python.org/downloads/
- selenium `pip install selenium`
- Chromedriver 需要查看Chrome版本下载对应Chromedriver
  具体步骤：https://blog.csdn.net/huixiaodezuotian/article/details/120225999
  Chromedriver下载地址：http://chromedriver.storage.googleapis.com/index.html
1、chrome://version/ 查看chrome的版本。
2、可以直接从这个官方的网址下载    https://googlechromelabs.github.io/chrome-for-testing/   ，其中稳定版本是可用的。
3、在进行2之前可以先进行chrome的卸载，控制面板，右键卸载。然后直接在2当中安装。记得将chrome的driver放到python路径下。
 #### 使用说明
本程序通过selenium模拟搜索并批量下载天眼查公司关系图，效果如下：
1. 用户需要先选择其需要打开的Excel文件和默认的下载路径
2. 程序会自己打开 https://www.tianyancha.com/relation 网站，但需要用户自行登录，预设时间为30秒
3. 程序会自动读取用户选择的Excel文件中“Sheet1”中名为“公司列表”的列，对所有对象进行遍历，查询其两两之间的关系，并下载其关系图。如果没有关系，程序会对页面进行截屏。
4. 并将其整理放入一份名为“公司关联关系图”的word文档中，其位置为程序所处的文件夹
