# 🎈🎗🎁🎄🎆VBA-VBE-VBE_AutoComplete
***VBE_AutoComplete**是一款VBE(Visual Basic Editor，也叫vba代码编辑器)插件  插件采用宿主无关性设计，完美兼容 Excel、Word、PPT 、AutoCAD 、SolidWorks、CorelDraw ……等所有内置 VBE 的应用程序  插件不影响宿主(如Excel)的运行效率，只有当你打开VBE界面，插件才开始加载

# VBE _AutoComplete使用说明

## 1、简介

名称 | 详情  
---|---  
插件名称 | VBE_AutoComplete  
测试系统 | win7，10，11  
适配VBE | 任何VBE均可使用完整功能，  
如Excel,Word,PPT,cad,solidworks,cdr...  
  
VBE_AutoComplete是一款VBE(Visual Basic Editor，也叫vba代码编辑器)插件

插件采用宿主无关性设计，完美兼容 Excel、Word、PPT 、AutoCAD 、SolidWorks、CorelDraw ……等所有内置 VBE 的应用程序

插件**不影响宿主(如Excel)的运行效率** ，只有当你打开VBE界面，插件才开始加载

[视频使用教程(评论区有Q群，群内有最新文件)](https://www.bilibili.com/video/BV1RRxPzHEeP)

🎥 相关提醒

  * **本文档** 包含插件的**最新功能和说明**
  * 由于视频制作周期较长，可能未能及时覆盖最新功能更新，请知悉



### 主要功能

 _/说明/_ 以下快捷键支持自定义，具体详见`设置`章节

功能 | 快捷键  
---|---  
代码buffer补全 |   
代码片段补全 |   
后缀补全 |   
符号补全 |   
片段管理 |   
代码库管理 |   
组件过滤 |   
CDR高版本中文注释修复 |   
往下移动行(支持多行) | `⎈ Ctrl` `J`  
往上移动行(支持多行) | `⎈ Ctrl` `K`  
复制行(支持多行) | `⎈ Ctrl` `Y`  
上方插入行 | `⇧ Shift` `⏎ Enter`  
下方插入行 | `⎈ Ctrl` `⏎ Enter`  
注释 | `F4`或`⎈ Ctrl` `/`  
删除当前行或选中行(支持多行) | `⎈ Ctrl` `D`  
拼接行 | `⎈ Ctrl` `⇧ Shift` `J`  
删除到行尾 | `⎈ Ctrl` `⇧ Shift` `⏎ Enter`  
删除光标前的词(VBE原生功能) | `⎈ Ctrl` `Backspace`  
删除光标后的词 | `⎈ Ctrl` `⇧ Shift` `Backspace`  
格式化代码并保存 | `⎈ Ctrl` `S`  
跳到上个位置 | `⎈ Ctrl` `O`  
跳到定义 | `⎈ Alt` `Left 🖱️`  
清空【立即窗口】 | `⇧ Shift` ``` (反引号键)  
清空【立即窗口】并运行到当前光标处 | `⎈ Ctrl` ``` (反引号键)  
  
## 2、安装及更新

### 安装前

添加信任 | 部分Win7用户需安装运行环境  
---|---  
_软件使用了加壳，杀毒软件会误报_  
 _本人保证软件绝对绿色无毒_  
 _如有误杀，请将插件加入信任_  
  
[WIN10 杀毒 添加信任](https://search.bilibili.com/all?vt=61735707&keyword=WIN10+%E6%9D%80%E6%AF%92+%E6%B7%BB%E5%8A%A0%E4%BF%A1%E4%BB%BB)  
[360 添加信任](https://search.bilibili.com/all?vt=61735707&keyword=360+%E6%B7%BB%E5%8A%A0%E4%BF%A1%E4%BB%BB)  
[腾讯电脑管家 添加信任](https://search.bilibili.com/all?vt=61735707&keyword=%E8%85%BE%E8%AE%AF%E7%94%B5%E8%84%91%E7%AE%A1%E5%AE%B6+%E6%B7%BB%E5%8A%A0%E4%BF%A1%E4%BB%BB)  
[火绒 添加信任](https://search.bilibili.com/all?vt=61735707&keyword=%E7%81%AB%E7%BB%92+%E6%B7%BB%E5%8A%A0%E4%BF%A1%E4%BB%BB) | win7早期用户`运行环境`可能缺失  
  
请群文件下载`.net48RunTime.zip`  
  
根据压缩包内`Readme.txt`进行安装`运行环境`  
  
WPS用户需关闭沙箱  

![download](https://github.com/user-attachments/assets/6f3f536f-f81f-4018-a0aa-7d25912033dd)

---  

  
WIN11用户需将压缩包【解除锁定】后再解压  
<img width="1473" height="601" alt="download" src="https://github.com/user-attachments/assets/8f8ee340-2892-490b-ad9c-52004f531c4c" />

---  
  
### 安装

  1. 解压压缩包，将解压后的文件夹移到你安装软件目录(任意路径，根据个人喜好而定)
  2. 双击`.一键安装.exe`进行安装(同理`.一键卸载.exe`可进行卸载)
  3. **/必须安装/** 安装自动弹出的【Maple Mono SC NF Light】字体



  * 安装完成后，此文件夹及里面的文件不可二次移动，如需移动请先卸载插件
  * 以下两图都符合，即为安装成功
![download](https://github.com/user-attachments/assets/0d78a052-7e82-44f0-858a-8c73b60c0dd3)




### 更新

  * _/说明/_
  * 【旧版本插件】简称【旧插件】
  * 【新版本插件】简称【新插件】
  *   * 旧插件文件夹内点`一键卸载`进行卸载
  * 新插件文件夹内点`一键安装`进行安装
  * 复制旧插件件夹内的`user_config`文件夹到新插件文件夹内即完成更新
  * 待打开vbe插件加载完毕确认配置无误后，旧插件文件夹即可删去
  * _/补充/_ `user_config`文件夹是用户文件夹，里面存有你所有的配置，插件运行前会判断，不存在会自动生成



## 3、代码补全
<img width="1934" height="993" alt="download" src="https://github.com/user-attachments/assets/f8800eaf-6069-4713-809b-65d91fb938b2" />


名称 | 内容  
---|---  
触发条件 | `中文输入法的英文状态`或`英文键盘`  
补全范围 | 声明区词语+模块名+过程名+当前过程内词语+工作表名  
候选词排序 | 程序会根据匹配级别自动调整顺序  
如上图右部分【候选排序图】  
A区和B区的区别为A区为连续匹配，B区为不连续匹配  
  
按匹配顺序依次为：  
/A区/  
1、行首匹配  
2、非行首匹配  
3、行首中文匹配  
4、非行首中文匹配  
  
图上仅列出了文本的排序，代码片段排序同理  
不过同级别下，代码片段总是排文本前面  
权重 | 本插件中权重功能比较有限，在【候选词排序同一级别】中，权重越大越靠前  
范围0-100，普通词默认`权重`10，片段默认`权重`20  
快捷操作 | 选择补全项：按`↑ Up`或`↓ Down`  
  
上屏：Tab或鼠标左键  
  
编辑片段：鼠标右键或Ctrl+i  
  
快捷添加片段：选中代码，按Alt+A快捷添加  
  
取消补全窗口：ESC  
  
上屏后片段中如带有占位符`$1...$n`，按TAB可跳到下一个占位符(如补全框存在，需先按esc取消补全框)，支持多个同序占位符：  
比如2个$1，填写完第一个$1时，按TAB跳转，全部(2个)$1均替换为新内容  
  
  * 右键菜单可进入`代码片段窗口`进行片段管理
  * Q群群文件也提供片段导入  



## 4、后缀补全

### ..se

<img width="2546" height="774" alt="download" src="https://github.com/user-attachments/assets/8f4a82d9-1b10-4902-bab8-0dec57199123" />

..se，se即set缩写

上图第一行，光标处输入`..se`后，变成第二行的效果，a  
即前端添加`set =`，且光标置中

### ..as
<img width="3432" height="1258" alt="download" src="https://github.com/user-attachments/assets/404715cc-9b5b-49fc-9947-358140dc4d64" />

上图底部行，光标处输入`..as`后，在其上行会新添一行`Dim`行，  
并弹框补全

### ..co

情况1  
---  
..co，co即copy缩写  
  <img width="2842" height="1306" alt="download" src="https://github.com/user-attachments/assets/86e2edc4-2c6b-4271-aab0-cc2d64de7c2a" />

上图第一行，/当前行/为`Dim`行时，光标处输入`..co`后，  
下下行插入带变量的行  
  
情况2  
---  
<img width="2724" height="1010" alt="download" src="https://github.com/user-attachments/assets/ff97cde8-df2b-4aa2-888f-e2d4e8db474f" />

上图第一行，/当前行/为`非Dim`行时,光标处输入`..co`后，  
下行插入带变量的行  
  
### ..va

<img width="1778" height="810" alt="download" src="https://github.com/user-attachments/assets/aa878497-d76b-406b-b36b-0119c17cb3ef" />

..va，va即variable缩写

上图第一行，光标处输入`..va`后，行首添加`=`，且光标移到前面

## 5、符号补全

  * 支持成对符号快捷输入，即输入左侧，自动补全右侧并光标居中，当前支持`""`,`()`,`[]`



## 6、模块过滤窗口
<img width="567" height="723" alt="download" src="https://github.com/user-attachments/assets/bb855bde-c1db-4891-956e-a142289ef746" />


  * 显示当前项目的模块结构 
  * 当模块发生变化时自动更新，如增删模块 
  * 插件安装后，右键【显/隐模块过滤窗体】，可调出本窗体，可再拖动嵌入vbe窗体中 
  * 搜索支持拼音首字母，即`mk`能匹配`模块`
  * 搜索框打完字后，可直接按`↑ Up`和`↓ Down`进行选择，`⏎ Enter`确认，也可鼠标直接双击子项跳转 
  * `⎈ Ctrl` `⇧ Shift` `/`可从其他窗口快速跳转到搜索框 
  * 设置里可设置【默认开启模块过滤窗口】选项 



## 7、大纲窗口

<img width="567" height="520" alt="download" src="https://github.com/user-attachments/assets/5e55891d-7fcc-4793-a4bc-874c2f261e4f" />

  * 显示当前模块的模块变量及过程(sub、func) 
  * 当模块代码发生变化时自动更新 
  * 插件安装后，右键【显/隐大纲窗口】，可调出本窗体，可再拖动嵌入vbe窗体中 
  * 搜索支持拼音首字母，即`mk`能匹配`模块`
  * 鼠标直接单击对应项跳转 
  * 设置里可设置【默认开启大纲窗口】选项 



## 8、代码片段窗口

### 进入

从主窗口右键进入本窗体  

### 界面说明

<img width="1401" height="812" alt="download" src="https://github.com/user-attachments/assets/f3787da6-3a27-417e-94b3-bca71e8c50f2" />

> 代码片段用于收录高频使用的代码块，通过快捷输入即可快速插入

**1、筛选功能区**

  * 清空全部筛选：点击以清空右侧4个筛选项
  * 触发词：输入内容，以筛选Trigger列
  * 权 重：输入内容，以筛选Weight列
  * 进程名：输入内容，以筛选ProcNm列
  * 内 容：输入内容，以筛选Content列



**2、代码片段**

  * 此处列出了全部的代码片段或筛选后的代码片段



**3、预览窗**

  * 当前触发词：显示左侧选中片段对应的触发词
  * 当前权重：范围0-100，普通词默认`权重`10，片段默认`权重`20，同`动态候选级别下`权重越大越靠前
  * 当前进程名：显示左侧选中片段对应的`进程名`，如`进程名`为空，即为`通用片段`，任何vbe界面下都可用，否则反之，如进程名若为`EXCEL`即只能在Excel下才可用，可直接下拉框中选取
  * 当前内容：显示左侧选中片段对应的内容
  * 预览窗内容更改后，需点击`保存`按钮才能生效



**4、按钮区**

  * 导入：导入外部csv。可先操作【导出】后，获得【Snip_Export.csv】文件，然后修改这文件后再导入即可。如导入发现乱码，请用记事本文件打开csv另存为`带BOM的UTF8格式`再次导入即可。
  * 导出：导出csv文件
  * 新建：新建一条空白项，快捷键a(在`2、代码片段`上生效)
  * 复制：新建一条重复项，快捷键c(在`2、代码片段`上生效)
  * 删除：删除当前项，快捷键d、del(在`2、代码片段`上生效)
  * 排版：排版内容格式
  * 保存：保存当前项全部内容，快捷键`⎈ Ctrl` `S`(在`2、代码片段`上生效)



## 9、代码库窗口

### 进入


从主窗口右键进入本窗体  

### 界面说明

<img width="1998" height="1160" alt="download" src="https://github.com/user-attachments/assets/250f76bf-1454-4b6f-86bb-6722fbb18426" />

> 代码笔记用于记录和归档不常使用的代码，以便将来查阅，这也是和代码片段最直接的区别

**1、搜索区**

  * 输入内容，同时搜索笔记的内容和标题
  * 点击右边RE图标，可开启正则搜索



**2、搜索标签**

  * 点击搜索标签框，弹出下拉标签多选列表，可钩选多个标签，此时将只显示这些标签下的笔记
  * 在上面基础上，继续在`1、搜索区`输入内容进行搜索
  * 点击右边扫帚图标，可清空搜索标签



**3、代码笔记**

  * 显示全部的代码笔记或过滤后的代码笔记
  * 点击笔记，`4、预览窗`将显示对应内容



**4、预览窗**

  * 当前标题：显示左侧选中笔记对应的标题
  * 当前文件标签：显示左侧选中笔记对应的标题，点击可从弹出下拉标签多选列表中选择标签，也可填入自定义标签，多个自定义标签请用`空格`分隔， 点击右边扫帚图标，可清空当前笔记标签
  * 当前内容：显示左侧选中笔记对应的内容
  * 预览窗内容更改后，需点击`保存`按钮才能生效



**5、按钮区**

  * 导入：导入外部csv。可先操作【导出】后，获得【CodeLibrary_Export.csv】文件，然后修改这文件后再导入即可。如导入发现乱码，请用记事本打开csv另存为`带BOM的UTF8格式`再次导入即可
  * 导出：导出csv文件
  * 新建：新建一条空白项，快捷键`A`(在`3、代码笔记`上生效)
  * 复制：复制`4、预览窗`的内容到剪贴板
  * 删除：删除当前项，快捷键`D`、`⌫ Delete`(在`3、代码笔记`上生效)
  * 排版：排版内容格式
  * 写回：点击vbe空行处，然后点此按钮将内容写回到vbe，如选中文本，则只写回选中部分
  * 保存：保存当前项全部内容，快捷键`⎈ Ctrl` `S`(在`3、代码笔记`上生效)



## 10、工具栏


### 显示

<img width="515" height="1127" alt="download" src="https://github.com/user-attachments/assets/c04a9f77-0d59-42b7-b5d9-976493147f46" />

从主窗口右键设置显示  
<img width="308" height="250" alt="download" src="https://github.com/user-attachments/assets/07232846-543c-4d64-9c24-ca08ebd65b4c" />

可设置为随插件启动  
<img width="1045" height="341" alt="download" src="https://github.com/user-attachments/assets/906c52a8-f86a-428a-93dc-de927f3a3e95" />

### 修改

  * 工具栏上任意一个按钮右键，即可跳到设置窗体
  * 支持对图标执行增加、删除、调序三种操作
  * 操作过程中，工具栏实时更新  

<img width="1597" height="906" alt="download" src="https://github.com/user-attachments/assets/d5ef0a5a-6afe-4274-bc3b-28bfc953dea1" />


> 增加或删除

  * 窗口左区域，命令钩选后，点击【添加到工具栏(右箭头按钮)】，即可添加到窗口右区域  
反之操作，从右区域到左区域，即可删除
  * 可以一次钩选多个命令进行增加或删除



> 调序

  * 长按调序按钮拖到适当位置后松开鼠标即可完成调序
  * 调序只支持右窗口



## 11、插件后台
<img width="319" height="175" alt="download" src="https://github.com/user-attachments/assets/c1a39321-5faa-4e36-aab4-9a53f2448a87" />

<img width="333" height="291" alt="download" src="https://github.com/user-attachments/assets/5212b456-cfa3-4377-979b-581899b477f3" />

  * _/左图/_
  * 插件安装完成后，任务栏会出现左图所示图标
  * 鼠标移至图标上，会显示对应的插件信息
  *   * _/右图/_
  * 右键图标可弹出功能菜单
  * 打开插件目录 W：打开插件所在的文件夹，快捷键为W
  * 打开代码片段 F：打开代码片段窗口，快捷键为F
  * 代码代码库 E：打开代码库窗口，快捷键为E
  * 设置 S：打开设置窗口，快捷键为S
  * 重启 R：重启全部窗体和后台程序，快捷键为R
  * 重启键盘钩子 V：如发现按钮无反应，可执行此操作恢复键盘钩子，快捷键为V
  * 退出程序 X：退出程序，快捷键为X



## 12、设置

### 进入

从主窗口右键进入设置  

### 设置页面

#### 【常规】页面

<img width="1318" height="1058" alt="download" src="https://github.com/user-attachments/assets/90a9ca9a-4a13-4a85-96cd-4695dd74a497" />

> 说明

  * 当`值列`是数值时，`状态列`为常开状态，用户无法关闭
  * 当`值列`为空白不可填时，用户可自由开关`状态列`
  * `值列`修改后，焦点离开后设置方能生效



> 内容

功能 | 说明  
---|---  
默认开启模块过滤窗口 | 插件加载后自动开启模块过滤窗口  
默认开启工具栏窗口 | 插件加载后自动开启工具栏窗口  
中括号补全 | 开启或关闭[]补全  
双引号补全 | 开启或关闭""补全  
括号补全 | 开启或关闭()补全  
注释窗体字体大小 | 详见【CDR优化】章节  
注释窗体宽度 | 详见【CDR优化】章节  
注释窗体高度 | 详见【CDR优化】章节  
注释后自动移到下一行 | 使用插件，注释代码后，自动移到下一行，对VBE原生的注释命令无效  
添加选中至片段时自动格式化 | 选中代码按ALT+A添加到代码片段窗口时，自动格式化  
接替VBE原生CTRL_C | 解决原生VBE，在英文键盘下(不是中文输入法的英文状态)，复制的中文粘贴到其他软件中乱码的问题  
接替VBE原生CTRL_V | 解决在部分软件中复制的中文粘贴回VBE会出现乱码的问题  
  
#### 【快捷键】页面

![download](https://github.com/user-attachments/assets/b87d97ad-f5e8-4120-acbc-82884641a7c2)

> 操作说明

  * 修改按键：先按`⟵ Backspace`删除已有按键，再按下待设置的按键
  * 新增按键：直接按下待设置的按键
  * 删除按键：按下`⟵ Backspace`即可删除已设置的按键，多次按下可删除多个按键



#### 【工具栏】页面

  * 请查看【工具栏】章节



## 13、CDR优化

### 介绍

  * 解决`高版本CDR不能中文注释`
  * 解决`高版本CDR属性编辑窗体不能输入中文`
  * 解决`高版本CDR搜索框不能输入中文`



### 操作

  * 在代码主窗口/属性编辑窗口/搜索框，按`ALT+/`开启`注释窗口`
  * <img width="655" height="297" alt="download" src="https://github.com/user-attachments/assets/d2ad7db6-a00c-44d7-94ad-0a222143b51b" />
  * 输入对应内容后，`回车`上屏`原文内容`，或`CTRL+回车`上屏`' `+`原文内容`
  * 注释窗体开启期间，如焦点丢失，窗体将自动关闭



### 其他设置

  * 【常规】页面，提供了注释窗体自定义设置，如下图
  * <img width="1308" height="1058" alt="download" src="https://github.com/user-attachments/assets/6a377147-a44e-408a-bfc2-7130b50cc578" />

  * 【快捷键】页面，提供了注释窗体相关快捷键设置  
  * <img width="1308" height="1058" alt="download" src="https://github.com/user-attachments/assets/aa6ee8f3-798a-49f7-bc00-df00ebe6edf1" />



## 14、更新日志
<img width="2093" height="847" alt="01" src="https://github.com/user-attachments/assets/e08d3efe-c44c-43ab-88a3-e47c77a64344" />

