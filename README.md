# clswindow v2.2 
#### VB6操作外部程序窗口的类clsWindow2.2使用说明
详细使用说明文档：https://www.kancloud.cn/sysdzw/clswindow

clsWindow是VB6环境下使用的一个操作外部程序窗口的类，比如得到窗口句柄，得到窗口里某个文本框的内容。非常方便，使用它可以让您脱身于一堆api函数，功能强大使用简单！

这个类楼主很早就开始封装了，原本打算做成类似DOM对象那样，通过一堆getElmentByXXX等方法实现对桌面程序下各个窗口以及里面各个控件对象的自由访问，但是具体要做的工作太多，目前只实现了一部分，期待大家一起加入更新维护。

目前该类封装了绝大部分对windows窗口的常用操作，例如：获取窗口句柄，设置窗口为活动窗口，设置窗口内文本框内容，点击窗口内的某些按钮等。

这个类现在还在一直不断地扩充，功能已经很强大很广泛，使用它可以轻而易举地设置窗口标题栏文字，移动窗体等等。以前要实现这些操作常常需要一大堆api函数，现在只需要一点点代码就可以了，完全让您脱身于api函数的海洋。当然您如果想知道里面的方法实现原理的话可以看一看源代码。

#### 备注：这个框架已经单身8年了，主要组成为一个bas模块和一个类模块，很多人好奇为什么会多出一个bas模块，都弄到一个cls类模块里不行吗？我也想过，主要是里面有个回调得放到模块里面，如果有谁能够做到将这bas模块消灭，合并到一个类模块中请贡献代码，到时更新注释中可以写上你的大名哦。开源框架的建设，需要你我每一个人的添砖加瓦！(*^▽^*) 测试模块下载: https://pan.baidu.com/s/1ZYMFQJ3fHXZiH75KM-JjSw 提取码: 5fc3



## 使用范例：
### 1）关闭腾讯新闻窗口“腾讯网迷你版”。
```vb
Dim window As New clsWindow
If window.GetWindowHwndByTitle("腾讯网迷你版").hwnd > 0 Then
    window.CloseWindow  '关闭窗口
End If
```
以上是不是很简洁呢？

### 2）获取某个打开的记事本里面的内容。假设记事本标题为“测试.txt - 记事本”，通过SPY等工具查看得知记事本的文本框类名为：Edit，那么我们编写程序如下：
```vb
Dim window As New clsWindow
If window.GetWindowHwndByTitle("测试.txt - 记事本").hwnd > 0 Then
    MsgBox window.GetElementTextByClassName("Edit")
End If
```
这个看起来也很简单，方法自由还可以使用正则匹配，可以写成下面这样：
```vb
Dim window As New clsWindow
If window.GetWindowHwndByTitleRegExp("测试\.txt.*?").hwnd > 0 Then
    MsgBox window.GetElementTextByClassName("Edi", , True)'第三个参数表示是否使用正则，默认为false
End If
```
v1.9以上版本已经可以使用连写功能。window.GetWindowHwndByTitle("腾讯网迷你版").CloseWindow 这样写是不是很酷呢？
更多演示案例：

类成员以及各个使用方法如下：



### csdn博客链接：
http://blog.csdn.net/sysdzw/article/details/9083313

## 更新日志

```vb
'==============================================================================================
'名    称：windows窗体控制类v2.2
'描    述：一个操作windows窗口的类，可对窗口进行很多常用的操作(类名为clsWindow)
'使用范例：Dim window As New clsWindow
'          window.GetWindowByTitle("计算器").closeWindow
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份，共同维护
'发布日期：2013/06/01
'博    客：https://blog.csdn.net/sysdzw
'用户手册：https://www.kancloud.cn/sysdzw/clswindow/
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                           2012/12/03
'          V1.1 修正了几个正则相关的函数，调整了部分类结构                       2013/05/28
'          V1.2 增加属性Caption，可以获取或设置当前标题栏                        2013/05/29
'          V1.3 增加了方法Focus，可以激活当前窗口                               2013/06/01
'               增加了方法Left,Top,Width,Height,Move，处理窗口位置等
'          V1.4 增加了窗口位置调整的几个函数                                    2013/06/04
'               增加了得到应用程序路径的函数AppName
'               增加了得到应用程序启动参数的函数AppCommandLine
'          V1.5 增加了窗口最大最小化，隐藏显示正常的几个函数                     2013/06/06
'               增加了获取控件相关函数是否使用正则的参数UseRegExp默认F
'          V1.6 将Left，Top函数改为属性，可获得可设置                           2013/06/10
'          V1.7 增加函数：CloseApp 结束进程                                    2013/06/13
'               修正了部分跟正则匹配相关的函数
'               增加函数：GetElementTextByText
'               增加函数：GetElementHwndByText
'          V1.8 增加函数：GetWindowByClassName                                 2013/06/26
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByAppName
'               增加私有变量hWnd_
'               增加属性hWnd，可设置，单设置时候会检查，非法则设置为0
'               更新GetWindowByTitleEx函数，使之可以选择性支持正则
'               删除GetWindowByTitleRegExp函数，合并到上面函数
'               增加SetFocus函数，调用Focus实现，为了是兼容VB习惯
'               扩了ProcessID、AppPath、AppName、AppCommandLine三个函数，可带参数
'               网友wwb(wwbing@gmail.com)提供了一些函数和方法属性：
'                 CheckWindow, Load, WindowState, Visible, hDC, ZOrder
'                 AlphaBlend, Enabled, Refresh, TransparentColor
'               采纳wwb网友的部分意见，将句柄变量改为hWnd_，但是hWnd作为公共属性
'          V1.9 修正函数：GetMatchHwndFromWindow 正则表达式的错误               2013/08/07
'               修正函数：GetMatchHwndFromWindow 函数中的一些错误               2014/09/23
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByPID 根据PID取窗口句柄
'               增加函数：GetCaptionByHwnd 根据句柄取得标题
'               增加函数：SetTop设置窗体置顶，传入参数false则取消                2014/09/24
'               增加函数：Shake、FadeIn、FadeOut 抖动、淡入、淡出特效
'          V2.0 修正函数：GetWindowByPID 遍历窗体Win7下有一些问题               2015/09/29
'               修正函数：GetWindowByAppName 遍历窗体Win7下有一些问题
'               修正函数：GetWindowByAppNameEx 遍历窗体Win7下有一些问题
'          V2.1 修正函数：ClickPoint 增加位置模式参数相对和绝对，默认相对        2018/06/05
'               增加函数：SelectComboBoxIndex 根据指定的index选择下拉框中的项
'                         上述方法得到网友Chen8013的不少帮助，特此感谢
'               增加函数：GetWindowByHwnd 根据指定的句柄确定窗口                2018/07/22
'               增加函数：GetWindowByCursorPos 根据当前光标获取窗口（控件）
'               增加函数：GetWindowByPoint 根据指定的位置获取窗口（控件）
'               升级ClickPoint函数，支持点击前后分别延时，默认延时为0            2018/07/23
'          V2.2 修正正则：网友小凡反应了句柄和id存在负数的情况                   2020/01/08
'               优化属性：Caption(Get)，根据网友小凡的建议改成可获得文本框内容
'               增加方法：Wait 此方法原为clsWaitableTimer模块中，现集成进来      2020/01/09
'               增加方法：ClickCurrentPoint 点击当前点                          2020/01/10
'               增加方法：SetCursor(别名:SetPoint MoveCursor MoveCursorTo)
'               更新函数：将所有默认等待超时60秒的函数中默认等待时间都改为10秒
'               增加属性：Text、Value、Title（均为Caption别名）                 2020/01/12
'               优化代码：GetCaptionByHwnd采用原Caption(Get)代码，后者也做了调整
'               增加函数：GetCursorPosCurrent(别名：GetCursorPoint)得到当前坐标
'               优化函数：所有窗口获取的函数增加了是否过滤可见的参数              2020/01/16
'               增加函数：GetTextByHwnd（同GetCaptionByHwnd）
'               优化代码结构。将模块中能移过来的都移到类模块中了                  2020/01/19
'               增加函数：myIsWindowVisibled 判断窗体可见，长宽为0也认为不可见   2020/01/31
'               优化函数：GetTextByHwnd 网友小凡提供                            2020/02/03
'               增加函数：CommandLine（同AppCommandLine）                      2020/02/05
'               增加函数：MakeTransparent 设置窗口透明度                        2020/02/18
'               增加函数：MoveToCenter 移动窗口到屏幕中心
'               增加函数：IsTopmost 判断窗口是否为置顶                          2020/02/20
'               增加函数：GetWindowTextByHwnd 获得窗口标题，给窗口句柄专用       2020/02/28
'               修正函数：Focus 旧方法使用后会改变置顶窗口属性                   2020/03/02
'               增加函数：IsWin64 网友小凡提供                                  2020/03/12
'               修正函数：AppPath 网友小凡提供兼容64系统的方法
'               修正函数：AppCommandLine 网友小凡做了兼容64位处理及其他代码优化  2020/03/15
'==============================================================================================
```


![](https://img-blog.csdn.net/20180423135213794)

应用案例（均附源码）：
微便签：https://github.com/sysdzw/WeNote
窗口置顶小插件：https://github.com/sysdzw/SetWindowTop
