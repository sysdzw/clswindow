# clswindow
vb6自动化操作窗体程序，可以让你在vb中自由控制其他程序

当前最新版本v2.1


'==============================================================================================
'名    称：windows窗体控制类v2.1
'描    述：一个操作windows窗口的类，可对窗口进行很多常用的操作(类名为clsWindow)
'使用范例：Dim window As New clsWindow
'          window.GetWindowByTitle("计算器").closeWindow
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份，共同维护
'发布日期：2013/06/01
'博    客：http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                        2012/12/03
'          V1.1 修正了几个正则相关的函数，调整了部分类结构                  2013/05/28
'          V1.2 增加属性Caption，可以获取或设置当前标题栏                   2013/05/29
'          V1.3 增加了方法Focus，可以激活当前窗口                           2013/06/01
'               增加了方法Left,Top,Width,Height,Move，处理窗口位置等
'          V1.4 增加了窗口位置调整的几个函数                                2013/06/04
'               增加了得到应用程序路径的函数AppName
'               增加了得到应用程序启动参数的函数AppCommandLine
'          V1.5 增加了窗口最大最小化，隐藏显示正常的几个函数                2013/06/06
'               增加了获取控件相关函数是否使用正则的参数UseRegExp默认F
'          V1.6 将Left，Top函数改为属性，可获得可设置                       2013/06/10
'          V1.7 增加函数：CloseApp 结束进程                                 2013/06/13
'               修正了部分跟正则匹配相关的函数
'               增加函数：GetElementTextByText
'               增加函数：GetElementHwndByText
'          V1.8 增加函数：GetWindowByClassName                              2013/06/26
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
'          V1.9 修正函数：GetMatchHwndFromWindow 正则表达式的错误           2013/08/07
'               修正函数：GetMatchHwndFromWindow 函数中的一些错误           2014/09/23
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByPID 根据PID取窗口句柄
'               增加函数：GetCaptionByHwnd 根据句柄取得标题
'               增加函数：SetTop设置窗体置顶，传入参数false则取消           2014/09/24
'               增加函数：Shake、FadeIn、FadeOut 抖动、淡入、淡出特效
'          V2.0 修正函数：GetWindowByPID 遍历窗体Win7下有一些问题           2015/09/29
'               修正函数：GetWindowByAppName 遍历窗体Win7下有一些问题
'               修正函数：GetWindowByAppNameEx 遍历窗体Win7下有一些问题
'          V2.1 修正函数：ClickPoint 增加位置模式参数相对和绝对，默认相对   2018/06/05
'               增加函数：SelectComboBoxIndex 根据指定的index选择下拉框中的项
'                         上述方法得到网友Chen8013的不少帮助，特此感谢
'               增加函数：GetWindowByHwnd 根据指定的句柄确定窗口            2018/07/22
'               增加函数：GetWindowByCursorPos 根据当前光标获取窗口（控件）
'               增加函数：GetWindowByPoint 根据指定的位置获取窗口（控件）
'               升级ClickPoint函数，支持点击前后分别延时，默认延时为0       2018/07/23
'==============================================================================================
