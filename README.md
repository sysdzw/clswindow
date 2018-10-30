# clswindow

			VB6操作外部程序窗口的类clsWindow使用说明

clsWindow是VB6环境下使用的一个操作外部程序窗口的类，比如得到窗口句柄，得到窗口里某个文本框的内容。非常方便，使用它可以让您脱身于一堆api函数，功能强大使用简单！

这个类楼主很早就开始封装了，原本打算做成类似DOM对象那样，通过一堆getElmentByXXX等方法实现对桌面程序下各个窗口以及里面各个控件对象的自由访问，但是具体要做的工作太多，目前只实现了一部分，期待大家一起加入更新维护。

目前该类封装了绝大部分对windows窗口的常用操作，例如：获取窗口句柄，设置窗口为活动窗口，设置窗口内文本框内容，点击窗口内的某些按钮等。

这个类现在还在一直不断地扩充，功能已经很强大很广泛，使用它可以轻而易举地设置窗口标题栏文字，移动窗体等等。以前要实现这些操作常常需要一大堆api函数，现在只需要一点点代码就可以了，完全让您脱身于api函数的海洋。当然您如果想知道里面的方法实现原理的话可以看一看源代码。



使用范例：
1）关闭腾讯新闻窗口“腾讯网迷你版”。
Dim window As New clsWindow
If window.GetWindowHwndByTitle("腾讯网迷你版") > 0 Then
    window.CloseWindow  '关闭窗口
End If
以上是不是很简洁呢？

2）获取某个打开的记事本里面的内容。假设记事本标题为“测试.txt - 记事本”，通过SPY等工具查看得知记事本的文本框类名为：Edit，那么我们编写程序如下：
Dim window As New clsWindow
If window.GetWindowHwndByTitle("测试.txt - 记事本") > 0 Then
    MsgBox window.GetElementTextByClassName("Edit")
End If
这个看起来也很简单，方法自由还可以使用正则匹配，可以写成下面这样：
Dim window As New clsWindow
If window.GetWindowHwndByTitleRegExp("测试\.txt.*?") > 0 Then
    MsgBox window.GetElementTextByClassName("Edi", , True)'第三个参数表示是否使用正则，默认为false
End If

v1.9以上版本已经可以使用连写功能。window.GetWindowHwndByTitle("腾讯网迷你版").CloseWindow 这样写是不是很酷呢？
更多演示案例：

类成员以及各个使用方法如下：


clsWindow2.1下载地址：
https://download.csdn.net/download/sysdzw/10558553

其他一些作品参考：
https://blog.csdn.net/sysdzw

http://blog.csdn.net/sysdzw/article/details/9083313
