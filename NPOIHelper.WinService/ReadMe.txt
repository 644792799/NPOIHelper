http://jingyan.baidu.com/article/fa4125acb71a8628ac709226.html  C#创建Windows服务与安装-图解
http://www.cnblogs.com/beniao/archive/2009/03/24/1420711.html   Net Remoting的双向通信和Windows Service的宿主服务
http://blog.csdn.net/kntao/article/details/6181133 Remoting 配置文件遇到的问题
http://blog.csdn.net/abc326321003/article/details/19162837 C#WinForm - 最小化或关闭时隐藏到系统托盘
http://blog.csdn.net/bigpudding24/article/details/50369904 emoting与socket、web service的比较及实例


错误描述
 
Microsoft Office Excel 不能访问文件“a.xls”。 可能的原因有:
? 文件名称或路径不存在。 
? 文件正被其他程序使用。 
? 您正要保存的工作簿与当前打开的工作簿同名
 
 
解决方案：
 
C:\Windows\System32\config\systemprofile和C:\Windows\SysWOW64\config\systemprofile目录下创建名为Desktop目录即可解决问题。


NPAPI是当今最流行的插件架构，由网景开发，后Mozilla维护，几乎所有浏览器都支持，不过存在很大的安全隐患，插件可以窃取系统底层权限，发起恶意攻击。

2010年，Google在原有网景NPAPI(Netscape Plugin API)基础上开发了新的PPAPI(Pepper Plugin API)，将外挂插件全部放到沙盒里运行，2012年Windows、Mac版本的Chrome浏览器先后升级了PPAPI Flash Player，并希望今年底值钱彻底淘汰NPAPI。

PPAPI的flash相较于NPAPI来讲，内存占用更大，因为全在沙盒里面运行