# VB6-Redis
【祭奠逝去的经典——VB6.0访问Redis缓存服务】(Access Redis with Visual Basic 6.0)

使用VB6.0工程引用mswinsck.ocx控件，创建TCP连接访问Redis缓存服务。并将常用的缓存变量读写、队列读写等操作功能封装成一个类RedisClass.

此项目完全开源，包括类文件和调用demo工程案例，欢迎使用、优化。


已知故障及解决方法：


故障：在类模块中使用set sckRedis=new Winsock时可能发生错误429，即无法创建ActiveX对象。

解决方法1：
  取消工程对mswinsck.ocx的引用，改为在“部件”中添加mswinsck.ocx控件，并拖到窗体form1界面上，即Winsock1控件。然后修改RedisClass类模块的代码为set sckRedis=form1.Winsock1即可。

解决方法2：
  使用vb6cli修复VB6许可证问题。然后重新编译生成exe，即可正常以“引用”方式调用mswinsck.ocx来创建socket对象实例。（但此方法生成的exe程序似乎只能在本机正常运行，在其它电脑运行依然会报错）
