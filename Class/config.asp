<%
'<><><><>定义变量<><><><>

'//// 连接数据库时的错误提示 /////////////////
 dim errConnStr   
     errConnStr="暂时无法读取数据库，请见谅~"

'//// 控制提示信息，用于程序升级暂时停止 //////
 dim UpdateTip
	 'UpdateTip="程序升级中..."  'UpdateTip不为空则程序提示

'//// 分页显示的数目 /////////////////
 dim pSize   
     pSize=11
	 
'//// 定义订单任务文件存放目录 /////////////////
 dim okPath   
     okPath="完成任务"	 

'//// 定义订单任务文件存放目录 /////////////////
 dim Fpath   
     Fpath="e:\"&okPath&"\"
	 
'//// 数据备份目录 /////////////////
 dim BackUPpath   
     BackUPpath="E:\webSystemV1.2 DbBackup\"
%>
