<%
'***********************************
'    重要的函数部分
'***********************************














'***********************************
'    更新所有网站目录
'***********************************
function cearPath()
if Fpath<>"" then
  dim newPath
  set rs=conn.execute("select * from user where typeB_id=38")
      do while not rs.eof
         if rs("UserName")<>"" then
            set rswww=conn.execute("select * from orders where jishuID="&cint(rs("id")))
                do while not rswww.eof
                   newPath=Fpath&"\"&rs("UserName")&"\"&rswww("L_SiteName")
                   newPath=replace(newPath,"/","\")
                   newPath=replace(newPath,"\\","\")
                '///为程序员生成，网站相关目录//////////////
                   call web_create(newPath,"","folder")
                   set rsPath=conn.execute("select * from setPath")
                       do while not rsPath.eof
                       if rsPath("tip")<>"" then call web_create(newPath&"\"&rsPath("tip"),"","folder")
                       rsPath.movenext
                       loop
                   set rsPath=nothing

                rswww.movenext
                loop
            set rswww=nothing
         end if
      rs.movenext
      loop
  set rs=nothing
end if
end function


'× --------------------------------------
'× ----------  备份文件 -----------
'× --------------------------------------
Function CopyFiles(TempSource,TempEnd)
    ON ERROR RESUME NEXT
    Dim FSO 
    Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
    IF FSO.FileExists(TempEnd) then 
      'Response.Write "目标备份文件 <b>" & TempEnd & "</b> 已存在，请先删除!" 
       Set FSO=Nothing 
       Exit Function 
    End IF 

    IF not FSO.FileExists(TempSource) Then 
      'Response.Write "要复制的源数据库文件 <b>"&TempSource&"</b> 不存在!" 
       Set FSO=Nothing 
       Exit Function 
    End If 

    FSO.CopyFile TempSource,TempEnd  
    Set FSO = Nothing 
End Function



'/////////////////////////////////////////////////////////////////////////////////
'-->> 生成文件/目录 的函数(go)    sub web_create("文件名","文件内容","类型")
'/////////////////////////////////////////////////////////////////////////////////

function web_create(s_name,s_centent,s_type)
    on error resume next  '容错模式(防止内存不足)
    select case s_type
	       case "folder"                          '文件夹不存在则生成
		       'folder_name=lcase(server.mappath(s_name))
		        folder_name=lcase(s_name)
	            folder_name=replace(folder_name,"/","\")
		        folder_name=split(folder_name,"\")
				i_paht=ubound(folder_name)
				if instr(folder_name(i_paht),".")<>0 then i_paht=i_paht-1
				
		    set fso=createobject("scripting.filesystemobject")
				for i=0 to i_paht
            		if folder_name(i)<>"" then
		      		   folder_names=folder_names&folder_name(i)&"\"
			   		if fso.folderexists(folder_names)=false then fso.createfolder(folder_names)
	        		end if 
				next
	        set fso=nothing 


		   case "file"                            '文件不存在则生成文件
		   
'               生成完整路径
		        temp_paths=lcase(s_name)
				temp_path =split(temp_paths,"\")
		        temp_path_str=temp_path(ubound(temp_path))
				full_path =replace(temp_paths,temp_path_str,"")
				call web_create(full_path,,"folder")

				set stm=server.createobject("adodb.stream") 
   				    stm.type=2 '以本模式读取 
   				    stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.writetext s_centent 
    				stm.savetofile server.mappath(s_name),2 
    				stm.flush 
    				stm.close
				set stm=nothing

		   case "getfile"                        '读取文件内容
				dim str 
				set stm=server.createobject("adodb.stream") 
    				stm.type=2 '以本模式读取 
    				stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.loadfromfile server.mappath(s_name) 
    				str=stm.readtext 
    				stm.close
				set stm=nothing 
   				    web_create=str 

		   case "copy"                        '复制文件
   				set fso=createobject("scripting.filesystemobject") 
   				set c=fso.getfile(s_name) '被拷贝文件位置
       				c.copy(s_centent)
   				set c=nothing
   				set fso=nothing
    end select
end function





Function GetLocationURL()
   Dim Url 
   Dim ServerPort,ServerName,ScriptName,QueryString 
       ServerName = Request.ServerVariables("SERVER_NAME") 
       ServerPort = Request.ServerVariables("SERVER_PORT") 
       ScriptName = Request.ServerVariables("SCRIPT_NAME") 
       QueryString = Request.ServerVariables("QUERY_STRING") 
       Url="http://"&ServerName 
       If ServerPort <> "80" Then Url = Url & ":" & ServerPort 
       Url=Url&ScriptName 
       If QueryString <>"" Then Url=Url&"?"& QueryString 
       GetLocationURL=Url 
End Function

 


'× --------------------------------------
'× ----------  返回登陆的真实ip -----------
'× --------------------------------------
   RealIp=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If RealIp="" Then RealIp=Request.ServerVariables("REMOTE_ADDR") 

'× --------------------------------------
'× ----------  返回提示信息 ---------------
'× --------------------------------------
  function backPage(backStr,backUrl,backType)
    back =""
	back =back&"<meta http-equiv=Content-Type content=text/html; charset=utf-8 />"
	back =back&"<link href='"&Rpath&"../Class/css.css' rel='stylesheet' type='text/css' />"
	
	if backType=0 then
	    'meta自动跳转到指定页面
        back =back&"<meta http-equiv=refresh content=1;url="&backUrl&">"
		back =back&"<body style=""overflow:hidden;"">"
	    back =back&"<br><TABLE border=0 align=center cellpadding=0 cellspacing=10 bgcolor=#FFFFFF><tr><td width=90% class=forumRow>"
		back =back&"<table width=100% border=0 align=center cellpadding=1 cellspacing=0 class=forMy><tr><td class=forumRow align=center>"
		back =back&backStr
		back =back&"</tr></table></td></tr></table>"
	elseif backType=1 then
	    'js弹出提示，返回指定页面
	    back =back&"<script language='javascript'>alert('"&backStr&"');"
		back =back&"window.location.href='"&backUrl&"';</script>"
	elseif backType=2 then
	    'js弹出提示，返回上一级
	    back =back&"<script language='javascript'>alert('"&backStr&"');history.back(1);</script>"
	elseif backType=3 then
	    'js弹出提示，返回指定页
	    back =back&"<script language='javascript'>window.location.href='"&backUrl&"';</script>"
	elseif backType=4 then
	    'js弹出提示，关闭窗口
	    back =back&"<script language='javascript'>alert('"&backStr&"');window.close();</script>"
	elseif backType=5 then
	    'js弹出提示，纯提示
	    back ="<script language='javascript'>alert('"&backStr&"');</script>"

	end if
	response.Write(back)
	response.End()
 end function

'× --------------------------------------
'× ---------  随机生成16位密码 ------------
'× --------------------------------------
 function getRunPass()
   Dim intCounter, intDecimal, strPassword
   For intCounter = 1 To 16
       Randomize
       intDecimal =  Int((26 * Rnd) + 1) + 64
       strPassword = strPassword & Chr(intDecimal)
   Next
     getRunPass=strPassword
 end function


'× --------------------------------------
'× ---------  返回相应的操作权限 -----------
'× --------------------------------------
 function Power(Uid)
     Power=false
   select case Uid
     case "6"       '//////// 只有超级权限用户可以操作
     if session("Loginpower")=6 then Power=true
     case "2_6"     '//////// 只有超级权限\技术员用户可以操
     if session("Loginpower")=2 or session("Loginpower")=6 then Power=true
     case "1"       '//////// 当前类型用户
     if session("Loginpower")=1 then Power=true
   end select

 end function
 
 
 
 '//// 返回是否符合空间的
 sub getOK(id)
    if id="1" or id=1 then
     response.Write("<div class='red'>√</div>")
    else
	 response.Write("&nbsp;")
    end if
 end sub

'× --------------------------------------
'× -------  用于返回热门最新推荐审核等  -----
'× --------------------------------------
  function getBool(num)
  	  if num="1" then
	     getBool =1
	  else
	     getBool =0
	  end if
  end function


 function getChecked(id)
    if cstr(id)="1" or cint(id)=1 then
       response.Write(" checked=""checked""")
    else
       response.Write("")
    end if
 end function



'× --------------------------------------
'× -------  用于返回Url值 ,重要 -----
'× --------------------------------------
function getUrl(key,value)
    on error resume next
    dim UrlStr,UrlKey,ReUrl,toUrl,NewUrl
    key   =lcase(key)
    UrlStr0=lcase(request.QueryString)
	UrlStr="?"&UrlStr0
    UrlKey=request.QueryString(key)

	if instr(UrlStr,"?"&key&"="&UrlKey)<>0 then
	   ReUrl =lcase("?"&key&"="&UrlKey)
       toUrl =lcase("?"&key&"="&value)
	   NewUrl=replace(UrlStr,ReUrl,toUrl)
	elseif instr(UrlStr,"&"&key&"="&UrlKey)<>0 then
	   ReUrl =lcase("&"&key&"="&UrlKey)
       toUrl =lcase("&"&key&"="&value)
	   NewUrl=replace(UrlStr,ReUrl,toUrl)
	else
	   if UrlStr0="" then
	      NewUrl=UrlStr&key&"="&value
	   else
	      NewUrl=UrlStr&"&"&key&"="&value
		  NewUrl=replace(NewUrl,"?&","?")
	   end if
	end if
	
    getUrl=NewUrl
end function
%>



<%sub reSetSize()%>
<div id="Set_H_V">500,300</div>
<script>
//设置宽高////
var main_H=document.getElementById("main_box").offsetHeight;
var main_W=document.getElementById("main_box").offsetWidth;
document.getElementById("Set_H_V").innerHTML=main_W+","+main_H;
</script>
<%end sub%>