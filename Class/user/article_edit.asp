<!--配置文件-->
<!--#include file="article_config.asp"-->
<!--#include file="../../edit/fckeditor.asp" -->

<%
'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if
   
if request.Form("edit")="ok" then
'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,conn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,conn,1,3
	end if

	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

'——————接收数据
    type_id   =request.Form("type_id")
	add_data  =request.Form("add_data")
 if isdate(add_data)=false then add_data=now()

TypeID     = cstr(request.Form("TypeID"))
UserID     = cstr(request.Form("UserID"))
password   = cstr(request.Form("password"))
UserName   = cstr(request.Form("UserName"))
User_Tel   = cstr(request.Form("User_Tel"))
User_Phone = cstr(request.Form("User_Phone"))
User_Email = cstr(request.Form("User_Email"))
User_QQ    = cstr(request.Form("User_QQ"))
User_Note  = cstr(request.Form("User_Note"))
LoginS     = cstr(request.Form("LoginS"))
UserPower      = cint(request.Form("UserPower"))

if UserPower="1" then
   UserPower=1
else
   UserPower=0
end if


'——————排序
 if order_id="" or isnumeric(order_id)=false then order_id=0




'/////// 会员类型的添加，修改 做出限制
dim typePower
    typePower=false  '/// 初始化为flash
if edit_id="" or isnumeric(edit_id)=false then 
   if session("LoginSuperPower")=1 or session("Loginpower")=1 then typePower=true
else
   if session("LoginSuperPower")=1 then typePower=true
end if


'——————分类
  '//系判断是否为数组，再判断是否为数字(否则失败)
if typePower then
  type_ids=split(type_id,",")
  if ubound(type_ids)=1 then
     if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
	    typeB_id=type_ids(0)
	    typeS_id=type_ids(1)
	    session("type_id")=type_id   '记录本次操作分类
     else
	    response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	    response.End()
     end if
  else
     if type_id<>"" and isnumeric(type_id) then
	    typeB_id=type_id
	    typeS_id=null
	    session("type_id")=type_id   '记录本次操作分类
	 else
	    response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	    response.End()
	 end if
  end if
end if

'/////// 会员类型的添加，修改 做出限制
if typePower then
	  rs("typeB_id") =typeB_id
	  rs("typeS_id") =typeS_id
end if


	rs("add_data")=add_data


'/// 该类型只能添加，不能修改
if edit_id="" and isnumeric(edit_id)=false then
   rs("TypeID")=TypeID
   rs("UserID")=UserID
end if


rs("password")   =password
rs("UserName")   =UserName
rs("User_Tel")   =User_Tel
rs("User_Phone") =User_Phone
rs("User_Email") =User_Email
rs("User_QQ")    =User_QQ
rs("User_Note")  =User_Note
rs("LoginS")     =LoginS
rs("LoginS_IP")  =RealIp

if session("LoginSuperPower")=1 then
   rs("UserPower")  =UserPower
end if



if LoginS="1" then
   rs("LoginTempPass")=getRunPass()
   response.Cookies("LoginTempPass")=rs("LoginTempPass")
   response.Cookies("LoginS_IP")    =RealIp
   response.Cookies("LoginS_id")    =rs("id")

   response.Cookies("LoginTempPass").Expires=DateAdd("m",60,now())
   response.Cookies("LoginS_IP").Expires=DateAdd("m",60,now())
   response.Cookies("LoginS_id").Expires=DateAdd("m",60,now())
  '设置60个月以后过期
end if


	end if
	rs.update
	rs.close
set rs=nothing

if edit_id<>"" and isnumeric(edit_id) then
   call backPage(editStr&"操作成功!","?id="&edit_id,0)
else
   call backPage(editStr&"操作成功!","article_manage.asp?UserID="&UserID&"&typdB_id="&typdB_id&"&typdS_id="&typdS_id,0)
end if

end if





'#################  读取数据部分 #################

   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
  '/// 当前状态:编辑
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&id
       rs.open exec,conn,1,1 
	   if not rs.eof then
	
	      title     =rs("title")
	      content   =rs("content")
		  typeB_id  =rs("typeB_id")
	      typeS_id  =rs("typeS_id")
		  
		  '记录分类,用于分类下拉
		  if typeS_id="" or isnumeric(typeS_id)=false then
		     session("type_id")=typeB_id
		  else
		     session("type_id")=typeB_id&","&typeS_id
		  end if


          TypeID    =rs("TypeID")
          UserID    =rs("UserID")
          password  =rs("password")
          UserName  =rs("UserName")
          User_Tel  =rs("User_Tel")
          User_Phone=rs("User_Phone")
          User_Email=rs("User_Email")
          User_QQ   =rs("User_QQ")
          User_Note =rs("User_Note")
          LoginS    =rs("LoginS")  
          UserPower     =rs("UserPower")

		  if order_id="" or isnumeric(order_id)=false then order_id=0

	   end if
	   rs.close
   set rs=nothing
   
else
  '/// 当前状态:添加
   edit_stat="add"
end if


  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
  UserPower=cint(UserPower)
%>


<body>
<!--#include file="../top.asp"-->
<td valign="top" class="forumRow">
<form name="article_update" method="post" action="" onSubmit="tijiao.disabled='disabled';">
<table width="100%" border="0" align=center cellpadding=6 cellspacing=0>
<tr>
<td class="forumRow">
<div class="tab_nav">
<a href="javascript:void(0);" class="on"><span style="font-weight: bold"><%if session("LoginID")=cint(edit_id) then%>我的<%end if%>信息<%=editStr%></span></a></div>
<div>
<table width="100%" border="0" align=center cellpadding=0 cellspacing=0>
<TR>
<TD width="16%" align="right" class="forumRow">类型：</TD>
<TD width="84%" class="forumRow">
<%if session("LoginSuperPower")<>1 then ThisStyle="disabled"%>
<!--#include file="../articles/articles_type.asp"-->
<label>
  <input name="loginS" type="checkbox" id="loginS" value="1" <%if LoginS="1" then%>checked<%end if%>/>
  自动登陆</label>

<%if session("LoginSuperPower")=1 then%>
<label>
  <input name="UserPower" type="checkbox" id="UserPower" value="1" <%if cint(UserPower)=1 then%>checked<%end if%>/>
  管理员</label>
<%if session("LoginID")=cint(edit_id) then%>
、超级管理员
<%end if%>

<%else%>
<label>
  <input disabled="disabled" type="checkbox" <%if cint(UserPower)=1 then%>checked<%end if%>/>
  管理员</label>
<%end if%>



</TD></TR>
<TR>
<TD width="16%" align="right" class="forumRow">账号：</TD>
<TD width="84%" class="forumRow">
  <span class="forumRow">
  <input NAME=UserID TYPE='text' class="input_1" id="UserID" VALUE='<%=UserID%>' <%if EDIT_TYPE="EDIT" then%>disabled<%end if%> />
  </span></TD></TR>
<TR>
<TD width="16%" align="right" class="forumRow">密码：</TD>
<TD width="84%" class="forumRow">
  <span class="forumRow">
  <input NAME=password TYPE='password' class="input_1" id="password" VALUE='<%=password%>' />
  </span></TD></TR>
<TR>
<TD width="16%" align="right" class="forumRow">称呼：</TD>
<TD width="84%" class="forumRow">
  <span class="forumRow">
  <input NAME=UserName TYPE='text' class="input_1" id="UserName" VALUE='<%=UserName%>' />
  </span></TD></TR>
<TR>
<TD align="right" class="forumRow">固话：</TD>
<TD class="forumRow">
  <span class="forumRow">
  <input NAME=User_Tel TYPE='text' class="input_1" id="User_Tel" VALUE='<%=User_Tel%>' />
  </span></TD></TR>
<TR>
<TD align="right" class="forumRow">手机：</TD>
<TD class="forumRow">
  <span class="forumRow">
  <input NAME=User_Phone TYPE='text' class="input_1" id="User_Phone" VALUE='<%=User_Phone%>' />
  </span></TD></TR>
<TR>
<TD align="right" class="forumRow">Email：</TD>
<TD class="forumRow">
  <span class="forumRow">
  <input NAME=User_Email TYPE='text' class="input_1" id="User_Email" VALUE='<%=User_Email%>' />
  </span></TD></TR>
<TR>
<TD align="right" class="forumRow">QQ：</TD>
<TD class="forumRow">
  <span class="forumRow">
  <input NAME=User_QQ TYPE='text' class="input_1" id="User_QQ" VALUE='<%=User_QQ%>' />
  </span></TD></TR>
<tr>
<TD align="right" valign="top" class="forumRow">其他说明：</TD>
<TD class="forumRow">
<%  
    User_Note=replace(User_Note,"'","&#39;")
	Set oFCKeditor = New FCKeditor 
	oFCKeditor.BasePath = "../../edit/"
	oFCKeditor.ToolbarSet = "Basic" 
	oFCKeditor.Width = "100%" 
	oFCKeditor.Height = "120" 
	oFCKeditor.Value = User_Note
	oFCKeditor.Create "User_Note"
    Set oFCKeditor = nothing
%></TD></TR></table>
</div>


<table width="100%" border="0" align=center cellpadding=0 cellspacing=0 style="margin-top:6px;">
<tr>
<td width="16%" class="forumRow">&nbsp;</td>
<td width="84%" class="forumRow">
<span class="forumRow">
<input class="INPUTBottom" name="tijiao" type="submit" id="tijiao" value="    提交<%=editStr%>   ">
<input onClick="history.back(1);" class="INPUTBottom" name="button2" type="reset" id="button2" value=" 返回 ">
<%if edit_id<>"" and isnumeric(edit_id) then%>
<input onClick="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?UserID=<%=UserID%>&act=del&id=<%=edit_id%>';}else{history.go(0);}" class="INPUTBottom" name="button3" type="reset" id="button3" value=" 删除 ">
<%end if%>
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
  </span></td>
</tr>
</table></td>
</tr>
</table>
</form>
<!--#include file="../bottom.asp"-->
</body>
</html>