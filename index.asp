<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CodePage=65001%>
<%Response.Charset="UTF-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>欢迎用户 <%=session("realname")%> 登陆! 当前IP为:<%=RealIp%></title>
<link href="Class/css.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="Class/js.js"></script>
</head>

<!--#include file="Class/config.asp"-->
<!--#include file="Class/function.asp"-->
<%
'//////////版权信息( GO )////////////
'   网站生成系统 By 天方。雨
'   Email:619835864@qq.com
'   Time: 2010-06-09
'//////////版权信息( END )///////////
  session.Timeout=45

'× --------------------------------------
'× ----------  数据库连接部分 -------------
'× --------------------------------------

  'ON ERROR RESUME NEXT
  DIM CONNS,CONNSTR,TIME1,TIME2,MDB,BackUP_MDB
      TIME1=TIMER
      MDB="edit/fckeditor.mdb"
      BackUP_MDB=year(now)&month(now)&day(now)&"."&session("LoginUserID")
      BackUP_MDB=BackUPpath&BackUP_MDB&".mdb"


      SET conn=SERVER.CREATEOBJECT("ADODB.CONNECTION")
          connSTR="DRIVER=MICROSOFT ACCESS DRIVER (*.MDB);DBQ="+SERVER.MAPPATH(MDB)
          conn.OPEN connSTR
'///////////////////////////////
  IF ERR THEN
     ERR.CLEAR
     SET conn = NOTHING
     call backPage(errConnStr&err.description,"javascript:;",0)
  END IF

%>


<%
'////////////////////////////////////////////////////////////
if session("LoginID")<>"" and isnumeric(session("LoginID")) then
'/////// 登录，则自动备份数据 /////////
'创建目录
    call web_create(BackUPpath,"","folder")

 if request.QueryString("loginIN")="yes" then
    call CopyFiles(SERVER.MAPPATH(MDB),BackUP_MDB)
    if session("LoginToUrl")<>"" then
       response.Redirect(session("LoginToUrl"))
    else
       response.Redirect("Class/notice/article_manage.asp?userID="&session("LoginID"))
    end if
 end if
end if




' 检测自动登陆 ------------
   LoginTempPass = request.Cookies("LoginTempPass")
   LoginS_IP     = request.Cookies("LoginS_IP")
   LoginS_id     = request.Cookies("LoginS_id")
if LoginTempPass<>"" and LoginS_IP<>"" and LoginS_id<>"" and isnumeric(LoginS_id) then
set rs=conn.execute("select top 1 * from user where ID="&LoginS_id&" and LoginS='1'") 
 if not rs.eof then 
    if rs("ok")<>1 then
       conn.execute("update user set LoginS='0' where ID="&LoginS_id&"") 
       call backPage("登录失败,该账号未通过审核!","?",0)
    else
 '#### 二重验证 ####-- 保证正常登陆
    if rs("LoginS_IP")=LoginS_IP and rs("LoginTempPass")=LoginTempPass then
	  session("LoginID")        = rs("id")
      session("LoginUserID")    = rs("UserID")
	  session("Loginrealname")  = rs("UserName")
	  session("LoginTypeId")    = rs("typeB_id")
	  session("Loginpower")     = rs("UserPower")
	  session("LoginSuperPower")= rs("SuperPower")
	  session("LoginToUrl")     = rs("LoginToUrl")

      response.redirect "?loginIN=yes"
    end if
    end if
 end if
set rs=nothing
end if
%>

<table border="0" align=center cellpadding=6 cellspacing=0 bgcolor="#C4D8ED" id="main_box">
<form id="ADMIN_EDIT_FORM" name="ADMIN_EDIT_FORM" method="post" action="" >
<tr>
<td class="forumRow">
<%if UpdateTip<>"" then%>
<table width="225" border="0" align=center cellpadding=3 cellspacing=0 bordercolor="#FFFFFF" bgcolor="#B4B4B4">
  <tr>
    <td colspan="2" align="left" class="forumRaw">&nbsp;&nbsp;温馨提示：<span style="font-weight: bold"></span></td>
  </tr>
  <tr>
    <td height="50" colspan="2" align="center" class="forumRow"><%=UpdateTip%></td>
  </tr>
</table>
<%else%>
<table width="100%" border="0" align=center cellpadding=3 cellspacing=0 bordercolor="#FFFFFF" bgcolor="#B4B4B4">
<TR>
    <TD colspan="2" align="center" class="forumRaw">&nbsp;&nbsp;<span style="font-weight: bold">管理登录</span></TD>
    </TR>
<TR>
  <TD width="50" align="center" class="forumRow">
账号</TD>
  <TD class="forumRow">
  <span class="forumRow">
  <input NAME=username TYPE='text' class="input_len" id="username" value="zhang" />
  </span></TD>
  </TR><TR>
    <TD align="center" class="forumRow">
密码</TD>
    <TD class="forumRow">
  <input name=password type='password' class="input_len" id="password" value="3389007" /></TD>
</TR>


<tr>
<td align="center" class="forumRow">
  <span class="forumRow">
  <input name="FORM_TYPE" type="hidden" value="yes" />
  &nbsp;&nbsp;</span></td>
<td align="left" class="forumRow"><input class="INPUTBottom" name="button" type="submit" id="button" value="Go 登陆 !" /></td>
</tr>
</table>
<%end if%>
</td>
</tr>
</form>
</table>

<% 
username=request.form("username")
password=request.form("password")
loginS  =request.form("loginS")

if request.Form("FORM_TYPE")="yes" then
'####################################################
if username=""  then call backPage("用户名不能为空","",2)
if password="" then call backPage("密码不能为空","",2)

    sql="select top 1 * from user where UserID='"&username&"' and password='"&password&"'" 
set rs=conn.execute(sql) 
if rs.eof or rs.bof then 
   call backPage("帐号或者密码错误，请重新输入!","",2)
else

if rs("ok")<>1 then
   
   call backPage("登录失败,该账号未通过审核!","?",0)
else
   '#### 二重验证 ####-- 保证正常登陆
   if password<>rs("password") then
      call backPage("帐号或者密码错误，请重新输入!","",2)
   else
	  session("LoginID")        = rs("id")
      session("LoginUserID")    = rs("UserID")
	  session("Loginrealname")  = rs("UserName")
	  session("LoginTypeId")    = rs("typeB_id")
	  session("Loginpower")     = rs("UserPower")
	  session("LoginSuperPower")= rs("SuperPower")
	  session("LoginToUrl")     = rs("LoginToUrl")
      response.redirect "?loginIN=yes"
   end if
end if

end if
set rs=nothing

end if
%>

<%call reSetSize()%>
