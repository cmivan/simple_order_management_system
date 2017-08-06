<!--配置文件-->
<!--#include file="article_config.asp"-->
<!--#include file="../../edit/fckeditor.asp" -->
<%
'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
   UserID    =request("UserID")

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
Tf_title  =request.form("Tf_title")
Tf_note   =request.form("Tf_note")
Tf_UserID =request.form("Tf_UserID")
Tf_endTime=request.form("Tf_endTime")
Tf_news   =request.form("Tf_news")
Tf_ok     =request.form("Tf_ok")


rs("title")  =Tf_title
rs("note")   =Tf_note
if edit_id="" or isnumeric(edit_id)=false then rs("UserID") =session("LoginID")

rs("endTime")=Tf_endTime

    end if

if err=0 then
   rs.update
   call backPage(editStr&"操作成功!","article_manage.asp"&getUrl("UserID",session("LoginID")),0)
else
   call backPage("操作失败!原因:"&err.description,"",0)
end if


	rs.close
set rs=nothing




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
	
Tf_title  =rs("title")
Tf_note   =rs("note")
Tf_UserID =rs("UserID")
Tf_endTime=rs("endTime")
Tf_news   =rs("news")
Tf_ok     =rs("ok")


	   end if
	   rs.close
   set rs=nothing
   
else
  '/// 当前状态:添加
   edit_stat="add"
end if



  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
  if F_L_Site_Pass="" then F_L_Site_Pass="zc2010admin"

%>


<body>
<!--#include file="../top.asp"-->
<td valign="top" class="forumRow">
<form name="article_update" method="post" action="" onSubmit="tijiao.disabled='disabled';">
<table width="100%" border="0" align=center cellpadding=5 cellspacing=0 style="border:0px">
<tr>
<td class="forumRow">
<div class="tab_nav">
<a href="javascript:void(0);" id="title_1_1" class="on" style="font-weight: bold">工作日记</a></div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" class="forMy">
<tr>
  <td width="70" align='right' class='forumRow'>标题：</td>
  <td class='forumRow'>
<input name="Tf_title" type="text" id="Tf_title" value="<%=Tf_title%>" size="40" /></td>
</tr>
<tr style="display:none">
  <td width="70" align='right' class='forumRow'>时间：</td>
  <td class='forumRow'>
  <input name="Tf_endTime" type="text" value="<%=Tf_endTime%>" size="40">
</td>
</tr>
<tr>
  <td width="70" align='right' valign="top" class='forumRow' style="padding-top:5px;">详情：</td>
  <td class='forumRow'>
<%  
    Tf_note=replace(Tf_note,"'","&#39;")
	Set oFCKeditor = New FCKeditor 
	oFCKeditor.BasePath = "../../edit/"
	oFCKeditor.ToolbarSet = "Basic" 
	oFCKeditor.Width = "100%" 
	oFCKeditor.Height = "163" 
	oFCKeditor.Value = Tf_note
	oFCKeditor.Create "Tf_note"
    Set oFCKeditor = nothing
%></td></tr>
</table>

<table width="100%" border="0" align=center cellpadding=0 cellspacing=0 style="margin-top:6px;">
<tr>
<td width="72" class="forumRow">&nbsp;</td>
<td class="forumRow">
<input class="INPUTBottom" name="tijiao" type="submit" id="tijiao" value="    提交<%=editStr%>    ">
<input onClick="history.back(1);" class="INPUTBottom" name="button" type="reset" id="button" value=" 返回列表 ">
<%if edit_id<>"" and isnumeric(edit_id) then%>
<input onClick="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?UserID=<%=UserID%>&act=del&id=<%=edit_id%>';}else{history.go(0);}" class="INPUTBottom" name="button3" type="reset" id="button3" value=" 删除 ">
<%end if%>
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</td>
</tr>
</table>
</form>
</td></tr>
</table>
<!--#include file="../bottom.asp"-->
</body>