<!--配置文件-->
<!--#include file="article_config.asp"-->
<!--#include file="../../edit/fckeditor.asp" -->

<%
on error resume next
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
if Tf_UserID<>"" then rs("UserID") =Tf_UserID
rs("endTime")=Tf_endTime
if edit_id="" or isnumeric(edit_id)=false then rs("insterID")=session("LoginID")

    end if

if err=0 then
   rs.update
   call backPage(editStr&"操作成功!","article_manage.asp?UserID="&UserID&"&typdB_id="&typdB_id&"&typdS_id="&typdS_id,0)
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
Tf_insterID  =rs("insterID")


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
	<script type="text/javascript" src="../../Edit/js/mootools.js"></script>
	<script type="text/javascript" src="../../Edit/js/calendar.rc4.js"></script>
	<script type="text/javascript">		
	//<![CDATA[
		window.addEvent('domready', function() { 
			myCal1 = new Calendar({ Tf_endTime: 'Y-m-d' }, { direction: 0, tweak: {x: 6, y: 0} });
		});
	//]]>
	</script>
<link rel="stylesheet" type="text/css" href="../../Edit/css/calendar.css" media="screen" />

<!--#include file="../top.asp"-->
<td valign="top" class="forumRow">
<table width="100%" border="0" align=center cellpadding=5 cellspacing=0 style="border:0px">
<tr>
<td class="forumRow">
<%
'////////////////////////////
if Power("2_6") then
   total=2
else
   total=1
end if
%>
<div class="tab_nav">
<a href="javascript:void(0);" id="title_1_1" class="on"><span style="font-weight: bold">通知消息</span></a></div>
<%
if edit_id<>"" and isnumeric(edit_id) then
   if session("Loginpower")=1 and cint(Tf_insterID)=cint(session("LoginID")) then
   elseif session("LoginSuperPower")=1 then 
   else
      SelStyle=" disabled"
   end if
end if
%>
<form name="article_update" method="post" action="" onSubmit="tijiao.disabled='disabled';">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="forMy">
<tr>
  <td width="70" align='right' class='forumRow'>消息：</td>
  <td class='forumRow'>
<input name="Tf_title" type="text" id="Tf_title" value="<%=Tf_title%>" size="40" <%=SelStyle%>/>
<select name="Tf_UserID" <%=SelStyle%>>
<option value="">请选择...</option>
<%
set CRs=conn.execute("SELECT * FROM user_type order by id asc")
    do while not CRs.eof
%>
<optgroup style="background-color:#CCCCCC;color:#FFFFFF" label="<%=CRs("title")%>"></optgroup>
<%
set collRs=conn.execute("SELECT * FROM user where typeB_id="&CRs("id")&" order by username asc,id desc")
    do while not collRs.eof
%>
<option value="<%=collRs("id")%>" <%if cint(Tf_UserID)=cint(collRs("id")) then%>selected style="background-color:#666666; color:#FFFFFF"<%end if%>>&nbsp;&nbsp;<%=collRs("UserName")%></option>
<%
    collRs.movenext
    loop
set collRs=nothing

    CRs.movenext
    loop
set CRs=nothing
%>
</select></td></tr>
<tr>
  <td width="70" align='right' class='forumRow'>任务时间：</td>
  <td class='forumRow'>
    <table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><input name="Tf_endTime" type="text" value="<%=Tf_endTime%>" size="40" style="float:left" <%=SelStyle%>></td>
        <td>&nbsp;
          <%if edit_id<>"" and isnumeric(edit_id) then%>
          <label>
          <input <%call getChecked(Tf_news)%> name="Tf_news" type="checkbox" id="Tf_news" value="1" disabled="disabled">
已经看到 </label>
&nbsp;&nbsp;
<label>
<input <%call getChecked(Tf_ok)%> name="Tf_ok" type="checkbox" id="Tf_ok" value="1" disabled="disabled">
处理完成 </label>
<%end if%></td>
        </tr>
    </table>
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

<table width="100%" border="0" align=center cellpadding=0 cellspacing=0 bordercolor="#FFFFFF" bgcolor="#B4B4B4" style="margin-top:6px;">
<tr>
<td width="72" class="forumRow">&nbsp;</td>
<td class="forumRow">
  <span class="forumRow">
<input class="INPUTBottom" name="tijiao" type="submit" id="tijiao" value="    提交<%=editStr%>    " <%=SelStyle%> />
<input onClick="history.back(1);" class="INPUTBottom" name="button" type="reset" id="button" value=" 返回列表 " />
<%if edit_id<>"" and isnumeric(edit_id) then%>
<input onClick="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?UserID=<%=UserID%>&act=del&id=<%=edit_id%>';}else{history.go(0);}" class="INPUTBottom" name="button3" type="reset" id="button3" value=" 删除 " <%=SelStyle%>/>
<%end if%>
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
  </span></td>
</tr>
</table>
</form>
</td></tr></table>

<!--#include file="../bottom.asp"-->
<script type="text/javascript"> 
    function showTab(ID,onID,aNum){
        for(var i=1;i<=aNum;i++){
  document.getElementById('tab_'+ID+'_'+i).style.display='none';
  document.getElementById('title_'+ID+'_'+i).className='';
			}
  document.getElementById('tab_'+ID+'_'+onID).style.display='block'; 
  document.getElementById('title_'+ID+'_'+onID).className='on';
  }
</script>
</body>
</html>