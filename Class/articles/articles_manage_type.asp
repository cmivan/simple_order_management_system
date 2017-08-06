<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" class="forMy forMAargin">
<TR>
<TD height="20" colspan="4" align="left" class="forumRow TypeNav" width="100%">
<%if db_table="orders" or db_table="user" then%>

<%
if UserID<>"" and isnumeric(UserID) then
  if cint(UserID)=cint(session("LoginID")) then
     UserIDName="<span class=red>我的</span> "
  else
  set rs=conn.execute("select * from user where id="&UserID)
      if not rs.eof then UserIDName="<span class=red>"&rs("UserName")&"</span> "
  set rs=nothing
  end if
end if
'/////////////
%>

<a hidefocus href="?UserID=<%=UserID%>"><%=UserIDName%><strong>全部<%=db_title%></strong></a>
 <%	
'### 搜索关键词 ####
    keyword=request("keyword")
set rs=server.createobject("adodb.recordset") 
	exec="select * from "&db_table&"_type where type_id=0 order by order_id asc" 
	rs.open exec,conn,1,1 
	do while not rs.eof
	'### 用于样式 ####
	if cstr(rs("id"))=cstr(typeB_id) then
	   styleB=" class=""on"""
	   else
	   styleB=""
	end if
%><a hidefocus href="?UserID=<%=UserID%>&typeB_id=<%=rs("id")%>&keyword=<%=keyword%>" <%=styleB%>>&nbsp;<%=rs("title")%>&nbsp;</a><%
	rs.movenext
	loop
    rs.close
set rs=nothing
%>
<%
else
response.Write("&nbsp;")
end if%>
</td>
<td class="forumRow">
<table border="0" align="right" cellpadding="0" cellspacing="0">
<Form name="search" method="get" action="">
 <tr>
<td><input name="keyword" type="text" id="keyword" style="font-size: 9pt" value="<%=request("keyword")%>" /></td>
<td width="70" align="center"><input name="submit" type="submit" class="INPUTBottom" value="<%=db_title%>搜索" align="absMiddle" width="43" height="18" border=0 />
 <input name="typeB_id" type="hidden" id="typeB_id" style="font-size: 9pt" value="<%=typeB_id%>" size="25" />
 <input name="typeS_id" type="hidden" id="typeS_id" style="font-size: 9pt" value="<%=typeS_id%>" size="25" />
 <input name="UserID" type="hidden" id="UserID" style="font-size: 9pt" value="<%=UserID%>" size="25" />
</td>
</tr>
</FORM>
</table></td>
</TR>

<%if typeB_id<>"" and isnumeric(typeB_id) then%>
<%	
set rs=server.createobject("adodb.recordset") 
	exec="select * from "&db_table&"_type where type_id="&typeB_id&" order by order_id asc" 
	rs.open exec,conn,1,1

if not rs.eof then
response.write "<TR><TD colspan=""3"" align=""left"" class=""forumRow"" style=""padding-left:50px;"">"
	do while not rs.eof
	'### 用于样式 ####
	if cstr(rs("id"))=cstr(typeS_id) then
	   styleB=" style=""color:#FF0000;"""
	   else
	   styleB=""
	end if
	
%>
&nbsp;- <a hidefocus href="?UserID=<%=UserID%>&typeB_id=<%=typeB_id%>&typeS_id=<%=rs("id")%>&keyword=<%=keyword%>" <%=styleB%>><%=rs("title")%></a>&nbsp;<%
	rs.movenext
	loop
response.write "</td></TR>"
end if
    rs.close
set rs=nothing
%>

<%end if%>
</TABLE>
