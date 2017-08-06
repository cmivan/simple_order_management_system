<!--配置文件-->
<!--#include file="article_config.asp"-->
<body>
<!--#include file="../top.asp"-->

<td width="100%" valign="top" class="forumRow">
<!--#include file="../articles/articles_manage_type.asp"-->
<%
'////////优先处理，查看和完成/////////////////////////////
news_id=request.QueryString("news_id")
ok_id  =request.QueryString("ok_id")
if news_id<>"" then
   conn.execute("update "&db_table&" set news=1 where id="&news_id)
   response.Redirect(getUrl("news_id",""))
end if

if ok_id<>"" then
   conn.execute("update "&db_table&" set ok=1 where id="&ok_id)
   response.Redirect(getUrl("ok_id",""))
end if



'//////////////////////////////////////////////////////


       userID =request.QueryString("userID")
	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and (title like '%"&keyword&"%' OR note like '%"&keyword&"%' OR endTime like '%"&keyword&"%' OR add_data like '%"&keyword&"%')"
	   keyword_sql2=" where HSX<>'1' "&keyword_sql
	end if

    '/// 按会员ID查看
    if userID<>"" and isnumeric(userID) then
       userSql=" and UserID="&userID
       keyword_sql=keyword_sql&userSql
'///////////////////////////
       if keyword_sql2<>"" then
          keyword_sql2=keyword_sql2&userSql
       else
          keyword_sql2=" where HSX<>'1' and UserID="&userID
       end if
    end if

    if keyword_sql2="" then keyword_sql2=" where HSX<>'1'"


	'########## 定义排序字符 #############
	   order_sql=" order by news asc,id desc"
	   exec="select * from "&db_table&keyword_sql2&order_sql
	

	'### 用于热门，审核 按钮链接,批量移动修改等
  	     FullUrl="?userID="&userID&"&page="&page&"&keyword="&keyword&""


set rs=server.createobject("adodb.recordset")
	rs.open exec,conn,1,1 
	if rs.eof then
%>
<table width="100%" border="0" align="center" cellpadding="60" cellspacing="0" class="forMy forMAargin">
<tr>
<td align="center" bgcolor="#DBF1F7" class="forumRow">
暂无相应 <%=db_title%> 内容!</td>
</tr>
</table>	
	
<%else%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="forMy forMAargin">
					<tr>
					  <td width="380" bgcolor="#DBF1F7" class="forumRaw">&nbsp;<%=db_title%>&nbsp;名称\标题</td>

						<td align="center" bgcolor="#DBF1F7" class="forumRaw">任务时间</td>
						<TD width="40" align="center" bgcolor="#D8D8D8" class="forumRaw">查看</TD>
						<TD width="40" align="center" bgcolor="#D8D8D8" class="forumRaw">完成</TD>
						<td width="40" colspan="2" align="center" bgcolor="#DBF1F7" class="forumRaw">管理</td>
				    </tr>

<%
	rs.PageSize =pSize '每页记录条数
	iCount=rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage=rs.PageCount 
	page=request("page")
	
	if Not IsNumeric(page) or page="" then
	   page=1
	else
	   page=cint(page)
	end if
	
	if page<1 then
	   page=1
	elseif  page>maxpage then
	   page=maxpage
	end if
	
	rs.AbsolutePage=Page
	if page=maxpage then
	   x=iCount-(maxpage-1)*iPageSize
	else
	   x=iPageSize
	end if



'------->分页必须的变量
PAGE=PAGE                      '当前页数
LIST_NUM=3                     '
'FRIST_PAGE=1                   ' 首页
LAST_PAGE=rs.PAGECOUNT  ' 最后页


LookID=request.QueryString("LookID")
CloseID=request.QueryString("CloseID")

For i=1 To x

if LookID<>"" and cint(LookID)=cint(i) then
   clickStr="CloseInfo("&i&");"
else
   clickStr="ShowInfo("&i&");"
end if

%>
					
<tr onDblClick="<%=clickStr%>" onMouseOver="MM_Color(this,'on');" onMouseOut="MM_Color(this,'');" class="forumRow_out">
					  <td>&nbsp;<a href="article_edit.asp?UserID=<%=UserID%>&id=<%=rs("id")%>"><img src="../../Edit/images/home.gif" width="10" height="11">
					    <%if keyword<>"" then%>
					  <%=replace(rs("title"),keyword,"<span style='color:#FF0000'>"&keyword&"</span>")%>
					  <%else%>
					  <%=rs("title")%>
					  <%end if%>
		  </a></td>
                      <td align="center"><%=rs("endTime")%></td>
                      <TD align="center" bgcolor="#D8D8D8">

<%if rs("news")<>1 then%>
<input name="update2" type="button" class="INPUTBottom" onClick="window.location.href='<%=getUrl("news_id",rs("id"))%>' " value="查看"/>
<%else%>
<%call getOK(rs("news"))%>
<%end if%>

</TD>
<TD align="center" bgcolor="#D8D8D8">
<%if rs("news")=1 and rs("ok")<>1 then%>
<input name="update2" type="button" class="INPUTBottom"onClick="javascript:if(confirm('确定已经完成<%=db_title%> 《<%=rs("title")%>》吗？')){window.location.href='<%=getUrl("ok_id",rs("id"))%>';}else{history.go(0);}" value="完成"/>
<%elseif rs("ok")=1 then%>
<%call getOK(rs("ok"))%>
<%else%>-<%end if%>


</TD>
          <td align="center"><input name="update" type="button" class="INPUTBottom" onClick="window.location.href='article_edit.asp?UserID=<%=UserID%>&id=<%=rs("id")%>' " value="管理"/></td>
</tr>

<%if cint(LookID)=cint(i) and cint(CloseID)<>cint(i) then%>
<tr id="info_<%=i%>" style="display:none1">
<TD colspan="10" class="forumRow" style="padding:1px; background-color:#FFFFFF"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="6">
  <tr>
    <td align="left"><%=rs("note")%></td>
    </tr>
  <tr>
    <td align="left" style="color:#B9B9B9;">发布时间：<%=rs("add_data")%>&nbsp;

</td>
  </tr>
  
</table></TD>
</tr>
<%end if%>


<%
		rs.movenext
	next
%>

<tr>
<td height="20" colspan="10" align="center" class="forumRaw"><!--#include file="../articles/articles_paging.asp"--></td>
</tr>
</table>
<%end if%>
<!--#include file="../bottom.asp"-->
</body>
</html>
<%
if request("act")="del" then
	    id=Request.QueryString("id")
if id<>"" and isnumeric(id) then
	set rs=server.createobject("adodb.recordset")
	    sql="select * from "&db_table&" where id="&id
	    rs.open sql,conn,1,3
        rs("HSX")="1"
	    rs.update
	    Response.Write "<script>alert('"&db_title&"刪除成功！');window.location.href='"&FullUrl&"';</script>" 
        rs.close
	set rs=nothing
end if
end if
%>