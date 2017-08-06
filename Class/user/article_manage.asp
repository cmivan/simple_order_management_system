<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
 typeB_id=request.QueryString("typeB_id")
 typeS_id=request.QueryString("typeS_id")
%>
<body>
<!--#include file="../top.asp"-->

<td width="100%" valign="top" class="forumRow">
<!--#include file="../articles/articles_manage_type.asp"-->
<%


'--------------------------------------------
r_ok=request.QueryString("ok")
r_hot=request.QueryString("hot")
r_id=request.QueryString("r_id")

if r_id<>"" and isnumeric(r_id) then
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&r_id
       rs.open exec,conn,1,3 
	   if not rs.eof then
	      
		  if r_ok<>"" and int(r_ok)=1 then
		     rs("ok")=0
		  elseif r_ok<>"" and int(r_ok)=0 then
		     rs("ok")=1
		  end if
		  
	      if r_hot<>"" and int(r_hot)=1 then
		     rs("hot")=0
		  elseif r_hot<>"" and int(r_hot)=0 then
		     rs("hot")=1
		  end if

	   rs.update
	   end if
	   rs.close
   set rs=nothing
end if
 
'---------------------------------------------
	
	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and UserName like '%"&request("keyword")&"%'"
	   keyword_sql2=" where UserName like '%"&request("keyword")&"%'"
	end if
	
	'########## 定义排序字符 #############
	   order_sql=" order by UserName asc,id desc"
    
	if typeS_id<>"" and isnumeric(typeS_id)then
	   exec="select * from "&db_table&" where typeS_id="&typeS_id&keyword_sql&order_sql
	elseif typeB_id<>"" and isnumeric(typeB_id)then 
	   exec="select * from "&db_table&" where typeB_id="&typeB_id&keyword_sql&order_sql
	else
	   exec="select * from "&db_table&keyword_sql2&order_sql
	end if

set rs=server.createobject("adodb.recordset")
	rs.open exec,conn,1,1 
	if rs.eof then
%>
<table width="100%" border="0" align="center" cellpadding="50" cellspacing="0" class="forMy forMAargin">
<tr>
<td align="center" class="forumRow">
暂无相应 <%=db_title%> 内容!</td>
</tr>
</table>	
	
<%else%>
<form name="news_category" method="post">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="forMy forMAargin">
					<tr>
					  <td class="forumRaw">&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<%
set ThisRS=conn.execute("select * from orders_type where type_id=0 order by order_id asc,id asc")
    ArrNum=0
    redim DbCount(3)
    do while not ThisRS.eof
    ArrNum=ArrNum+1
    DbCount(ArrNum)=conn.execute("select count(id) from orders where typeB_id="&ThisRS("id"))(0)
%>
<td width="100" align="center" class="forumRaw"><%=ThisRS("title")%>(<%=DbCount(ArrNum)%>)</td>
<%
    ThisRS.movenext
    loop
set ThisRS=nothing
%>

<td width="25" align="center" class="forumRaw">总</td>

<%if session("LoginSuperPower")=1 then%>
<td width="20" align="center" class="forumRaw">审</td>
<%end if%>

<td width="150" align="center" class="forumRaw">管理</td>
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


For i=1 To x
%>
					
<tr onMouseOver="MM_Color(this,'on');" onMouseOut="MM_Color(this,'')" class="forumRow_out" title="录入时间:<%=rs("add_data")%>">
					  <td class="forumRow">
<a href="article_edit.asp?id=<%=rs("id")%>">&nbsp;<img src="../../Edit/images/home.gif" width="10" height="11">
<%if keyword<>"" then%>
					  <%=replace(rs("UserName"),keyword,"<span style='color:#FF0000'>"&keyword&"</span>")%>
					  <%else%>
					  <%=rs("UserName")%>
					  <%end if%>
					  </a></td>

<%
dim AllOrders
    AllOrders=0
set ThisRS=conn.execute("select * from orders_type where type_id=0 order by order_id asc,id asc")
    dim ArrNum
    ArrNum=0
    do while not ThisRS.eof
    ArrNum=ArrNum+1

'/////////////////////////////
    dim ThisNum
set conRS=server.CreateObject("adodb.recordset")
    conRSSql="select id from orders where HSX<>'1' and typeB_id="&ThisRS("id")&" and (salerID="&rs("id")&" or jishuID="&rs("id")&") order by id desc"
    conRS.open conRSSql,conn,1,1
    if not conRS.eof then
       ThisNum=conRS.recordcount
    else
       ThisNum=0
    end if
    conRS.close
set conRS=nothing

    AllOrders=cint(AllOrders)+cint(ThisNum)

'/////////////////////////////
    dim CountNum
 if DbCount(ArrNum)<>"" and isnumeric(DbCount(ArrNum)) then
    CountNum=ThisNum/DbCount(ArrNum)*90
 else
    CountNum=0
 end if
%>
<td align="left" class="forumRaw" title="该位置显示相应订单的总数! ">
<table width="90" border="0" align="center" cellpadding="0" cellspacing="0" style="border:#66CCFF 1px solid; border-bottom:0; border-right:0;" title="总记录数:<%=ThisNum%>">
  <tr>
<td align="left"><div style="font-size:8px;color:#FFFFFF;width:<%=int(CountNum)%>px; background-image:url(../../Edit/images/bg100.jpg);"></div></td>
  </tr>
</table></td>

<%
    ThisRS.movenext
    loop
set ThisRS=nothing
%>


<td align="center" class="forumRaw" title="该位置显示相应订单的总数! "><%=AllOrders%></td>

<%if session("LoginSuperPower")=1 then%>
<td width="18" align="center" class="forumRow" title="通过审核的用户才可以正常登录! ">
<%
	'### 用于热门，审核 按钮链接,批量移动修改等
  	     FullUrl="?typeB_id="&typeB_id&"&typeS_id="&typeS_id&"&page="&page&"&keyword="&keyword&""
%>

<%if rs("ok")=1 then%>
<a href="<%=FullUrl%>&ok=1&r_id=<%=rs("id")%>" class="yes">√</a>
<%else%>
<a href="<%=FullUrl%>&ok=0&r_id=<%=rs("id")%>" class="no">×</a>
<%end if%></td>
<%end if%>

<td align="center" class="forumRow">
<%
'///// 获取通知总数 //////
if session("LoginID")<>"" and isnumeric(session("LoginID")) then
SET noticeRS=SERVER.CREATEOBJECT("ADODB.RECORDSET")
    noticeRS_STR="SELECT * FROM notice where UserID="&session("LoginID")&" and news<>1 order by id desc"	
	noticeRS.OPEN noticeRS_STR,conn,1,1
    if not noticeRS.eof then
       noticeNum="("&noticeRS.recordcount&")"
       noticePic="notice.gif"
    else
       noticeNum=""
       noticePic="notice_off.gif"
    end if
    noticeRS.close
set noticeRS=nothing
else
       noticeNum=""
       noticePic="notice_off.gif"
end if
%>



<button class="INPUTBottom" onClick="window.location.href='article_edit.asp?id=<%=rs("id")%>' " >编辑</button>
<button class="INPUTBottom" onClick="window.location.href='../notice/article_manage.asp?userID=<%=rs("id")%>'" style="background-image:url(../Edit/images/<%=noticePic%>)">通知</button>
<%if session("LoginSuperPower")=1 then%>
<button class="INPUTBottom" onClick="window.location.href='../diary/article_manage.asp?userID=<%=rs("id")%>' ">日记</button>
<%end if%>
<button class="INPUTBottom" onClick="window.location.href='../order/article_manage.asp?userID=<%=rs("id")%>' ">订单</button></td>
</tr>

<%
		rs.movenext
	next
%>

<tr <%=onTable%>>
<td height="20" colspan="13" align="center" class="forumRaw"><!--#include file="../articles/articles_paging.asp"--></td>
		  </tr>
		</table>
			
</form>

<%
end if
%>

<!--#include file="../bottom.asp"-->






<%
if errStr="ok" then
   if edits="del" then
      errStr="成功删除!"
   elseif edits="check" then
      errStr="成功审核!"
   elseif edits="not_check" then
      errStr="成功取消审核!"
   elseif edits="move" then
      errStr="成功移动!"
   end if
end if

if errStr<>"" then
   response.Write("<script>alert('"&errStr&"');window.location.href='"&FullUrl&"';</script>")
end if
%>
</body>
</html>
<%
if request("act")="del" then
	set rs=server.createobject("adodb.recordset")
	id=Request.QueryString("id")
	sql="select * from "&db_table&" where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	Response.Write "<script>alert('"&db_title&"刪除成功！');window.location.href='"&FullUrl&"';</script>" 
end if
%>