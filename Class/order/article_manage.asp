<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
 typeB_id=request.QueryString("typeB_id")
 typeS_id=request.QueryString("typeS_id")
 userID  =request.QueryString("userID")
%>
<body>
<!--#include file="../top.asp"-->

<td width="100%" valign="top" class="forumRow">
<!--#include file="../articles/articles_manage_type.asp"-->
<%
	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and (L_SiteName like '%"&keyword&"%' OR User_Name like '%"&keyword&"%' OR User_Company like '%"&keyword&"%' OR L_Url like '%"&keyword&"%' OR L_Site like '%"&keyword&"%' OR L_Ip like '%"&keyword&"%' OR L_Note like '%"&keyword&"%')"
	   keyword_sql2=" where HSX<>'1' "&keyword_sql
	end if

    '/// 按会员ID查看
    if userID<>"" and isnumeric(userID) then
       userSql=" and (salerID="&userID&" or jishuID="&userID&")"
       keyword_sql=keyword_sql&userSql
'///////////////////////////
       if keyword_sql2<>"" then
          keyword_sql2=keyword_sql2&userSql
       else
          keyword_sql2=" where HSX<>'1' and (salerID="&userID&" or jishuID="&userID&")"
       end if
    end if

    if keyword_sql2="" then keyword_sql2=" where HSX<>'1'"


	'########## 定义排序字符 #############
	   order_sql=" order by order_id asc,id desc"
    
	if typeS_id<>"" and isnumeric(typeS_id)then
	   exec="select * from "&db_table&" where HSX<>'1' and typeS_id="&typeS_id&keyword_sql&order_sql
	elseif typeB_id<>"" and isnumeric(typeB_id)then 
	   exec="select * from "&db_table&" where HSX<>'1' and typeB_id="&typeB_id&keyword_sql&order_sql
	else
	   exec="select * from "&db_table&keyword_sql2&order_sql
	end if
	

	'### 用于热门，审核 按钮链接,批量移动修改等
  	     FullUrl="?userID="&userID&"&typeB_id="&typeB_id&"&typeS_id="&typeS_id&"&page="&page&"&keyword="&keyword&""

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
					  <td width="420" bgcolor="#DBF1F7" class="forumRaw">&nbsp;<%=db_title%>&nbsp;名称\标题</td>
						<TD width="35" align="center" bgcolor="#D8D8D8" class="forumRaw">备案</TD>
						<TD width="35" align="center" bgcolor="#D8D8D8" class="forumRaw">完成</TD>
						<TD width="35" align="center" bgcolor="#D8D8D8" class="forumRaw">上传</TD>
						<TD width="35" align="center" bgcolor="#D8D8D8" class="forumRaw">收款</TD>
						<td width="60" colspan="3" align="center" bgcolor="#DBF1F7" class="forumRaw">管理</td>
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
					  <%=replace(rs("L_SiteName"),keyword,"<span style='color:#FF0000'>"&keyword&"</span>")%>
					  <%else%>
					  <%=rs("L_SiteName")%>
					  <%end if%>
		  </a></td>
          <TD height="19" align="center" bgcolor="#D8D8D8"><%call getOK(rs("Web_BA"))%></TD>
          <TD align="center" bgcolor="#D8D8D8"><%call getOK(rs("Web_WC"))%></TD>
          <TD align="center" bgcolor="#D8D8D8"><%call getOK(rs("Web_SC"))%></TD>
          <TD align="center" bgcolor="#D8D8D8"><%call getOK(rs("Web_SK"))%></TD>
            <td width="25" align="center"><a href="http://www.<%=rs("L_Site")%>" target="_blank" hidefocus><img src="../../edit/images/url.gif" alt="访问网站" width="16" height="16" border="0" /></a></td>
			
<%if (cstr(userID)=cstr(session("LoginID")) or session("LoginSuperPower")=1) and session("LoginTypeId")=38 then%>
<td width="25" align="center"><a href="\\<%=Request.ServerVariables("LOCAL_ADDR")%>\<%=okPath%>\<%=session("Loginrealname")%>\<%=rs("L_SiteName")%>" target="_blank" hidefocus><img src="../../Edit/images/home.gif" width="10" height="11" border="0"></a></td>
<%end if%>

          <td align="center"><input title="录入时间:<%=rs("add_data")%>" name="update" type="button" class="INPUTBottom" onClick="window.location.href='article_edit.asp<%=getUrl("id",rs("id"))%>' " value="修改"/></td>
</tr>

<%if cint(LookID)=cint(i) and cint(CloseID)<>cint(i) then%>
<tr id="info_<%=i%>" style="display:none1">
<TD colspan="9" class="forumRow" style="padding:1px;">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" class="forMy forMAargin" style="word-break:keep-all" >
  <tr>
    <td width="50%" class="forumRow"><img src="../../Edit/images/type_ico.gif" width="12" height="12">客户信息</td>
    <td width="50%" class="forumRow"><img src="../../Edit/images/type_ico.gif" width="12" height="12">FTP信息</td>
    </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF">
<table width="100%" border="0" align=center cellpadding=2 cellspacing=0 >
<TR>
<TD align="right" valign="top">销售：</TD>
<TD valign="top">
<%
set salerrs=conn.execute("select * from user where id="&rs("salerID"))
	if not salerrs.eof then
       response.Write(salerrs("UserName"))
    else
       response.Write("&nbsp;")
    end if
SET salerrs=nothing
%></TD>
</TR>
<TR>
<TD align="right" valign="top">客户称呼：</TD>
<TD valign="top">
<%=rs("User_Name")%></TD></TR>
<TR>
<TD align="right" valign="top">固话：</TD>
<TD valign="top">
<%=rs("User_Tel")%></TD></TR>
<TR>
<TD align="right" valign="top">手机：</TD>
<TD valign="top">
<%=rs("User_Phone")%></TD></TR>
<TR>
<TD align="right" valign="top">Email：</TD>
<TD valign="top">
<%=rs("User_Email")%></TD></TR>
  <TR>
<TD align="right" valign="top">联系QQ：</TD>
<TD valign="top">
<%=rs("User_QQ")%></TD></TR>
<TR>
<TD align="right" valign="top">身份证号：</TD>
<TD valign="top">
<%=rs("User_SFZ")%></TD></TR>
<TR>
<TD align="right" valign="top">公司名称：</TD>
<TD valign="top">
<%=rs("User_Company")%></TD>
</TR>
<tr>
<TD width="100" align="right" valign="top">执照号/备案号：</TD>
<TD valign="top">
  <%=rs("User_BAH")%></TD></TR>
<tr>
<TD align="right" valign="top">公司地址：</TD>
<TD valign="top">
 <%=rs("User_Address")%></TD></TR>


<TR>
  <TD align="right" valign="top">
站名：</TD>
  <TD valign="top">
 <%=rs("L_SiteName")%></TD>
</TR>
<TR>
  <TD align="right" valign="top">
类型：</TD>
  <TD valign="top">
<%
ThisStyle="disabled"
%>
<!--#include file="../articles/articles_type.asp"--></TD>
</TR>
<TR>
<TD align="right" valign="top">域名：</TD>
<TD valign="top">
 <%=rs("L_Site")%></TD></TR>
<TR>
<TD align="right" valign="top">签单时间：</TD>
<TD valign="top">
 <%=rs("L_OrderTime")%></TD></TR>
</table></td>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" style="word-break:keep-all" >
<TR>
<TD width="100" align="right">技术：</TD>
<TD>
<%
set salerrs=conn.execute("select * from user where id="&rs("jishuID"))
	if not salerrs.eof then
       response.Write(salerrs("UserName"))
    else
       response.Write("&nbsp;")
    end if
SET salerrs=nothing
%></TD>
</TR>
 <tr>
        <td width="100" align="right">站名：</td>
        <td><%=rs("L_SiteName")%></td>
      </tr>
      <tr>
        <td align="right">控制面板：</td>
        <td><a hidefocus href="http://www.ns365.net/" target="_blank">http://www.ns365.net/</a></td>
      </tr>
      <tr>
        <td align="right">域名：</td>
        <td><a hidefocus href="http://www.<%=rs("L_Site")%>" target="_blank"><%=rs("L_Site")%></a></td>
      </tr>
      <tr>
        <td align="right">域名密码：</td>
        <td><%=rs("L_Site_Pass")%></td>
      </tr>

<TR>
    <TD colspan="2" align="left">&nbsp;&nbsp;<span style="font-weight: bold">[网站后台登陆]</span></TD>
    </TR>
<TR>
    <TD align="right">
后台地址：</TD>
    <TD>
<%=rs("L_Url")%></TD></TR>
	
<TR>
  <TD align="right">
账号：</TD>
  <TD>
<%=rs("L_UserID")%></TD>
  </TR><TR>
    <TD align="right">
密码：</TD>
<TD>
<%=rs("L_PassWord")%></TD>
</TR>

      <tr>
        <td colspan="2">&nbsp;&nbsp;<span style="font-weight: bold">[Ftp登陆信息]</span></td>
      </tr>
      <tr>
        <td align="right">服务IP：</td>
        <td><%=rs("L_IP")%></td>
      </tr>
      <tr>
        <td align="right">Ftp账号：</td>
        <td><%=rs("L_FtpUserID")%></td>
      </tr>
      <tr>
        <td align="right">Ftp密码：</td>
        <td><%=rs("L_FtpPassWord")%></td>
      </tr>
      <tr>
        <td align="right" valign="top">其他信息：</td>
        <td valign="top"><%=rs("L_Note")%></td>
      </tr>
    </table></td>
  </tr>
</table></TD>
</tr>
<%end if%>


<%
		rs.movenext
	next
%>

<tr>
<td height="20" colspan="9" align="center" class="forumRaw"><!--#include file="../articles/articles_paging.asp"--></td>
	    </tr>
	  </table>

<%
end if
%>

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