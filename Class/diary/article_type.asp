<!--配置文件-->
<!--#include file="article_config.asp"-->
<script>
//进入编辑状态
function getEdit(id){
window.location.href='?id='+id;}
</script>


<body>
<%
'################| 处理删除排序问题 |################
   add_id_b=request.QueryString("add_id_b")
if add_id_b<>"" and isnumeric(add_id_b) then session("type_id")=add_id_b


'################| 处理删除排序问题 |################
   class_b_del=request.QueryString("class_b_del")
   class_s_del=request.QueryString("class_s_del")
if class_b_del<>"" and isnumeric(class_b_del) then 
'-===================| 删除大类 |=====================-
  sql="delete from "&db_table&"_type where id="&class_b_del   
  conn.execute(sql)  
  '删除相应小类
  sql_s ="delete from "&db_table&"_type where type_id="&class_b_del
  conn.execute(sql_s)
  '删除相应文章
  sql_art="delete from "&db_table&" where typeB_id="&class_b_del
  conn.execute(sql_art)
  del_Info="成功删除相应的大类、小类、文章"
  call backPage(del_info,"?",0)
end if

if class_s_del<>"" and isnumeric(class_s_del) then 
'-===================| 删除小类 |=====================-
  sql  ="delete from "&db_table&"_type where id="&class_s_del
  conn.execute(sql)
  sql_s="delete from "&db_table&" where typeS_id="&class_s_del
  conn.execute(sql_s)
  del_info="成功删除分类、相应的内容!"
  call backPage(del_info,"?",0)
end if





'################| 处理分类排序问题 |################
  order     =request.QueryString("order")
  id_b=request.QueryString("id_b")
  id_s=request.QueryString("id_s")
  
  if order<>"" and id_b<>"" and isnumeric(id_b) then
     if id_s<>"" and isnumeric(id_s) and id_b<>"" and isnumeric(id_b) then
	    order_sql="select * from "&db_table&"_type where id="&id_s
	 else
	    order_sql="select * from "&db_table&"_type where id="&id_b
	 end if
	 
	 row_now=conn.execute(order_sql)
     if not row_now.eof then
        row_now_order_id=row_now("order_id")
		row_now_id      =row_now("id")
	 else
		call backPage("参数有误!","?",0)
	 end if

	 '%%%%%%%%%%%%%%%%%%%%%%%
 	 if id_s<>"" and isnumeric(id_s) and id_b<>"" and isnumeric(id_b) then
    	orderStr=" and type_id="&id_b
	 else
    	orderStr=" and type_id=0"
 	 end if
	 
  end if


  if order="up" then
    '---------------------------------------------
set row_up=server.CreateObject("adodb.recordset")
    order_sql_up="select * from "&db_table&"_type where order_id<"&row_now_order_id&orderStr&" order by order_id desc"
    row_up.open order_sql_up,conn,1,1
	 if not row_up.eof then
        row_up_order_id=row_up("order_id")
		conn.execute("update "&db_table&"_type set order_id="&row_now_order_id&" where order_id="&row_up_order_id&orderStr)
		conn.execute("update "&db_table&"_type set order_id="&row_up_order_id&" where id="&row_now_id&orderStr)
     else
		call backPage("排序已到上限!","?",0)
	 end if
	row_up.close
set row_up=nothing

  elseif order="down" then
set row_down=server.CreateObject("adodb.recordset")
    order_sql_down="select * from "&db_table&"_type where order_id>"&row_now_order_id&orderStr&" order by order_id asc"
    row_down.open order_sql_down,conn,1,1
	 if not row_down.eof then
        row_down_order_id=row_down("order_id")
		row_down_query ="update "&db_table&"_type set order_id="&row_now_order_id&" where order_id="&row_down_order_id&orderStr
		row_now_query  ="update "&db_table&"_type set order_id="&row_down_order_id&" where id="&row_now_id&orderStr
		conn.execute(row_down_query)
		conn.execute(row_now_query)
	 else
		call backPage("排序已到下限!","?",0)
	 end if
	row_down.close
set row_down=nothing
	 
  end if
  
  
'################| 处理添加分类问题 |################
   id=request("id")
   edit=request("edit")
   
   title   = request.form("title")
   order_id= request.form("order_id")
   pic     = request.form("pic")
   type_id = request.form("type_id")
   
IF order_id="" or isNumeric(order_id)=false then order_id=0
IF type_id="" or isNumeric(type_id)=false then type_id=0

if edit<>"" then
   IF title=""  then 
      response.Write("<script language=javascript>alert('分类名称不能为空!');history.go(-1)</script>") 
      response.end()
   End IF

set rs=server.createobject("adodb.recordset")
 if edit="update" and id<>"" and isnumeric(id) then
    editStr="修改"
    sql="select * from "&db_table&"_type where id="&id 
    rs.open sql,conn,1,3
 else
    editStr="添加"
    sql="select * from "&db_table&"_type"
    rs.open sql,conn,1,3
    rs.addnew
 end if

'######### 写入数据 #############
    rs("title")   = title
    rs("order_id")= order_id
	rs("pic")     = pic
	
 if edit="add" then rs("type_id") =type_id
    session("type_id")=type_id

    rs.update
    rs.close
set rs=nothing
Response.Write "<script>window.location.href='article_type.asp';</script>" 
end if 
%>




<!--#include file="../top.asp"-->
<td valign="top" class="forumRow">


<%	
	set row_b=server.createobject("adodb.recordset") 
	    row_b_sql="select * from "&db_table&"_type where type_id=0 order by order_id asc" 
	    row_b.open row_b_sql,conn,1,3
	 if row_b.eof then
%>
暂无<%=db_title%>分类!
<%	
	 else
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="forMy forMAargin">
<form name="mclass_type_add" method="post" action="article_type.asp?edit=add">    
<tr <%=onTable%> >
            <td colspan="2" align="center" bgcolor="#279ED6" class="forumRow"><span class="mainTitle">分类管理</span>
              <input name="type_id" value="0" type="hidden">
</td>
            <td colspan="3" align="left" bgcolor="#279ED6" class="forumRow">
            <input name="title" type="text" id="title" style="width:100%" size="20">            </td>
            <td width="85" align="center" bgcolor="#279ED6" class="forumRow"><input style="width:100%" name="update22" type="submit" class="INPUTBottom" value="添加" id="update2" />
              <span class="forumRaw">
              <input name="FORM_TYPE" type="hidden" value="ADD" />
            </span></td>
</tr>
</form>

<tr <%=onTable%> >
            <td colspan="2" align="center" bgcolor="#279ED6" class="forumRaw" width="70"><span class="mainTitle">编号</span></td>
            <td align="left" bgcolor="#279ED6" class="forumRaw"><span class="mainTitle">&nbsp;<%=db_title%>分类名称</span></td>
            <td colspan="2" align="center" bgcolor="#279ED6" class="forumRaw" width="85"><span class="mainTitle">排序</span></td>
            <td width="85" align="center" bgcolor="#279ED6" class="forumRaw"><span class="mainTitle">修改操作</span></td>
          </tr>


<%
    do while not row_b.eof
	'#### 重写排序 ####
	r_b=r_b+1
	row_b("order_id")=r_b
	row_b.update()
%>
        <form name="article_type<%=row_b("id")%>" method="post" action="article_type.asp?edit=update&id=<%=row_b("id")%>">
         
<%if cstr(request("id"))=cstr(row_b("id")) then%>
<tr <%=onTable%> onDblClick="submit();" title="双击可完成编辑!">
<td align="center" class="forumRow"><%=row_b("id")%><a href="article_manage.asp?typeB_id=<%=row_b("id")%>">
              <input name="id" type="hidden" id="id" value="<%=row_b("id")%>" />
            </a></td>
           
            <td align="center" bgcolor="#E2E2C7" class="forumRow"><a href="?add_id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/add_class_s.gif" width="9" height="9" border="0" /></a></td>
            <td align="left" class="forumRow"><input name="title" type="text" class="input2" id="title" value="<%=row_b("title")%>" /></td>
		    <td width="85" align="center" class="forumRow"><input style="background-color:#F7F7EE; text-align:center" name="order_id" type="text" value="<%=row_b("order_id")%>" size="4"></td>
            <td width="40" align="center" class="forumRow">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/up_ico.gif" width="12" height="12" border="0"></a>

<a href="?order=down&id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/down_ico.gif" width="12" height="12"></a></td>
            <td width="85" align="center" class="forumRow"><input name="update2" type="submit" class="INPUTBottom" value="修改" id="input_but" />
            <input name="update3" type="button" class="INPUTBottom" id="update" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>'" /></td>
</tr>
<%else%>
<tr <%=onTable%> onDblClick="getEdit(<%=row_b("id")%>);" title="双击可编辑该分类!">
<td width="70" align="center" class="forumRow"><%=row_b("id")%><a href="article_manage.asp?typeB_id=<%=row_b("id")%>">
              <input name="id" type="hidden" id="id" value="<%=row_b("id")%>" />
            </a></td>
           
            <td width="20" align="center" bgcolor="#E2E2C7" class="forumRow"><a href="?add_id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/add_class_s.gif" width="9" height="9" border="0" /></a></td>
            <td align="left" class="forumRow">&nbsp;<%=row_b("title")%></td>
		    <td width="85" align="center" class="forumRow red" style="font-weight:bold"><%=row_b("order_id")%></td>
            <td align="center" class="forumRow">
<a href="?order=up&id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/up_ico.gif" width="12" height="12" border="0"></a>

<a href="?order=down&id_b=<%=row_b("id")%>"><img src="../../Edit/images/ico/down_ico.gif" width="12" height="12"></a>			</td>
            <td width="85" align="center" class="forumRow">
<%if db_types<>3 then%>
<input name="delete" type="button" class="INPUTBottom" onClick="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?class_b_del=<%=row_b("id")%>';}else{history.go(0);}" value="删除"/>
<%end if%>
<input name="update3" type="button" class="INPUTBottom" id="update" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>'" /></td>
</tr>
<%end if%> 
        </form>
<%
set row_s=server.CreateObject("ADODB.Recordset")
    sql1="select * from "&db_table&"_type where type_id="&row_b("id")&" order by order_id asc,id desc"
	row_s.open sql1,conn,1,3
	r_s=0
    do while not row_s.eof
   '#### 重写排序 ####
	r_s=r_s+1
	row_s("order_id")=r_s
	row_s.update()
		%>
		<form name="article_m_type<%=row_s("id")%>" method="post" action="?edit=update">
		


<%if cstr(request("id"))=cstr(row_s("id")) then%>
<tr <%=onTable%> onDblClick="submit();" title="双击可完成编辑!">
<td width="70" align="center" class="forumRow"><%=row_s("id")%><a href="article_manage.asp?typeS_id=<%=row_s("id")%>">
              <input name="id" type="hidden" id="id" value="<%=row_s("id")%>" />
            </a></td>
          
            <td width="20" align="center" class="forumRow"><img src="../../Edit/images/ico/type_ico.gif" /></td>
            <td align="left" class="forumRow"><input name="title" type="text" class="input2" id="title" value="<%=row_s("title")%>" /></td>
           <td width="85" align="center" class="forumRow"><input style="background-color:#FAFAF5;text-align:center" name="order_id" type="text" value="<%=row_s("order_id")%>" size="4"></td>
            <td align="center" class="forumRow">
<a href="?order=up&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>"><img src="../../Edit/images/ico/up_ico.gif" width="12" height="12"></a>

<a href="?order=down&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>"><img src="../../Edit/images/ico/down_ico.gif" width="12" height="12"></a>			</td>
            <td width="85" align="center" class="forumRow"><input name="update" type="submit" class="INPUTBottom" value="修改" id="input_but" />
            <input name="update23" type="button" class="INPUTBottom" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>&typeS_id=<%=row_s("id")%>'" /></td>
</tr>
<%else%>
<tr <%=onTable%> onDblClick="getEdit(<%=row_s("id")%>);" title="双击可编辑该分类!">
<td width="70" align="center" class="forumRow"><%=row_s("id")%><a href="article_manage.asp?typeS_id=<%=row_s("id")%>">
              <input name="id" type="hidden" id="id" value="<%=row_s("id")%>" />
            </a></td>
          
            <td width="20" align="center" class="forumRow"><img src="../../Edit/images/ico/type_ico.gif" /></td>
            <td align="left" class="forumRow">&nbsp;<%=row_s("title")%></td>
           <td width="85" align="center" class="forumRow"><%=row_s("order_id")%></td>
            <td align="center" class="forumRow">
<a href="?order=up&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>"><img src="../../Edit/images/ico/up_ico.gif" width="12" height="12"></a>

<a href="?order=down&id_b=<%=row_b("id")%>&id_s=<%=row_s("id")%>"><img src="../../Edit/images/ico/down_ico.gif" width="12" height="12"></a>			</td>
            <td width="85" align="center" class="forumRow">
<%if db_types<>3 then%>
<input name="delete" type="button" class="INPUTBottom" onClick="javascript:if(confirm('确定要删除此导航栏吗？删除后不可恢复!')){window.location.href='?class_s_del=<%=row_s("id")%>';}else{history.go(0);}" value="删除"/>
<%end if%>
<input name="update23" type="button" class="INPUTBottom" value="管理" onClick="window.location.href='article_manage.asp?typeB_id=<%=row_b("id")%>&typeS_id=<%=row_s("id")%>'" /></td>
</tr>
<%end if%> 
        </form>

<%
		row_s.movenext
		loop
		row_s.close
	set row_s=nothing
	
	
	row_b.movenext
	loop
	end if
	row_b.close
set row_b=nothing
%>
      </table>
<!--#include file="../bottom.asp"-->
</body>
</html>