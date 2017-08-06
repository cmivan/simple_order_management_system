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
    title     =request.Form("title")
    content   =request.Form("content")
    type_id   =request.Form("type_id")
    order_id  =request.Form("order_id")

	toUrl     =request.Form("toUrl")
	note      =request.Form("note")
	add_data  =request.Form("add_data")
 if isdate(add_data)=false then add_data=now()

'——————审核\热门\最新\推荐
	ok  =request.Form("ok")
	hot =request.Form("hot")
	news=request.Form("news")
	tj  =request.Form("tj")


'——————排序
 if order_id="" or isnumeric(order_id)=false then order_id=0



'/////////////////////////////////
salerID        =cint(request.Form("salerID"))
jishuID        =cstr(request.Form("jishuID"))
User_Name      =cstr(request.Form("User_Name"))
User_Tel       =cstr(request.Form("User_Tel"))
User_Phone     =cstr(request.Form("User_Phone"))
User_Email     =cstr(request.Form("User_Email"))
User_QQ        =cstr(request.Form("User_QQ"))
User_SFZ       =cstr(request.Form("User_SFZ"))
User_BAH       =cstr(request.Form("User_BAH"))
User_Company   =cstr(request.Form("User_Company"))
User_Address   =cstr(request.Form("User_Address"))
User_OrderNote =cstr(request.Form("User_OrderNote"))
'-------------------------------------------
F_L_Name       =cstr(request.Form("F_L_Name"))
F_L_UserID     =cstr(request.Form("F_L_UserID"))
F_L_PassWord   =cstr(request.Form("F_L_PassWord"))
F_L_FtpUserID  =cstr(request.Form("F_L_FtpUserID"))
F_L_FtpPassWord=cstr(request.Form("F_L_FtpPassWord"))
F_L_SiteName   =cstr(request.Form("F_L_SiteName"))
F_L_Site       =cstr(request.Form("F_L_Site"))
F_L_Site_Pass  =cstr(request.Form("F_L_Site_Pass"))
F_L_Url        =cstr(request.Form("F_L_Url"))
F_L_Ip         =cstr(request.Form("F_L_Ip"))
F_L_Note       =cstr(request.Form("F_L_Note"))
F_L_Type       =cstr(request.Form("F_L_Type"))
F_L_OrderTime  =cstr(request.Form("F_L_OrderTime"))

'-------------------------------------------
web_BA=request.Form("web_BA")
web_WC=request.Form("web_WC")
web_SC=request.Form("web_SC")
web_SK=request.Form("web_SK")

if salerID=0 or isnumeric(salerID)=false then
   strMSG="请选择相应的销售!"
elseif User_Name="" then
   strMSG="请填写客户名称!"
elseif F_L_SiteName="" then
   strMSG="请填写网站名称!"
elseif isdate(F_L_OrderTime)=false then
   strMSG="请填写正确的签单时间!"
else

'///////////////////////////////
   if lcase(left(F_L_Site,4))="www." or lcase(left(F_L_Site,7))="http://" then
      strMSG="域名填写有误,填写示例: baidu.com !"
   end if


   if edit_id<>"" and isnumeric(edit_id) then
   set crs=conn.execute("SELECT * FROM " & db_table & " WHERE L_SiteName='"&F_L_SiteName&"' or (L_Site<>'' and L_Site='"&F_L_Site&"')")
       'if not crs.eof then strMSG="网站信息已经存在!"
   set crs=nothing
   end if

end if



'——————分类
  '//系判断是否为数组，再判断是否为数字(否则失败)
  type_ids=split(type_id,",")
  if ubound(type_ids)=1 then
     if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
	    typeB_id=type_ids(0)
	    typeS_id=type_ids(1)
	    session("type_id")=type_id   '记录本次操作分类
     else
        strMSG="分类有误!请重新选择." 
     end if
  else
     if type_id<>"" and isnumeric(type_id) then
	    typeB_id=type_id
	    typeS_id=null
	    session("type_id")=type_id   '记录本次操作分类
	 else
	    strMSG="分类有误!请重新选择." 
	 end if
  end if






'/// 返回提示
if strMSG<>"" then call backPage(strMSG,"?",2)

	rs("title")    =title
	rs("content")  =content
	rs("typeB_id") =typeB_id
	rs("typeS_id") =typeS_id

	rs("order_id") =order_id
	
	rs("ok")  =getBool(ok)
	rs("hot") =getBool(hot)
	rs("news")=getBool(news)
	rs("tj")  =getBool(tj)

	rs("toUrl") =toUrl
	rs("note")  =note
	rs("add_data")=add_data

'///////////////////////////////

'////临时的
if F_L_Url="" then F_L_Url="http://www."&F_L_Site&"/admin/"


rs("User_Name")      =User_Name
rs("User_Tel")       =User_Tel
rs("User_Phone")     =User_Phone
rs("User_Email")     =User_Email
rs("User_QQ")        =User_QQ
rs("User_SFZ")       =User_SFZ
rs("User_BAH")       =User_BAH
rs("User_Company")   =User_Company
rs("User_Address")   =User_Address
rs("User_OrderNote") =User_OrderNote
'--------------------------------
rs("salerID")=salerID
rs("jishuID")=jishuID
rs("L_UserID")=F_L_UserID
rs("L_PassWord")=F_L_PassWord
rs("L_FtpUserID")=F_L_FtpUserID
rs("L_FtpPassWord")=F_L_FtpPassWord
rs("L_SiteName")=F_L_SiteName
rs("L_Site")=F_L_Site
rs("L_Site_Pass")=F_L_Site_Pass
rs("L_Url")=F_L_Url
rs("L_Ip")=F_L_Ip
rs("L_Note")=F_L_Note
rs("L_Type")=F_L_Type
rs("L_OrderTime")=F_L_OrderTime
'--------------------------------
rs("web_BA")=web_BA
rs("web_WC")=web_WC
rs("web_SC")=web_SC
rs("web_SK")=web_SK

'--------------------------------
'// 记录添加者的ip，id
if EDIT_TYPE="ADD" then
   rs("insterID") =session("LoginID")
   rs("addIP")    =RealIp
end if
'// 记录编辑者的ip，id
rs("EditID")=session("LoginID")
rs("editIP")=RealIp

    end if
	rs.update
	rs.close
set rs=nothing

'/// 返回提示
 call backPage(editStr&"操作成功!","article_manage.asp"&getUrl("id",""),0)


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

	      order_id  =rs("order_id")

		  
		  ok   =rs("ok")
		  hot  =rs("hot")
		  news =rs("news")
		  tj   =rs("tj")
		  toUrl=rs("toUrl")
		  note =rs("note")
		  add_data=rs("add_data")
		  if order_id="" or isnumeric(order_id)=false then order_id=0


'//////////////////////////////////
          User_Name      =rs("User_Name")
          User_Tel       =rs("User_Tel")
          User_Phone     =rs("User_Phone")
          User_Email     =rs("User_Email")
		  User_QQ        =rs("User_QQ")
          User_SFZ       =rs("User_SFZ")
          User_BAH       =rs("User_BAH")
          User_Company   =rs("User_Company")
          User_Address   =rs("User_Address")
          User_OrderNote =rs("User_OrderNote")
          '------------------------------------
          salerID=rs("salerID")
          jishuID=rs("jishuID")
          F_L_UserID=rs("L_UserID")
          F_L_PassWord=rs("L_PassWord")
          F_L_FtpUserID=rs("L_FtpUserID")
          F_L_FtpPassWord=rs("L_FtpPassWord")
          F_L_SiteName=rs("L_SiteName")
          F_L_Site =rs("L_Site")
          F_L_Site_Pass=rs("L_Site_Pass")
          F_L_Url=rs("L_Url")
          F_L_Ip=rs("L_Ip")
          F_L_Note=rs("L_Note")
          F_L_Type=rs("L_Type")
          F_L_OrderTime=rs("L_OrderTime")
          '------------------------------------
          web_BA=rs("web_BA")
          web_WC=rs("web_WC")
          web_SC=rs("web_SC")
		  web_SK=rs("web_SK")



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
		window.addEvent('domready', function() { 
			myCal1 = new Calendar({ F_L_OrderTime: 'Y-m-d' }, { direction: 0, tweak: {x: 6, y: 0} });
		});
	</script>
<link rel="stylesheet" type="text/css" href="../../Edit/css/calendar.css" media="screen" />


<!--#include file="../top.asp"-->
<td valign="top" class="forumRow">
<form name="article_update" method="post" action="" onSubmit="tijiao.disabled='disabled';">
<table width="100%" border="0" align=center cellpadding=5 cellspacing=0 style="border:0px">
<tr>
<td class="forumRow">
<%
'////////////////////////////
if session("Loginpower")=1 or session("LoginTypeId")=38 then
   total=2
else
   total=1
end if
%>
<div class="tab_nav">
<a href="javascript:void(0);" onClick="showTab(1,1,<%=total%>);" id="title_1_1" class="on" style="font-weight: bold">网站客户信息</a>
<%if session("Loginpower")=1 or session("LoginTypeId")=38 then%>
<a href="javascript:void(0);" onClick="showTab(1,2,<%=total%>);" id="title_1_2" style="font-weight: bold">网站登录信息</a>
<%end if%>
</div>
<div>
<div id="tab_1_1">
<table width="100%" border="0" align=center cellpadding=0 cellspacing=0>
<TR>
<TD align="right" class="forumRow">销售：</TD>
<TD class="forumRow">
<select name="salerID" id="salerID">
  <option value="">请选择</option>
  <%
set rs=conn.execute("select * from user where typeB_id=14 order by UserName asc")
	do while not rs.eof
%>
  <%if int(salerID)=int(rs("id")) then%>
  <option value="<%=rs("id")%>" selected="selected" style="background-color:#666666; color:#FFFFFF"><%=rs("UserName")%></option>
  <%else%>
  <option value="<%=rs("id")%>"><%=rs("UserName")%></option>
  <%end if%>
  <%  
    rs.movenext
    loop
SET rs=nothing
%>
</select></TD>
</TR>
<TR>
<TD align="right" class="forumRow">客户称呼：</TD>
<TD class="forumRow">
  
  <input NAME=User_Name TYPE='text' class="input_1" id="User_Name" VALUE='<%=User_Name%>' />
  </TD></TR>
<TR>
<TD align="right" class="forumRow">固话：</TD>
<TD class="forumRow">
  
  <input NAME=User_Tel TYPE='text' class="input_1" id="User_Tel" VALUE='<%=User_Tel%>' />
  </TD></TR>
<TR>
<TD align="right" class="forumRow">手机：</TD>
<TD class="forumRow">
  
  <input NAME=User_Phone TYPE='text' class="input_1" id="User_Phone" VALUE='<%=User_Phone%>' />
  </TD></TR>
<TR>
<TD align="right" class="forumRow">Email：</TD>
<TD class="forumRow">
  
  <input NAME=User_Email TYPE='text' class="input_1" id="User_Email" VALUE='<%=User_Email%>' />
  </TD></TR>
  <TR>
<TD align="right" class="forumRow">联系QQ：</TD>
<TD class="forumRow">
  
  <input NAME=User_QQ TYPE='text' class="input_1" id="User_QQ" VALUE='<%=User_QQ%>' />
  </TD></TR>
<TR>
<TD align="right" class="forumRow">身份证号：</TD>
<TD class="forumRow">
  
  <input NAME=User_SFZ TYPE='text' class="input_1" id="User_SFZ" VALUE='<%=User_SFZ%>' />
  </TD></TR>
<TR>
<TD align="right" class="forumRow">公司名称：</TD>
<TD class="forumRow">
  
  <input NAME=User_Company TYPE='text' class="input_1" id="User_Company" VALUE='<%=User_Company%>' />
  </TD></TR>
<tr>
<TD width="100" align="right" class="forumRow">执照号/备案号：</TD>
<TD class="forumRow">
  
  <input NAME=User_BAH TYPE='text' class="input_1" id="User_BAH" VALUE='<%=User_BAH%>' />
  </TD></TR>
<tr>
<TD align="right" class="forumRow">公司地址：</TD>
<TD class="forumRow">
  
  <input NAME=User_Address TYPE='text' class="input_1" id="User_Address" VALUE='<%=User_Address%>' />
  </TD></TR>


<TR>
  <TD align="right" class="forumRow">
站名：</TD>
  <TD class="forumRow">
  <input NAME=F_L_SiteName TYPE='text' class="input_1" id="F_L_SiteName" style="width:200px;" VALUE='<%=F_L_SiteName%>' />
<!--#include file="../articles/articles_type.asp"-->
</TD>
</TR>
<TR>
<TD align="right" class="forumRow">域名：</TD>
<TD class="forumRow">
  
  <input NAME=F_L_Site TYPE='text' class="input_1" VALUE='<%=F_L_Site%>' onChange="ADMIN_EDIT_FORM.F_L_FtpUserID.value='webmaster@'+this.value;" />
  </TD></TR>

<TR>
<TD align="right" class="forumRow">签单时间：</TD>
<TD class="forumRow" style="float:left">
  <input NAME=F_L_OrderTime TYPE='text' class="input_1" id="F_L_OrderTime" VALUE='<%=F_L_OrderTime%>' readonly  style="float:left" /></TD>
</TR>

<tr>
<TD align="right" valign="top" class="forumRow">其他订单信息：</TD>
<TD class="forumRow">
  
<%  
    if User_Note<>"" then User_Note=replace(User_OrderNote,"'","&#39;")
	Set oFCKeditor = New FCKeditor 
	oFCKeditor.BasePath = "../../edit/"
	oFCKeditor.ToolbarSet = "Basic" 
	oFCKeditor.Width = "100%" 
	oFCKeditor.Height = "136" 
	oFCKeditor.Value = User_OrderNote
	oFCKeditor.Create "User_OrderNote"
    Set oFCKeditor = nothing
%>
 </TD></TR>
</table>
</div>

<%if session("Loginpower")=1 or session("LoginTypeId")=38 then%>
<div id="tab_1_2" style="display:none">
<table width="100%" border="0" align=center cellpadding=0 cellspacing=0>

<%if session("Loginpower")=1 or session("LoginTypeId")=38 then%>
<TR>
<TD width="100" align="right" class="forumRow">技术：</TD>
<TD class="forumRow"><select name="jishuID" id="jishuID">
  <option value="">请选择</option>
  <%
set rs=conn.execute("select * from user where typeB_id=38 order by UserName asc")
	do while not rs.eof
%>
  <%if int(jishuID)=int(rs("id")) then%>
  <option value="<%=rs("id")%>" selected="selected" style="background-color:#666666; color:#FFFFFF"><%=rs("UserName")%></option>
  <%else%>
  <option value="<%=rs("id")%>"><%=rs("UserName")%></option>
  <%end if%>
  <%  
    rs.movenext
    loop
SET rs=nothing
%>
</select></TD>
</TR>
<TR>
<TD align="right" class="forumRow">控制面板密码：</TD>
<TD class="forumRow">
  
  <input NAME=F_L_Site_Pass TYPE='text' class="input_1" VALUE='<%=F_L_Site_Pass%>' />
  </TD></TR>

<TR>
    <TD colspan="2" align="left" class="forumRow" style="font-weight: bold">&nbsp;&nbsp;[网站后台登陆]</TD>
    </TR>
<TR>
    <TD align="right" class="forumRow">
后台地址：</TD>
    <TD class="forumRow">
<%
'////临时的
if F_L_Url="" then F_L_Url="http://www."&F_L_Site&"/admin/"%>
  <input NAME=F_L_Url TYPE='text' class="input_1" VALUE='<%=F_L_Url%>' />
  </TD></TR>
	
<TR>
  <TD align="right" class="forumRow">
账号：</TD>
  <TD class="forumRow">
  
  <input NAME=F_L_UserID TYPE='text' class="input_1" VALUE='<%=F_L_UserID%>' />
  </TD>
  </TR><TR>
    <TD align="right" class="forumRow">
密码：</TD>
    <TD class="forumRow">
  <input name=F_L_PassWord type='text' class="input_1" VALUE='<%=F_L_PassWord%>' /></TD>
</TR>

<TR>
    <TD colspan="2" align="left" class="forumRow" style="font-weight: bold">&nbsp;&nbsp;[Ftp登陆信息]</TD>
    </TR>
<TR>
    <TD align="right" class="forumRow">
服务IP：</TD>
    <TD class="forumRow">
  <input NAME=F_L_Ip TYPE='text' class="input_1" VALUE='<%=F_L_Ip%>' />  </TD></TR>
<TR>
  <TD align="right" class="forumRow">
Ftp账号：</TD>
  <TD class="forumRow">
  
  <input NAME=F_L_FtpUserID TYPE='text' class="input_1" VALUE='<%=F_L_FtpUserID%>' />
  </TD>
  </TR><TR>
    <TD align="right" class="forumRow">
Ftp密码：</TD>
    <TD class="forumRow">
  <input name=F_L_FtpPassWord type='text' class="input_1" VALUE='<%=F_L_FtpPassWord%>' /></TD>
</TR>



<TR>
    <TD align="right" class="forumRow">是否已：</TD>
    <TD align="left" class="forumRow">
<label>
<input name="web_BA" type="checkbox" id="web_BA" value="1"<%if web_BA="1" then%> checked<%end if%> />
&nbsp;备案&nbsp;</label>&nbsp;&nbsp;
<label>
<input name="web_WC" type="checkbox" id="web_WC" value="1"<%if web_WC="1" then%> checked<%end if%> />
&nbsp;完成&nbsp;</label>&nbsp;&nbsp;
<label>
<input name="web_SC" type="checkbox" id="web_SC" value="1"<%if web_SC="1" then%> checked<%end if%> />
&nbsp;上传&nbsp;</label>
<label>
<input name="web_SK" type="checkbox" id="web_SK" value="1"<%if web_SK="1" then%> checked<%end if%> />
&nbsp;收款&nbsp;</label>
</TD>
</TR>
  <TR>
    <TD align="right" valign="top" class="forumRow">其他信息：</TD>
    <TD class="forumRow">
<%  
if F_L_Note<>"" then F_L_Note=replace(F_L_Note,"'","&#39;")

	Set oFCKeditor = New FCKeditor 
	oFCKeditor.BasePath = "../../edit/"
	oFCKeditor.ToolbarSet = "Basic" 
	oFCKeditor.Width = "100%" 
	oFCKeditor.Height = "148" 
	oFCKeditor.Value = F_L_Note
	oFCKeditor.Create "F_L_Note"
    Set oFCKeditor = nothing
%>
</TD>
  </TR>
<%end if%>
</table>
</div>
<%end if%>
</div>


<table width="100%" border="0" align=center cellpadding=0 cellspacing=0 style="margin-top:6px;">
<tr>
<td width="100" class="forumRow">&nbsp;</td>
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
</td>
</tr>
</table>
</form>

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
  <!--#include file="../bottom.asp"-->
</body>
</html>