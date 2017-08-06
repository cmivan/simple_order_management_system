<%
'///// 获取通知总数 //////
dim noticeNum
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
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<%if session("Loginpower")=1 then%>
	<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
      <tr>
       
        <td align="left"><form action="http://www.gzidc.com/login.php" method="post" target="_blank" style="margin:0">
            <input type="hidden" name="module" value="dologin" />
            <input type="hidden" name="goto_page" value="member.php" />
            <input name="login" type="hidden" value="walter0552@gmail.com"/>
            <input name="password" type="hidden" value="38218062"/>
            <input type="hidden" name="dologin" value="1" />
            <input name="submit" type="submit" class="INPUTBottom" value="新一代" style="padding-right:0; margin-right:0;width:100%" /></form></td>
        <td align="left"><form action="http://www.mofine.cn/admin/ynet_login.asp?ok" method="post" name="login_frm" target="_blank" id="login_frm">
            <input type="hidden" name="admin_log" value="ok" />
            <input type="hidden" name="username" value="zhicheng" />
            <input type="hidden" name="password" value="walter1982" />
            <input class="INPUTBottom" type="submit" value="魔方" name="enter.x" tabindex="4" style="padding-right:0; margin-right:0;width:100%" />
        </form></td>
       
      </tr>
    </table>
 <%end if%>
	
	</td>
  </tr>
  <tr>
    <td>
<div class="type_list">
<%if session("Loginpower")=1 then%>
<a hidefocus href="../../Class/user/article_manage.asp">部门信息</a>
<a hidefocus href="../../Class/order/article_manage.asp">订单管理</a>
<a hidefocus href="../../Class/order/article_edit.asp">添加订单</a>
<a hidefocus href="../../Class/notice/article_edit.asp">发布通知</a>
<a hidefocus href="../../Class/notice/article_manage.asp" style="border-bottom:#999999 3px solid;">通知管理</a>
<%end if%>
<a hidefocus href="../../Class/order/article_manage.asp?userID=<%=session("LoginID")%>" style="background-image:url(../../edit/images/rss.gif);" >我的订单</a>
<a hidefocus href="../../Class/notice/article_manage.asp?userID=<%=session("LoginID")%>" style="background-image:url(../../edit/images/<%=noticePic%>);" >通知消息<%=noticeNum%></a>
<%if session("LoginTypeId")=38 then%>
<a hidefocus href="../../Class/diary/article_edit.asp" style="background-image:url(../../edit/images/visited.gif);" >添加记录</a>
<a hidefocus href="../../Class/diary/article_manage.asp?userID=<%=session("LoginID")%>" style="background-image:url(../../edit/images/visited.gif);" >工作记录</a>
<%end if%>
<a hidefocus href="../../Class/user/article_edit.asp?id=<%=session("LoginID")%>" style="background-image:url(../../edit/images/visited.gif);" >我的信息</a>
<a hidefocus href="javascript:history.back(1);">上一步<img src="../../Edit/images/back.png" width="16" align="absmiddle" /></a>
<a hidefocus href="<%=getUrl("login","this")%>">登录到这</a>
<a hidefocus style="background-image:url(../../edit/images/out.gif);" href="?login=out">退出系统</a>
</div>
</td>
  </tr>
</table>

