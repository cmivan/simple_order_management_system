<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CodePage=65001%>
<%Response.Charset="UTF-8" %>
<link href="../css.css" rel="stylesheet" type="text/css" />
<!--#include file="../config.asp"-->
<!--#include file="../function.asp"-->
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
  ON ERROR RESUME NEXT
  DIM CONNS,CONNSTR,TIME1,TIME2,MDB,BackUP_MDB
      TIME1=TIMER
      MDB="../../edit/fckeditor.mdb"
      BackUP_MDB=year(now)&month(now)&day(now)&"."&session("LoginUserID")&".mdb"
      BackUP_MDB=BackUPpath&BackUP_MDB

      SET conn=SERVER.CREATEOBJECT("ADODB.CONNECTION")
          connSTR="DRIVER=MICROSOFT ACCESS DRIVER (*.MDB);DBQ="+SERVER.MAPPATH(MDB)
          conn.OPEN connSTR
'///////////////////////////////
  IF ERR THEN
     ERR.CLEAR
     SET conn = NOTHING
     call backPage(errConnStr&err.description,"javascript:;",0)
  END IF



'////开始生成任务目录 ///////////
 call cearPath()
%>






<%
'//// 退出登陆 //////////////////////
if session("LoginID")="" or isnumeric(session("LoginID"))=false then
   call backPage("登录失败,或登录超时!","../../index.asp",0)
end if

dim login
    login=lcase(request("login"))
'//// 退出登陆 //////////////////////
if login="out" then
   session.Abandon()
   call backPage("成功退出系统!","../../index.asp",0)
end if

'//// 设置下次登录的页面 //////////////////////
if login="this" then
   if session("LoginID")<>"" and isnumeric(session("LoginID")) then
      getThisUrl=GetLocationURL()
      if getThisUrl<>"" then
         getThisUrl=lcase(getThisUrl)
         getThisUrl=replace(getThisUrl,"?login=this","")
         getThisUrl=replace(getThisUrl,"&login=this","")
         conn.execute("update user set LoginToUrl='"&getThisUrl&"' where id="&session("LoginID"))
         errTip="2"
      else
         errTip="3"
      end if
   end if
end if
%>
