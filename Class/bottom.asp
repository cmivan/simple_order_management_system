</td></TR>
<TR><td colspan="2" class="forumRow">
<%
select case errTip
       case "1"
            'call backPage("成功退出系统!","../../index.asp",0)
       case "2"
            call backPage("登录位置设置成功,下次登录后 会直接转到此页面!","",2) 
       case "3"
            call backPage("设置失败,暂时无法正确获取所在位置!","",2) 
end select
%>
</td></TR>
</TABLE>
<script>
function MM_Color(thisID,onType) {
if(onType=='on'){
thisID.className='forumRow_on';
}else{
thisID.className='forumRow_out';
}}
function ShowInfo(id){
window.location.href='?userID=<%=userID%>&typeB_id=<%=typeB_id%>&typeS_id=<%=typeS_id%>&keyword=<%=keyword%>&page=<%=page%>&LookID='+id;
}
function CloseInfo(id){
window.location.href='?userID=<%=userID%>&typeB_id=<%=typeB_id%>&typeS_id=<%=typeS_id%>&keyword=<%=keyword%>&page=<%=page%>&CloseID='+id;
}
</script>
<%call reSetSize()%>
<script language="JavaScript" src="../js.js"></script>