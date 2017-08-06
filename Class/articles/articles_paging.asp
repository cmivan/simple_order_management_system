<div class="paging_nav">
<%
'//// 分页 ///
NUM1=PAGE-LIST_NUM
NUM2=PAGE+LIST_NUM
IF NUM1<1 THEN NUM1=1
IF NUM2>LAST_PAGE THEN NUM2=LAST_PAGE
%>
<a hidefocus href="<%=getUrl("page",1)%>"><img src="../../edit/images/arrow_left.gif" align="absmiddle" /></a>
<%

FOR NUM=NUM1 TO NUM2
%>
<%if int(NUM)=int(PAGE) then%>
<a hidefocus href="<%=getUrl("page",num)%>" class="on"><%=NUM%></a>
<%else%>
<a hidefocus href="<%=getUrl("page",num)%>"><%=NUM%></a>
<%end if%>
<%
NEXT
%>
<a hidefocus href="<%=getUrl("page",LAST_PAGE)%>"><img src="../../edit/images/arrow_right.gif" align="absmiddle" /></a>
</div>