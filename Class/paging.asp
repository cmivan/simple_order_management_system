<div class="paging_nav">
<%
'//// 分页 ///
num1=page-list_num
num2=page+list_num
if num1<1 then num1=1
if num2>last_page then num2=last_page
%>
<a hidefocus href="?page=1"><img src="edit/images/arrow_left.gif" align="absmiddle" /></a>
<%for num=num1 to num2%>
<%if int(num)=int(page) then%>
<a hidefocus href="?page=<%=num%>" class="on"><%=num%></a>
<%else%>
<a hidefocus href="?page=<%=num%>"><%=num%></a>
<%end if%>
<%
next
%>
<a hidefocus href="?page=<%=last_page%>"><img src="edit/images/arrow_right.gif" align="absmiddle" /></a>
</div>