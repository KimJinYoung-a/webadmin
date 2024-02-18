<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="33%">
    	<a href="allFeeditem.asp?mallid=ggshop&menupos=<%=menupos%>">전체Feed리스트</a>
    </td>
    <td width="33%">
    	<a href="notinmakerid.asp?mallid=ggshop&menupos=<%=menupos%>">구글쇼핑 등록제외 브랜드</a>
    </td>
    <td width="33%">
    	<a href="notinitemid.asp?mallid=ggshop&menupos=<%=menupos%>">구글쇼핑 등록제외 상품</a>
    </td>
</tr>
</table>
<br>
<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td><a href="https://merchants.google.com/" target="_blank">Google Merchant Center</a> <font color='GREEN'>[ tenbytencorp@gmail.com | cube1010?? ]</font> </td>
</tr>
</table>
<br>
<% End If %>
