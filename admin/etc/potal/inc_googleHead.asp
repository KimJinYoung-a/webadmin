<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="33%">
    	<a href="allFeeditem.asp?mallid=ggshop&menupos=<%=menupos%>">��üFeed����Ʈ</a>
    </td>
    <td width="33%">
    	<a href="notinmakerid.asp?mallid=ggshop&menupos=<%=menupos%>">���ۼ��� ������� �귣��</a>
    </td>
    <td width="33%">
    	<a href="notinitemid.asp?mallid=ggshop&menupos=<%=menupos%>">���ۼ��� ������� ��ǰ</a>
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
