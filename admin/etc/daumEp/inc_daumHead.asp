<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="20%">
    	<a href="allEPitem.asp?menupos=<%=menupos%>">��üEP����Ʈ</a>
    </td>
    <td width="20%">
    	<a href="best100EPitem.asp?menupos=<%=menupos%>">����Ʈ100EP����Ʈ</a>
    </td>
    <td width="20%">
    	<a href="notinmakerid.asp?menupos=<%=menupos%>">����EP ������� �귣��</a>
    </td>
    <td width="20%">
    	<a href="notinitemid.asp?menupos=<%=menupos%>">����EP ������� ��ǰ</a>
    </td>
    <td width="20%">
    	<a href="3depthmakerid.asp?menupos=<%=menupos%>">Ư���귣��3Depth�� ����</a>
    </td>
</tr>
</table>
<br>
<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td><a href="https://adcenter.shopping.naver.com/member/login/form.nhn;jsessionid=4FA352F7CD9516343C239A40D11A3EC5?targetUrl=%2Fproduct%2Fproduct_receive_status.nhn" target="_blank">���ļ���ADMIN</a> <font color='GREEN'>[ 10x10 | cube101010 ]</font> </td>
</tr>
</table>
<br>
-->
<% End If %>
