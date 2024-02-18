<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="20%">
    	<a href="allEPitem.asp?menupos=<%=menupos%>">전체EP리스트</a>
    </td>
    <td width="20%">
    	<a href="best100EPitem.asp?menupos=<%=menupos%>">베스트100EP리스트</a>
    </td>
    <td width="20%">
    	<a href="notinmakerid.asp?menupos=<%=menupos%>">다음EP 등록제외 브랜드</a>
    </td>
    <td width="20%">
    	<a href="notinitemid.asp?menupos=<%=menupos%>">다음EP 등록제외 상품</a>
    </td>
    <td width="20%">
    	<a href="3depthmakerid.asp?menupos=<%=menupos%>">특정브랜드3Depth명 정의</a>
    </td>
</tr>
</table>
<br>
<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td><a href="https://adcenter.shopping.naver.com/member/login/form.nhn;jsessionid=4FA352F7CD9516343C239A40D11A3EC5?targetUrl=%2Fproduct%2Fproduct_receive_status.nhn" target="_blank">지식쇼핑ADMIN</a> <font color='GREEN'>[ 10x10 | cube101010 ]</font> </td>
</tr>
</table>
<br>
-->
<% End If %>
