<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/allEpitem.asp?mallid=naverEP&menupos=<%=menupos%>">전체EP리스트</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/chgEPitem.asp?menupos=<%=menupos%>">요약EP리스트</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/potal/notinmakerid.asp?mallid=naverEP&menupos=<%=menupos%>">네이버EP 등록제외 브랜드</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/potal/notinitemid.asp?mallid=naverEP&menupos=<%=menupos%>">네이버EP 등록제외 상품</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/3depthmakerid.asp?menupos=<%=menupos%>">특정브랜드3Depth명 정의</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/3depthitemid.asp?menupos=<%=menupos%>">특정상품3Depth명 정의</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/eventName.asp?menupos=<%=menupos%>&mallid=nvshop">이벤트문구</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/chgsocname.asp?menupos=<%=menupos%>&mallid=nvshop">브랜드명 변경</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/diffItems.asp?menupos=<%=menupos%>">전일비교DATA</a>
    </td>
</tr>
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/couponMakerid.asp?mallid=naverEP&menupos=<%=menupos%>">네이버EP 쿠폰적용 브랜드</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/couponItem.asp?mallid=naverEP&menupos=<%=menupos%>">네이버EP 쿠폰적용 상품</a>
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
    <td width="11.1%">
    </td>
</tr>
</table>
<br>
<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td><a href="https://adcenter.shopping.naver.com/member/login/form.nhn;jsessionid=4FA352F7CD9516343C239A40D11A3EC5?targetUrl=%2Fproduct%2Fproduct_receive_status.nhn" target="_blank">지식쇼핑ADMIN</a> <font color='GREEN'>[ 10x10 | cube101010 ]</font> </td>
</tr>
</table>
<br>
<% End If %>
