<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/allEpitem.asp?mallid=naverEP&menupos=<%=menupos%>">��üEP����Ʈ</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/chgEPitem.asp?menupos=<%=menupos%>">���EP����Ʈ</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/potal/notinmakerid.asp?mallid=naverEP&menupos=<%=menupos%>">���̹�EP ������� �귣��</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/potal/notinitemid.asp?mallid=naverEP&menupos=<%=menupos%>">���̹�EP ������� ��ǰ</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/3depthmakerid.asp?menupos=<%=menupos%>">Ư���귣��3Depth�� ����</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/3depthitemid.asp?menupos=<%=menupos%>">Ư����ǰ3Depth�� ����</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/eventName.asp?menupos=<%=menupos%>&mallid=nvshop">�̺�Ʈ����</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/chgsocname.asp?menupos=<%=menupos%>&mallid=nvshop">�귣��� ����</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/diffItems.asp?menupos=<%=menupos%>">���Ϻ�DATA</a>
    </td>
</tr>
<tr align="center" bgcolor="#f1f1f1" height="30">
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/couponMakerid.asp?mallid=naverEP&menupos=<%=menupos%>">���̹�EP �������� �귣��</a>
    </td>
    <td width="11.1%">
    	<a href="/admin/etc/naverEP/couponItem.asp?mallid=naverEP&menupos=<%=menupos%>">���̹�EP �������� ��ǰ</a>
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
    <td><a href="https://adcenter.shopping.naver.com/member/login/form.nhn;jsessionid=4FA352F7CD9516343C239A40D11A3EC5?targetUrl=%2Fproduct%2Fproduct_receive_status.nhn" target="_blank">���ļ���ADMIN</a> <font color='GREEN'>[ 10x10 | cube101010 ]</font> </td>
</tr>
</table>
<br>
<% End If %>
