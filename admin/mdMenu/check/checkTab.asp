<script>
function goTab(v){
	if(v == 1){
		location.href='/admin/mdMenu/check/optAddpriceChecklist.asp?menupos=<%=menupos%>&vTab='+v;	
	}else if(v == 2){
		location.href='/admin/mdMenu/check/optUseChecklist.asp?menupos=<%=menupos%>&vTab='+v;
	}else if(v == 3){
		location.href='/admin/mdMenu/check/UpBeaSongErrList.asp?menupos=<%=menupos%>&vTab='+v;
	}else if(v == 4){
		location.href='/admin/mdMenu/check/deliverytypeErrList.asp?menupos=<%=menupos%>&vTab='+v;
	}else if(v == 5){
		location.href='/admin/mdMenu/check/buycashPrimeList.asp?menupos=<%=menupos%>&vTab='+v;
	}
}
</script>
<%
Dim vTab
vTab = requestCheckvar(request("vTab"),1)
If vTab = "" Then vTab = 1
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td width="20%" onclick="goTab(1);" style="cursor:pointer;" <%= chkiif(vTab=1, "bgcolor='#D2FFFF'", "") %> >�ɼ��߰��ݾ׸��԰�����</td>
	<td width="20%" onclick="goTab(2);" style="cursor:pointer;" <%= chkiif(vTab=2, "bgcolor='#D2FFFF'", "") %>>�ɼǻ�뿩�ο���</td>
	<td width="20%" onclick="goTab(3);" style="cursor:pointer;" <%= chkiif(vTab=3, "bgcolor='#D2FFFF'", "") %>>��ü���ǹ�� üũ</td>
	<td width="20%" onclick="goTab(4);" style="cursor:pointer;" <%= chkiif(vTab=4, "bgcolor='#D2FFFF'", "") %>>����&��۱��п���</td>
	<td width="20%" onclick="goTab(5);" style="cursor:pointer;" <%= chkiif(vTab=5, "bgcolor='#D2FFFF'", "") %>>��ǰ/�ɼǰ��ް��Ҽ���</td>
</tr>
</table>
<br>