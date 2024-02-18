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
	<td width="20%" onclick="goTab(1);" style="cursor:pointer;" <%= chkiif(vTab=1, "bgcolor='#D2FFFF'", "") %> >옵션추가금액매입가오류</td>
	<td width="20%" onclick="goTab(2);" style="cursor:pointer;" <%= chkiif(vTab=2, "bgcolor='#D2FFFF'", "") %>>옵션사용여부오류</td>
	<td width="20%" onclick="goTab(3);" style="cursor:pointer;" <%= chkiif(vTab=3, "bgcolor='#D2FFFF'", "") %>>업체조건배송 체크</td>
	<td width="20%" onclick="goTab(4);" style="cursor:pointer;" <%= chkiif(vTab=4, "bgcolor='#D2FFFF'", "") %>>매입&배송구분오류</td>
	<td width="20%" onclick="goTab(5);" style="cursor:pointer;" <%= chkiif(vTab=5, "bgcolor='#D2FFFF'", "") %>>상품/옵션공급가소수점</td>
</tr>
</table>
<br>