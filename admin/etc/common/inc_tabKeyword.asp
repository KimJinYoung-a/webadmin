<%
Dim tabSelect
tabSelect = request("tabSelect")
If tabSelect = "" Then
	tabSelect = "1"
End If
%>
<script language='javascript'>
// ������� �귣��
function goTab(v){
	if (v == "1"){
		location.replace("/admin/etc/common/popKeywordList.asp?mallgubun=ssg");
	}else{
		location.replace("/admin/etc/common/popNotKeywordList.asp?mallgubun=ssg&tabSelect=2");
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td style="cursor:pointer;" width="50%" height="50" bgcolor="<%= Chkiif(tabSelect="1", "#BABAFF", "") %>" onclick="goTab('1');">Ű���� ������</td>
	<td style="cursor:pointer;" width="50%" height="50" bgcolor="<%= Chkiif(tabSelect="2", "#BABAFF", "") %>" onclick="goTab('2');">Ű���� ����</td>
</tr>
</table>