<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ�
' History : 2009.04.07 ������ ����
'			2010.05.12 �ѿ�� ����
'####################################################
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
	Response.CharSet = "euc-kr" 
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<%
Dim code_large , code_mid
	code_large = RequestCheckvar(request("code_large"),3)
	code_mid = RequestCheckvar(request("code_mid"),3)
%>
<script src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
// ī�װ� ����� ���
function changecontent(){
		location.href = "?code_large=" + poplecfrm.code_large.value + "<%=CHKIIF(code_large<>"","&code_mid="&chr(34)&" + poplecfrm.code_mid.value + "&chr(34)&"","")%>";
}
function checkval(){
	
	var code_large = $("#code_large").attr("value");
	var code_large_nm = $("#code_large option:selected").text();
	var code_mid = $("#code_mid").attr("value");
	var code_mid_nm = $("#code_mid option:selected").text();

	var frm = opener.lecfrm;

	if (code_mid == "" || code_mid_nm == "")
	{
		alert("��ī�װ��� ���� �ϼ���");
		return false;
	} else {
		frm.code_large.value = code_large;
		frm.code_mid.value = code_mid;
		frm.large_name.value = code_large_nm;
		frm.mid_name.value = code_mid_nm;

		self.close();
	}
}
</script>
<table class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="poplecfrm" method="get">
	<tr bgcolor="#DDDDFF">
		<td>���� ī�װ� ����</td>
		<td>
			<% DrawSelectBoxLecCategoryLarge "code_large" ,  code_large , "Y" %>
			<% 
				if code_large <> "" Then
				  response.write "&nbsp;"
				  Call DrawSelectBoxLecCategoryMid("code_mid", code_large , code_mid  , "N")
				End If 
			%>
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" align="right">
		<td colspan="2"><input type="button" value="�Է�" onclick="checkval()"></td>
	</tr>
	</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->