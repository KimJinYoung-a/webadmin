<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim mallid, vAction
mallid = request("mallid")
vAction 		= Request("action")

If vAction = "reg" Then
	Call Proc()
End If
%>
<script>
function frmSubmit(){
	if(confirm("���� �Ͻðڽ��ϱ�?")){
		document.frm.action.value = "reg";
		document.frm.submit();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="action" value="">
<input type="hidden" mallid="<%= mallid %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">API �׼�</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<select name="apiaction" class="select">
			<option value="CONFIRM">��ȸ</option>
			<option value="EDIT">����</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">Mall��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF" align="LEFT">
		<textarea name="mallgoodno" cols="20" rows="20"></textarea>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="2">
		<input type="button" class="button" value = "����" onclick="frmSubmit();">
	</td>
</tr>
</form>
</table>
<%
Function Proc()
	Dim strSql, vAction, vItemid, vMallGubun, vMallgoodno, vApiaction
	vAction		= Request("action")
	vMallGubun	= Request("mallid")
	vApiaction	= Request("apiaction")
	vMallgoodno	= Request("mallgoodno")
	If vMallgoodno <> "" then
		Dim iA2, arrTemp2, arrvMallgoodno
		vMallgoodno = replace(vMallgoodno,",",chr(10))
		vMallgoodno = replace(vMallgoodno,chr(13),"")
		arrTemp2 = Split(vMallgoodno,chr(10))
		iA2 = 0
		Do While iA2 <= ubound(arrTemp2)
			If Trim(arrTemp2(iA2))<>"" then
				arrvMallgoodno = arrvMallgoodno& "'"& trim(arrTemp2(iA2)) & "',"
			End If
			iA2 = iA2 + 1
		Loop
		vMallgoodno = left(arrvMallgoodno,len(arrvMallgoodno)-1)
	End If

	strSql = ""
	strSql = strSql & " SELECT itemid "
	strSql = strSql & " INTO #tmpTBL "
	strSql = strSql & " FROM db_item.dbo.tbl_cjmall_regitem "
	strSql = strSql & " WHERE cjmallprdno in ( "
	strSql = strSql & vMallgoodno
	strSql = strSql & " ) "
	dbget.Execute(strSql)

	strSql = ""
	strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_outmall_API_Que "
	strSql = strSql & " (mallid, apiAction, itemid, priority, lastUserid) "
	strSql = strSql & " SELECT 'cjmall', '"& vApiaction &"', itemid, 999, 'system' "
	strSql = strSql & " FROM #tmpTBL "
	dbget.Execute(strSql)

	strSql = ""
	strSql = strSql & " DROP TABLE #tmpTBL "
	dbget.Execute(strSql)
	Response.Write "<script>alert('ó���Ǿ����ϴ�.');location.href='popQueUserReg.asp?mallid=" & vMallGubun & "';</script>"
	Response.End
End Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
