<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
Dim sRect, mode
sRect = requestCheckVar(request("sRect"),32)
mode = requestCheckVar(request("mode"),32)
Dim sqlStr
Dim iRowsData
sqlStr = ""
sqlStr = sqlStr & " SELECT * FROM [db_temp].dbo.tbl_interpark_Tmp_DispCategory"
sqlStr = sqlStr & " WHERE dispyn='Y'"
If (sRect<>"") then
	sqlStr = sqlStr & " and dispcatename like '%" & sRect & "%'"
	sqlStr = sqlStr & " order by dispcatecode"
End If

If (sRect<>"") or (mode="all") then
	rsget.Open sqlStr,dbget,1
	If Not Rsget.Eof Then
		iRowsData = rsget.GetRows
	End If
	rsget.close
End If

Dim i,RowCnt
If IsArray(iRowsData) Then
	RowCnt = UBound(iRowsData,2)
Else
	RowCnt = -1
End If
%>
<script language='javascript'>
function CopyCode(comp){
	var compval = comp.value;
	var comptxt = comp[comp.selectedIndex].text;
	if (compval.length<1) { return };
	parent.frmSvr.interparkdispcategory.value = compval;
	parent.frmSvr.interparkdispcategoryText.value = comptxt;
}
</script>
<table border="1" cellspacing="1" cellpadding="1">
<tr>
	<td>
		<select name="dispcatecode" size="10" style="width:600px" onDblClick="CopyCode(this);">
		<% For i=0 to RowCnt %>
			<option value="<%= iRowsData(0,i) %>"><%= iRowsData(1,i) %>
		<% Next %>
		</select>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->