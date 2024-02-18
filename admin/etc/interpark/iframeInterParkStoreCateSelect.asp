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
sqlStr = sqlStr & " SELECT * FROM [db_temp].dbo.tbl_interpark_Tmp_StoreCategory"
sqlStr = sqlStr & " WHERE dispyn = 'Y'"
If (sRect<>"") Then
	sqlStr = sqlStr & " and storecatename like '%" & sRect & "%'"
	sqlStr = sqlStr & " order by storecatecode"
End If 

If (sRect<>"") or (mode="all") Then
	rsget.Open sqlStr,dbget,1
	If Not Rsget.Eof Then
		iRowsData = rsget.GetRows
	End if
	rsget.close
End If

Dim i,RowCnt
If IsArray(iRowsData) Then
	RowCnt = UBound(iRowsData,2)
Else
	RowCnt = -1
End If

Function getSupplyCtrtSeqName(iSupplyCtrtSeq)
	If IsNULL(iSupplyCtrtSeq) Then Exit Function
	
	If (iSupplyCtrtSeq=2) Then
		getSupplyCtrtSeqName = "리빙"
	ElseIf (iSupplyCtrtSeq=3) Then
		getSupplyCtrtSeqName = "잡화"
	ElseIf (iSupplyCtrtSeq=4) Then
		getSupplyCtrtSeqName = "의류"
	End If
End Function
%>
<script language='javascript'>
function CopyCode(comp){
	var compval = comp.value;
	var compid  = comp[comp.selectedIndex].id;
	var comptxt = comp[comp.selectedIndex].text;
	if (compval.length<1) { return };
	var CtrtSeqName = '';
	if (compid=="2") {
		CtrtSeqName = "리빙";
	}else if (compid=="3") {
		CtrtSeqName = "잡화";
	}else if (compid=="4") {
		CtrtSeqName = "의류";
	}
	parent.frmSvr.SupplyCtrtSeq.value = compid;
	parent.frmSvr.SupplyCtrtSeqName.value = CtrtSeqName;
	parent.frmSvr.interparkstorecategory.value = compval;
	parent.frmSvr.interparkstorecategoryText.value = comptxt;
}
</script>
<table border="1" cellspacing="1" cellpadding="1">
<tr>
	<td>
		<select name="dispcatecode" size="10" style="width:600px" onDblClick="CopyCode(this);">
		<% For i=0 to RowCnt %>
			<option id="<%= iRowsData(0,i) %>" value="<%= iRowsData(1,i) %>"><%= getSupplyCtrtSeqName(iRowsData(0,i)) %>&nbsp;&nbsp;|&nbsp;&nbsp;<%= iRowsData(2,i) %>
		<% Next %>
		</select>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->