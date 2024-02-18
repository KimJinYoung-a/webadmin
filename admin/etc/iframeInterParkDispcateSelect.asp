<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->
<%
dim sRect, mode
sRect = requestCheckVar(request("sRect"),32)
mode = requestCheckVar(request("mode"),32)

dim sqlStr
dim iRowsData
sqlStr = "select * from [db_temp].dbo.tbl_interpark_Tmp_DispCategory"
sqlStr = sqlStr + " where dispyn='Y'"
if (sRect<>"") then
    sqlStr = sqlStr + " and dispcatename like '%" + sRect + "%'"
    sqlStr = sqlStr + " order by dispcatecode"
end if 

if (sRect<>"") or (mode="all") then
rsget.Open sqlStr,dbget,1
if Not Rsget.Eof then
    iRowsData = rsget.GetRows
end if
rsget.close
End if

dim i,RowCnt

IF IsArray(iRowsData) then
    RowCnt = UBound(iRowsData,2)
else
    RowCnt = -1
End if
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
    <% for i=0 to RowCnt %>
    <option value="<%= iRowsData(0,i) %>"><%= iRowsData(1,i) %>
    <% next %>
    </select>    
    </td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->