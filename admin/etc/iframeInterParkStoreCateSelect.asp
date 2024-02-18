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
sqlStr = "select * from [db_temp].dbo.tbl_interpark_Tmp_StoreCategory"
sqlStr = sqlStr + " where dispyn='Y'"
if (sRect<>"") then
    sqlStr = sqlStr + " and storecatename like '%" + sRect + "%'"
    sqlStr = sqlStr + " order by storecatecode"
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

function getSupplyCtrtSeqName(iSupplyCtrtSeq)
    if IsNULL(iSupplyCtrtSeq) then Exit Function
    
    if (iSupplyCtrtSeq=2) then  
        getSupplyCtrtSeqName = "리빙"
    elseif (iSupplyCtrtSeq=3) then  
        getSupplyCtrtSeqName = "잡화"    
    elseif (iSupplyCtrtSeq=4) then  
        getSupplyCtrtSeqName = "의류"  
    end if
end function
    
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
    <% for i=0 to RowCnt %>
    <option id="<%= iRowsData(0,i) %>" value="<%= iRowsData(1,i) %>"><%= getSupplyCtrtSeqName(iRowsData(0,i)) %>&nbsp;&nbsp;|&nbsp;&nbsp;<%= iRowsData(2,i) %>
    <% next %>
    </select>    
    </td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->