<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->

<%
Dim taxIdx : taxIdx = requestCheckvar(request("taxIdx"),10)
Dim mode   : mode = requestCheckvar(request("mode"),10)
Dim neotaxno : neotaxno = requestCheckvar(request("neotaxno"),15)
Dim no_iss : no_iss = requestCheckvar(request("no_iss"),20)
Dim oTax

dim sqlStr,assignedRow
if (mode="act") then
    sqlStr = "update db_order.dbo.tbl_taxSheet"
    sqlStr = sqlStr & " set isueyn='Y'"
    sqlStr = sqlStr & " ,neotaxno='"&Trim(neotaxno)&"'"
    sqlStr = sqlStr & " ,no_iss='"&Trim(no_iss)&"'"
    sqlStr = sqlStr & " ,curUserId='"&session("ssBctId")&"'"
    sqlStr = sqlStr & " ,printdate=getdate()"
    sqlStr = sqlStr & " where  taxidx="&taxIdx
    sqlStr = sqlStr & " and delyn='N'"
    
    dbget.Execute sqlStr,assignedRow
    
    if (assignedRow>0) then
        response.write "<script>alert('저장 되었습니다.');opener.location.reload();window.close();</script>"
        response.end
    end if
end if

set oTax = new CTax
oTax.FRecttaxIdx = taxIdx

oTax.GetTaxRead
	
if (oTax.FTaxList(0).FisueYn = "Y") then
    response.write "<script>alert('이미 발행 완료된 세금계산서');</script>"
    response.end
end if
%>
<script language='javascript'>
function saveFrm(){
    var frm = document.frmAct;
    
    if (frm.neotaxno.value.length<1){
        alert('더존 TX 번호를 입력하세요.');
        return;
    }
    
    if (frm.no_iss.value.length<1){
        alert('국세청번호를 입력하세요.');
        return;
    }  
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a">
<form name="frmAct" method="post" action="">
<input type="hidden" name="mode" value="act">
<tr bgcolor="#FFFFFF">
    <td width="120" bgcolor="#DDDDFF">요청번호</td>
    <td><%= taxIdx %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="120" bgcolor="#DDDDFF">더존TX번호</td>
    <td><input type="text" name="neotaxno" value="" size="20" maxlength="15">
    (TX2012XXXXXXXXX)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="120" bgcolor="#DDDDFF">국세청번호</td>
    <td><input type="text" name="no_iss" value="" size="24" maxlength="24">
    (201202294100009XXXXXXXXX)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <input type="button" value="저장" onClick="saveFrm()">
    </td>
</tr>
</form>
</table>
<%
set oTax = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->