<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
Dim didx, mwgubun, mode, chgmwgubun
didx  = requestCheckvar(request("didx"),10)
mode  = requestCheckvar(request("mode"),10)
chgmwgubun = requestCheckvar(request("chgmwgubun"),10)

Dim sqlStr, ArrList, i, AssignedRow

IF (mode="act") then
    sqlStr = "  update [db_storage].[dbo].tbl_acount_storage_detail"& VbCRLF
    sqlStr = sqlStr & " set mwgubun='"&chgmwgubun&"'"& VbCRLF
    sqlStr = sqlStr & " where id="&didx
    
    dbget.Execute sqlStr,AssignedRow

    IF (AssignedRow>0) then
        response.write "<script>alert('저장되었습니다.');opener.location.reload();window.close()</script>"
        dbget.close() : response.end
    end if
END IF

sqlStr = " select d.mwgubun "
sqlStr = sqlStr & " from [db_storage].[dbo].tbl_acount_storage_detail d"
sqlStr = sqlStr & " where id="&didx

rsget.Open sqlStr,dbget,1
if  not rsget.EOF  then
	mwgubun = rsget("mwgubun")
end if
rsget.Close


if isNULL(mwgubun) then mwgubun=""

function getchulgoMwgubunName(imwgubun)
    if (imwgubun="") then
        getchulgoMwgubunName ="미지정"
    elseif (imwgubun="C") then
        getchulgoMwgubunName ="위탁=>출고매입"
    elseif (imwgubun="M") then
        getchulgoMwgubunName ="온라인매입"
    elseif (imwgubun="F") then
        getchulgoMwgubunName ="오프매입"
    elseif (imwgubun="W") then
        getchulgoMwgubunName ="출고위탁"
    elseif (imwgubun="X") then
        getchulgoMwgubunName ="위탁=>위탁"
    else
        getchulgoMwgubunName ="??"
    end if
end function

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;
    
    if (frm.chgmwgubun.value.length<1){
        if (!confirm('출고 구분이 지정되지 않았습니다. 계속하시겠습니까?')) return;
    }

    if (confirm('저장 하시겠습니까?')){
        frm.mode.value="act";
        frm.submit();
    }
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmAct" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="didx" value="<%=didx%>">
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >현재 </td>
	<td ><%= getchulgoMwgubunName(mwgubun) %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >변경 </td>
    <td >
        <select name="chgmwgubun">
        <option value=""> 미지정
        <option value="C" <%=CHKIIF(mwgubun="C","selected","")%> >(C) 위탁=>출고매입
        <option value="M" <%=CHKIIF(mwgubun="M","selected","")%> >(M) 온라인매입
        <option value="F" <%=CHKIIF(mwgubun="F","selected","")%> >(F) 오프매입
        <option value="W" <%=CHKIIF(mwgubun="W","selected","")%> >(W) 출고위탁
        <option value="X" <%=CHKIIF(mwgubun="X","selected","")%> >(X) 위탁=>위탁
        </select>
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="저장" onClick="saveThis()">
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
