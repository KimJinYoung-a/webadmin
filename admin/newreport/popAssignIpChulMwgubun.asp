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
        response.write "<script>alert('����Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
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
        getchulgoMwgubunName ="������"
    elseif (imwgubun="C") then
        getchulgoMwgubunName ="��Ź=>������"
    elseif (imwgubun="M") then
        getchulgoMwgubunName ="�¶��θ���"
    elseif (imwgubun="F") then
        getchulgoMwgubunName ="��������"
    elseif (imwgubun="W") then
        getchulgoMwgubunName ="�����Ź"
    elseif (imwgubun="X") then
        getchulgoMwgubunName ="��Ź=>��Ź"
    else
        getchulgoMwgubunName ="??"
    end if
end function

%>
<script language='javascript'>
function saveThis(){
    var frm = document.frmAct;
    
    if (frm.chgmwgubun.value.length<1){
        if (!confirm('��� ������ �������� �ʾҽ��ϴ�. ����Ͻðڽ��ϱ�?')) return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
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
    <td width="80" bgcolor="#F3F3FF" >���� </td>
	<td ><%= getchulgoMwgubunName(mwgubun) %></td>
</tr>
<tr align="center"  bgcolor="#FFFFFF" height="20">
    <td width="80" bgcolor="#F3F3FF" >���� </td>
    <td >
        <select name="chgmwgubun">
        <option value=""> ������
        <option value="C" <%=CHKIIF(mwgubun="C","selected","")%> >(C) ��Ź=>������
        <option value="M" <%=CHKIIF(mwgubun="M","selected","")%> >(M) �¶��θ���
        <option value="F" <%=CHKIIF(mwgubun="F","selected","")%> >(F) ��������
        <option value="W" <%=CHKIIF(mwgubun="W","selected","")%> >(W) �����Ź
        <option value="X" <%=CHKIIF(mwgubun="X","selected","")%> >(X) ��Ź=>��Ź
        </select>
    </td>
</tr>

<tr bgcolor="#FFFFFF" height="40">
    <td colspan="7" align="center">
    <input type="button" class="button" value="����" onClick="saveThis()">
    </td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
