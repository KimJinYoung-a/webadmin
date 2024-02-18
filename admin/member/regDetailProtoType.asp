<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->
<%
dim ContractType, detailKey, detailDesc, mode
ContractType = request("ContractType")
detailKey    = request("detailKey")
detailDesc   = request("detailDesc")
mode         = request("mode")

dim sqlStr

if (mode="regmode") then
    sqlStr =" insert into db_partner.dbo.tbl_partner_contractDetailType"
    sqlStr = sqlStr & " (ContractType, detailKey, detailDesc) "
    sqlStr = sqlStr & " values(" & ContractType 
    sqlStr = sqlStr & " ,'" & detailKey & "'"
    sqlStr = sqlStr & " ,'" & html2DB(detailDesc) & "'"
    sqlStr = sqlStr & " )" 
    
    dbget.Execute sqlStr
elseif (mode="editmode") then
    sqlStr = " update db_partner.dbo.tbl_partner_contractDetailType"
    sqlStr = sqlStr & " set detailDesc='" & html2DB(detailDesc) & "'" & VbCrlf
    sqlStr = sqlStr & " where ContractType=" & ContractType & VbCrlf
    sqlStr = sqlStr & " and detailKey='" & detailKey & "'"
    
    dbget.Execute sqlStr
end if

dim onecontractProtoType 
set onecontractProtoType = new CPartnerContract
onecontractProtoType.FRectContractType = ContractType
onecontractProtoType.getOneContractProtoType

dim onecontractDetailProtoType 
set onecontractDetailProtoType = new CPartnerContract
onecontractDetailProtoType.FRectContractType = ContractType
onecontractDetailProtoType.FRectdetailKey    = detailKey
onecontractDetailProtoType.getOneContractDetailProtoType


%>
<script language='javascript'>
function ConFirmNsubmit(frm){
    if (frm.detailKey.value.length<6){
        alert('변수 Key 값은 최소 6자 이상입니다.');
        frm.detailKey.focus();
        return;
    }
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
<form name="frmSvr" method="post" action="">
<% if onecontractDetailProtoType.FResultCount>0 then %>
<input type="hidden" name="mode" value="editmode">
<% else %>
<input type="hidden" name="mode" value="regmode">
<% end if %>

<tr bgcolor="#DDDDFF">
    <td width="100" >계약서 명</td>
    <td bgcolor="#FFFFFF"><%= onecontractProtoType.FOneItem.FcontractName %></td>
</tr>
<tr bgcolor="#DDDDFF">
    <td >변수 Type(KEY)</td>
    <td bgcolor="#FFFFFF">
        <% if onecontractDetailProtoType.FResultCount>0 then %>
        <input type="text" name="detailKey" value="<%= onecontractDetailProtoType.FOneItem.FdetailKey %>" size="40" maxlength="40" readOnly >
        <% else %>
        <input type="text" name="detailKey" value="" size="40" maxlength="40">
        <% end if %>
    </td>
</tr>
<tr bgcolor="#DDDDFF">
    <td >변수 설명</td>
    <td bgcolor="#FFFFFF">
        <% if onecontractDetailProtoType.FResultCount>0 then %>
        <input type="text" name="detailDesc" value="<%= onecontractDetailProtoType.FOneItem.FdetailDesc %>" size="40" maxlength="40">
        <% else %>
        <input type="text" name="detailDesc" value="" size="40" maxlength="40">
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
    <td align="center" colspan="2"><input type="button" value="저장" onClick="ConFirmNsubmit(frmSvr);"></td>
</tr>
</form>
</table>

<%
set onecontractProtoType = Nothing
set onecontractDetailProtoType = Nothing
%>

<% if mode<>"" then %>
<script language='javascript'>
opener.location.reload();
</script>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->