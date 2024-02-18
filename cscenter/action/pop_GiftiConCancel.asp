<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%

dim id
id = requestCheckVar(request("id"),10)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

''주문 마스타
dim ogifticonordermaster
set ogifticonordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    IF (ocsaslist.FOneItem.Frefminusorderserial<>"") then
        ogifticonordermaster.FRectOrderSerial = ocsaslist.FOneItem.Frefminusorderserial
    ELSE
        ogifticonordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    ENd IF

    ogifticonordermaster.QuickSearchOrderMaster
end if

if (ocsaslist.FResultCount<1) or (orefund.FResultCount<1) then
    response.write "<script>alert('환불내역이 없거나 유효하지 않은 내역입니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    response.write "<script>alert('접수 상태가 아닙니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'' 기프티콘 만 취소만 가능
if (IsNumeric(orefund.FOneItem.FpaygateTid)<>True) or orefund.FOneItem.Freturnmethod<>"R560" then
    response.write "<script>alert('기프티콘 만 취소 가능합니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
''rw ogifticonordermaster.FOneItem.FOrderSERIAL

if (ogifticonordermaster.FResultCount>0) then
    if (ogifticonordermaster.FOneItem.FCancelYn="N") and (ogifticonordermaster.FOneItem.Fjumundiv<>"9")  then
        response.write "<script>alert('반품주문 또는 주문이 취소된 경우만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if

dim i
dim IsDirectCancelAvail
IsDirectCancelAvail = True

dim CancelCase , etcCancelCase

CancelCase = "기프티콘 결제취소"
etcCancelCase = "기프티콘 결제취소"
%>
<script language='javascript'>
function ActCancel(frm){
    if (frm.msg.value.length<1){
        alert('취소사유를 입력해 주세요.');
        frm.msg.focus();
        return;
    }

    if (confirm('결제 취소 하시겠습니까?')){
        frm.action="pop_GiftiConCancel_process.asp";
        frm.submit();
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmCanncel" method="post" action="pop_giftcard_CardCancel_process.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="returnmethod" value="<%= orefund.FOneItem.Freturnmethod %>">
<% if (ogifticonordermaster.FResultCount>0) then %>
<input type="hidden" name="rdsite" value="">
<input type="hidden" name="buyemail" value="<%= ogifticonordermaster.FOneItem.Fbuyemail%>">
<% end if %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소자</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">주문번호</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FOrderSerial %>

        <% if (ocsaslist.FOneItem.Frefminusorderserial<>"") then %>
        (마이너스 주문번호 : <%= ocsaslist.FOneItem.Frefminusorderserial %>)
        <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">상태</td>
    <td bgcolor="#FFFFFF">
    <% if (ogifticonordermaster.FResultCount>0) then %>
        <font color="<%= ogifticonordermaster.FOneItem.CancelYnColor %>"><%= ogifticonordermaster.FOneItem.CancelYnName %></font> <font color="<%= ogifticonordermaster.FOneItem.IpkumDivColor %>"><%= ogifticonordermaster.FOneItem.GetJumunDivName %>

        <% if (ogifticonordermaster.FOneItem.Fjumundiv="9") then %>
        <font color=red><strong>[마이너스 주문]</strong></font>
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">구매자ID</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FUserID %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소방식</td>
    <td bgcolor="#FFFFFF">
        <%= orefund.FOneItem.FreturnmethodName %>
        <% if (orefund.FOneItem.Freturnmethod="R120") then %>
        (<strong><%= orefund.FOneItem.Freturnmethod %></strong>)
        <% else %>
		(<%= orefund.FOneItem.Freturnmethod %>)
		<% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소금액</td>
    <td bgcolor="#FFFFFF">
        <%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="tid" value="<%= orefund.FOneItem.FpaygateTid %>" size="60" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">취소사유</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="msg" value="<%= ChkIIF(IsDirectCancelAvail,CancelCase, etcCancelCase) %>" size="50" maxlength="60" >
    	<% if (ogifticonordermaster.FResultCount>0) then %>
    	<% if ((C_ADMIN_AUTH) and (ogifticonordermaster.FOneItem.Fjumundiv="9")) or (session("ssBctID")="icommang") or (session("ssBctID")="iroo4")  then %>
    	<input type="checkbox" name="force" >금액검토안함
    	<% end if %>
    	<% end if %>
    </td>
</tr>

<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <input type="button" class="button" value=" 결제 취소 " onClick="ActCancel(frmCanncel)">
    </td>
</tr>
</form>
</table>
<%
set ocsaslist = Nothing
set orefund = Nothing
set ogifticonordermaster = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->