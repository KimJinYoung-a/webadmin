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

''�ֹ� ����Ÿ
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
    response.write "<script>alert('ȯ�ҳ����� ���ų� ��ȿ���� ���� �����Դϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    response.write "<script>alert('���� ���°� �ƴմϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'' ����Ƽ�� �� ��Ҹ� ����
if (IsNumeric(orefund.FOneItem.FpaygateTid)<>True) or orefund.FOneItem.Freturnmethod<>"R560" then
    response.write "<script>alert('����Ƽ�� �� ��� �����մϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
''rw ogifticonordermaster.FOneItem.FOrderSERIAL

if (ogifticonordermaster.FResultCount>0) then
    if (ogifticonordermaster.FOneItem.FCancelYn="N") and (ogifticonordermaster.FOneItem.Fjumundiv<>"9")  then
        response.write "<script>alert('��ǰ�ֹ� �Ǵ� �ֹ��� ��ҵ� ��츸 ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if

dim i
dim IsDirectCancelAvail
IsDirectCancelAvail = True

dim CancelCase , etcCancelCase

CancelCase = "����Ƽ�� �������"
etcCancelCase = "����Ƽ�� �������"
%>
<script language='javascript'>
function ActCancel(frm){
    if (frm.msg.value.length<1){
        alert('��һ����� �Է��� �ּ���.');
        frm.msg.focus();
        return;
    }

    if (confirm('���� ��� �Ͻðڽ��ϱ�?')){
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
    <td width="90" bgcolor="<%= adminColor("topbar") %>">�����</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">�ֹ���ȣ</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FOrderSerial %>

        <% if (ocsaslist.FOneItem.Frefminusorderserial<>"") then %>
        (���̳ʽ� �ֹ���ȣ : <%= ocsaslist.FOneItem.Frefminusorderserial %>)
        <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">����</td>
    <td bgcolor="#FFFFFF">
    <% if (ogifticonordermaster.FResultCount>0) then %>
        <font color="<%= ogifticonordermaster.FOneItem.CancelYnColor %>"><%= ogifticonordermaster.FOneItem.CancelYnName %></font> <font color="<%= ogifticonordermaster.FOneItem.IpkumDivColor %>"><%= ogifticonordermaster.FOneItem.GetJumunDivName %>

        <% if (ogifticonordermaster.FOneItem.Fjumundiv="9") then %>
        <font color=red><strong>[���̳ʽ� �ֹ�]</strong></font>
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">������ID</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FUserID %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">��ҹ��</td>
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
    <td width="90" bgcolor="<%= adminColor("topbar") %>">��ұݾ�</td>
    <td bgcolor="#FFFFFF">
        <%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">PG�� ID</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="tid" value="<%= orefund.FOneItem.FpaygateTid %>" size="60" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">��һ���</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="msg" value="<%= ChkIIF(IsDirectCancelAvail,CancelCase, etcCancelCase) %>" size="50" maxlength="60" >
    	<% if (ogifticonordermaster.FResultCount>0) then %>
    	<% if ((C_ADMIN_AUTH) and (ogifticonordermaster.FOneItem.Fjumundiv="9")) or (session("ssBctID")="icommang") or (session("ssBctID")="iroo4")  then %>
    	<input type="checkbox" name="force" >�ݾװ������
    	<% end if %>
    	<% end if %>
    </td>
</tr>

<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <input type="button" class="button" value=" ���� ��� " onClick="ActCancel(frmCanncel)">
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