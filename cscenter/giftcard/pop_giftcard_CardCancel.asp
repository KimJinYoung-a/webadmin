<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->

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
dim ogiftcardordermaster
set ogiftcardordermaster = new cGiftCardOrder

if (ocsaslist.FResultCount>0) then
	ogiftcardordermaster.FRectgiftorderserial = ocsaslist.FOneItem.Forderserial

    ogiftcardordermaster.getCSGiftcardOrderDetail
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

'' �ſ�ī�� ��Ҹ� ����
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('���� �ſ�ī�� �ŷ��� ��� �����մϴ�.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

'' IniPay �� ��Ҹ� ����  Stdpay, INIMX_
if (Left(orefund.FOneItem.FpaygateTid,10)<>"IniTechPG_") AND Left(orefund.FOneItem.FpaygateTid,6)<>"Stdpay" AND Left(orefund.FOneItem.FpaygateTid,6)<>"INIMX_" AND orefund.FOneItem.Freturnmethod<>"R400" then
    response.write "<script>alert('�̴Ͻý� �ŷ��� ��� �����մϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
''rw ogiftcardordermaster.FOneItem.FOrderSERIAL

if (ogiftcardordermaster.FResultCount>0) then
    if (ogiftcardordermaster.FOneItem.FCancelYn="N") and (ogiftcardordermaster.FOneItem.Fjumundiv<>"9") and (orefund.FOneItem.Freturnmethod<>"R120")  then
        response.write "<script>alert('��ǰ�ֹ� �Ǵ� �ֹ��� ��ҵ� ���, �Ǵ� �ſ�ī���Ϻ���Ҹ� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if

dim i
dim IsDirectCancelAvail
IsDirectCancelAvail = True

dim CancelCase , etcCancelCase

CancelCase = "Giftī�� �������"
etcCancelCase = "Giftī�� �������"
%>
<script language='javascript'>
function ActCancel(frm){
    if (frm.msg.value.length<1){
        alert('��һ����� �Է��� �ּ���.');
        frm.msg.focus();
        return;
    }

    if (frm.returnmethod.value=="R120"){
        //�κ����
        if (confirm('�κ� ��� ���� �Ͻðڽ��ϱ�?')){
            frm.action="pop_PartialCardCancel_process.asp";
            frm.submit();
        }
    }else{
        if (confirm('���� ��� �Ͻðڽ��ϱ�?')){
            frm.action="pop_giftcard_CardCancel_process.asp";
            frm.submit();
        }
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmCanncel" method="post" action="pop_giftcard_CardCancel_process.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="returnmethod" value="<%= orefund.FOneItem.Freturnmethod %>">
<% if (ogiftcardordermaster.FResultCount>0) then %>
<input type="hidden" name="rdsite" value="">
<input type="hidden" name="buyemail" value="<%= ogiftcardordermaster.FOneItem.Fbuyemail%>">
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
    <% if (ogiftcardordermaster.FResultCount>0) then %>
        <font color="<%= ogiftcardordermaster.FOneItem.CancelYnColor %>"><%= ogiftcardordermaster.FOneItem.CancelYnName %></font> <font color="<%= ogiftcardordermaster.FOneItem.IpkumDivColor %>"><%= ogiftcardordermaster.FOneItem.GetJumunDivName %>

        <% if (ogiftcardordermaster.FOneItem.Fjumundiv="9") then %>
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
    	<% if (ogiftcardordermaster.FResultCount>0) then %>
    	<% if ((C_ADMIN_AUTH) and (ogiftcardordermaster.FOneItem.Fjumundiv="9")) or (session("ssBctID")="icommang") or (session("ssBctID")="iroo4")  then %>
    	<input type="checkbox" name="force" >�ݾװ������
    	<% end if %>
    	<% end if %>
    </td>
</tr>

<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <% if (orefund.FOneItem.Freturnmethod="R120") then %>
    <input type="button" class="button" value=" ���� �κ� ��� " onClick="ActCancel(frmCanncel)">
    <% else %>
    <input type="button" class="button" value=" ���� ��� " onClick="ActCancel(frmCanncel)">
    <% end if %>
    </td>
</tr>
</form>
</table>
<%
set ocsaslist = Nothing
set orefund = Nothing
set ogiftcardordermaster = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->