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

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "�̹� ������ �����Դϴ�." + Chr(34) + ")</script>"
	    response.write "�̹� ������ �����Դϴ�."
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		response.write "<script>alert(" + Chr(34) + "�̹� �Ϸ�� �����Դϴ�." + Chr(34) + ")</script>"
		response.write "�̹� �Ϸ�� �����Դϴ�."
		dbget.close()	:	response.End
	end if
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

''�ֹ� ����Ÿ
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    IF (ocsaslist.FOneItem.Frefminusorderserial<>"") then
        oordermaster.FRectOrderSerial = ocsaslist.FOneItem.Frefminusorderserial
    ELSE
        oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    ENd IF

    oordermaster.QuickSearchOrderMaster
end if

''�ֹ� ������
dim oorderdetail
set oorderdetail = new COrderMaster

if (oordermaster.FResultCount>0) then
    oorderdetail.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oorderdetail.QuickSearchOrderDetail
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

''�������� ������� '' ���޸��������� ��������� �Ϸ� //2013/01/14
if (oordermaster.FOneItem.Faccountdiv="550") then
    response.write "<script>alert('������ ��Ҵ� ����θ� �����մϴ�.\n\n���޸����� ����� ���� �����ÿ� ��ҿ�û�� ����Ϸ� �Ͻñ� �ٶ��ϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'' IniPay �� ��Ҹ� ���� INIMX_CARD INIMX_ISP_
dim IsInicisTID : IsInicisTID = False
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="IniTechPG_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_CARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_ISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtCARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,6)="Stdpay")
''IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_AUTH")

if (oordermaster.FOneItem.Fpggubun <> "KA") and (oordermaster.FOneItem.Fpggubun <> "KK") and (oordermaster.FOneItem.Fpggubun <> "CH") and (oordermaster.FOneItem.Fpggubun <> "NP") and (oordermaster.FOneItem.Fpggubun <> "PY") and (oordermaster.FOneItem.Fpggubun <> "TS") and Not IsInicisTID AND orefund.FOneItem.Freturnmethod<>"R400" AND orefund.FOneItem.Freturnmethod<>"R420" then
    response.write "<script>alert('�̴Ͻý�,īī��PAY, ���̹�����, ������, �佺 �ŷ��� ��� �����մϴ�.["&oordermaster.FOneItem.Fpggubun&"]aaaa');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
''rw oordermaster.FOneItem.FOrderSERIAL

if (oordermaster.FResultCount>0) then
    if (oordermaster.FOneItem.FCancelYn="N") and (oordermaster.FOneItem.Fjumundiv<>"9") and (orefund.FOneItem.Freturnmethod<>"R120") and (orefund.FOneItem.Freturnmethod<>"R022") and (orefund.FOneItem.Freturnmethod<>"R420")  then
        response.write "<script>alert('��ǰ�ֹ� �Ǵ� �ֹ��� ��ҵ� ���\n\n�ſ�ī���Ϻ���� �Ǵ� �ڵ����κ���Ҹ� ��� �����մϴ�.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if

dim i
dim IsDirectCancelAvail
IsDirectCancelAvail = True

for i=0 to oorderdetail.FResultCount - 1
    if (oorderdetail.FItemList(i).FItemId<>0) then
        if (Not (IsNULL(oorderdetail.FItemList(i).Fcurrstate) or (oorderdetail.FItemList(i).Fcurrstate<3))) then
            IsDirectCancelAvail = False
        end if
    end if
next

dim CancelCase , etcCancelCase

if (Left(ocsaslist.FOneItem.FOrderSerial,1)="A") or (Left(ocsaslist.FOneItem.FOrderSerial,1)="B") then
    CancelCase="�����ּ�"
else
    CancelCase="��������"

    if (oordermaster.FOneItem.Fjumundiv="9") then
        etcCancelCase = "��ǰ"
    elseif (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") or (orefund.FOneItem.Freturnmethod="R420") then
        etcCancelCase = "������κ����"
    end if
end if
%>
<script language='javascript'>
function ActCancel(frm){
    if (frm.msg.value.length<1){
        alert('��һ����� �Է��� �ּ���.');
        frm.msg.focus();
        return;
    }

    if ((frm.returnmethod.value=="R120") || (frm.returnmethod.value=="R022") || (frm.returnmethod.value=="R420")) {
        //�κ����(�ſ�ī��, �ڵ���)
        if (confirm('�κ� ��� ���� �Ͻðڽ��ϱ�?')){
            frm.action="pop_PartialCardCancel_process_test.asp";
            frm.submit();
        }
    }else{
        if (confirm('���� ��� �Ͻðڽ��ϱ�?')){
            frm.action="pop_CardCancel_process_20150812.asp";
            frm.submit();
        }
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmCanncel" method="post" action="pop_CardCancel_process_20150812.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="returnmethod" value="<%= orefund.FOneItem.Freturnmethod %>">
<% if (oordermaster.FResultCount>0) then %>
<input type="hidden" name="rdsite" value="<%= oordermaster.FOneItem.Frdsite%>">
<input type="hidden" name="buyemail" value="<%= oordermaster.FOneItem.Fbuyemail%>">
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
    <% if (oordermaster.FResultCount>0) then %>
        <font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font> <font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %>

        <% if (oordermaster.FOneItem.Fjumundiv="9") then %>
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
        <% if (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") then %>
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
    	<% if (oordermaster.FResultCount>0) then %>
    	<% if ((C_ADMIN_AUTH) and (oordermaster.FOneItem.Fjumundiv="9")) or (session("ssBctID")="icommang") or (session("ssBctID")="iroo4")  then %>
    	<input type="checkbox" name="force" >�ݾװ������
    	<% end if %>
    	<% end if %>
    </td>
</tr>

<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <% if (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") then %>
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
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
