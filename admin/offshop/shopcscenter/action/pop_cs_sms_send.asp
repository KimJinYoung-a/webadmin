<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ������
' Hieditor : 2012.03.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
dim reqhp, smstext ,masteridx, defaultMsg , orderno
	reqhp 	= requestCheckVar(request("reqhp"),32)
	smstext = request("smstext")
	masteridx = requestCheckVar(request("masteridx"),10)
	orderno 	= requestCheckVar(request("orderno"),16)
	defaultMsg = request("defaultMsg")

if (defaultMsg="") then defaultMsg="[�ٹ�����]"

dim sqlstr
dim smstext1, smstext2, smstext3

if (reqhp<>"" and smstext<>"") then
	if defaultMsg <> "" then
		if checkNotValidHTML(defaultMsg) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if
	if smstext <> "" then
		if checkNotValidHTML(smstext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

    call SendOverLengthSMS(reqhp,"",smstext)

    response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"   
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
%>

<script language='javascript'>

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
 }
 
function SendSMS(frm){

	if (frm.reqhp.value.length<12){
		alert('�ڵ�����ȣ�� ��Ȯ�� �Է��ϼ���.');
		return;
    }
	if ((frm.smstext.value.length<1)||(frm.smstext.value=="[�ٹ�����]")){
		alert('���ڳ����� �Է��ϼ���.');
		return;
	}
    
    if (getByteLength(frm.smstext.value)>240){
        alert('SMS�� �ι� �̻����� ������ ���� �� �����ϴ�. ���ڼ��� �ٿ��ּ���');
        return;
    }
    
    if (getByteLength(frm.smstext.value)>80){
        var varconfirmMsg = '�޼����� 2ȸ �̻����� ������ ���� �մϴ�. ���� �Ͻðڽ��ϱ�?';
    }else{
        var varconfirmMsg = '���� �Ͻðڽ��ϱ�?';
    }
    
	var ret= confirm(varconfirmMsg);
	
	if(ret){
		frm.submit();
	}
}

window.resizeTo('280','320')

</script>

<!-- ǥ ��ܹ� ����-->
<table width="230" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="10" valign="bottom">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td>
    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>SMS�߼�</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="230" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/offshop/cscenter/action/pop_cs_sms_send.asp">
<input type="hidden" name="mode" value="send">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="orderno" class="text_ro" value="<%= orderno %>" size="20" maxlength="20" readonly></td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reqhp" class="text" value="<%= reqhp %>" size="15" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���ڳ���</td>
	<td bgcolor="#FFFFFF"><textarea name="smstext" class="textarea" cols="22" rows="5"><%= defaultMsg %></textarea></td>
</tr>
</form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="230" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
        <input type="button" class="button" value="SMS�߼�" onclick="SendSMS(frm);">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->