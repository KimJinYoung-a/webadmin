<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/sp_tenGiftCardCls.asp" -->

<%

dim userid, orderserial, currentCash
dim sqlStr

userid      = request("userid")



'==============================================================================
dim oTenGiftCard

set oTenGiftCard = new CTenGiftCard

oTenGiftCard.FRectUserID = userid

currentCash = 0
if (userid<>"") then
    oTenGiftCard.getUserCurrentTenGiftCard

    currentCash = oTenGiftCard.FcurrentCash
end if



'==============================================================================
if (userid = "") then
	Response.Write "<script>alert('���̵� �����ϴ�.');</script>"
	dbget.close()
	Response.End
end if

if (CLng(FormatNumber((100 * oTenGiftCard.FspendCash / oTenGiftCard.FgainCash),0)) < 60) and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") and (userid<>"eiddr0705") then
	Response.Write "<script>alert('Giftī�������( = ��ǰ�����Ѿ�/����Ѿ�) �� 60% �̻��� ��츸 �ܾ��� ȯ���� �����մϴ�.');</script>"
	dbget.close()
	Response.End
end if



'==============================================================================
'// ����Ʈī���� �����ڿ� ����ڴ� �ٸ��� �ִ�.
'// ���� ��ϳ����� �ƴ� ��볻������ �ֹ���ȣ�� �����´�.
sqlStr = " select top 1 orderserial "
sqlStr = sqlStr + "	from "
sqlStr = sqlStr + "	db_user.dbo.tbl_giftCard_log "
sqlStr = sqlStr + "	where userid = '" + CStr(userid) + "' and jukyocd = 200 and deleteyn = 'N' "
sqlStr = sqlStr + "	order by idx desc "
rsget.Open sqlStr,dbget,1
If Not rsget.Eof Then
	orderserial = rsget("orderserial")
End IF
rsget.close()

if (orderserial = "") and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") then
	Response.Write "<script>alert('Giftī�� ��ϳ����� �����ϴ�.[������ ����]');</script>"
	dbget.close()
	Response.End
end if

%>


<script language="javascript">
function refundByBank()
{
	if((document.getElementById("refundrequire").value * 0) != 0) {
		alert("���ڷθ� �Է��ϼ���.");
		document.getElementById("refundrequire").focus();
		document.getElementById("refundrequire").select();
		return;
	}

	if((<%= currentCash %> - document.getElementById("refundrequire").value*1) < 0) {
		alert("��ȯ�� ��ġ���� <%= FormatNumber(currentCash, 0) %> ���� Ů�ϴ�.\n<%= FormatNumber(currentCash, 0) %> ���Ϸ� �Է��� �ּ���.");
		document.getElementById("refundrequire").focus();
		document.getElementById("refundrequire").select();
		return;
	}

	if(confirm("������ ȯ���Ͻðڽ��ϱ�?") == true) {
		document.frmRefundByBank.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr height="30">
	<td style="padding-left:8px;"><img src="http://webadmin.10x10.co.kr/images/icon_arrow_link.gif"></td>
	<td style="padding-top:5px;"><b>������ ȯ��</b></td>
</tr>
</table>
<form name="frmRefundByBank" method="post" action="cs_popRefundByBank_process.asp" style="margin:0px;">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="380" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
  	<td bgcolor="#FFFFFF"><%= userid %></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�ֱ� �ֹ���ȣ</td>
  	<td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="orderserial" value="<%= orderserial %>" readonly></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
  	<td bgcolor="#FFFFFF">
	  	<input class="text" type="text" size="20" name="rebankaccount" value="">
	  	<input class="csbutton" type="button" value="��������" onClick="popPreReturnAcct('<%= userid %>','frmRefundByBank','rebankaccount','rebankownername','rebankname');">
  	</td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
  	<td bgcolor="#FFFFFF">
  		<input class="text" type="text" size="20" name="rebankownername" value="">
  	</td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
  	<td bgcolor="#FFFFFF"><% DrawBankCombo "rebankname", "" %></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">ȯ�Ҿ�</td>
  	<td bgcolor="#FFFFFF">
  		<input type="text" class="text" name="refundrequire" id="refundrequire" value="<%= currentCash %>" size="10"> �� (Giftī�� �ܾ� : <%= FormatNumber(currentCash, 0) %> ��)
  	</td>
</tr>
</table>
</form>
<table class="a" width="390">
<tr height="30">
	<td align="right"><input type="button" value="ȯ���ϱ�" class="button" onClick="refundByBank()"></td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
