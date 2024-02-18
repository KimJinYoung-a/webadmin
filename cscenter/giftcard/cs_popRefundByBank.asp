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
	Response.Write "<script>alert('아이디가 없습니다.');</script>"
	dbget.close()
	Response.End
end if

if (CLng(FormatNumber((100 * oTenGiftCard.FspendCash / oTenGiftCard.FgainCash),0)) < 60) and (userid<>"danbi2612") and (userid<>"setjddms") and (userid<>"dadareda") and (userid<>"eiddr0705") then
	Response.Write "<script>alert('Gift카드사용비율( = 상품구매총액/등록총액) 이 60% 이상인 경우만 잔액의 환불이 가능합니다.');</script>"
	dbget.close()
	Response.End
end if



'==============================================================================
'// 기프트카드의 구매자와 사용자는 다를수 있다.
'// 따라서 등록내역이 아닌 사용내역에서 주문번호를 가져온다.
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
	Response.Write "<script>alert('Gift카드 등록내역이 없습니다.[관리자 문의]');</script>"
	dbget.close()
	Response.End
end if

%>


<script language="javascript">
function refundByBank()
{
	if((document.getElementById("refundrequire").value * 0) != 0) {
		alert("숫자로만 입력하세요.");
		document.getElementById("refundrequire").focus();
		document.getElementById("refundrequire").select();
		return;
	}

	if((<%= currentCash %> - document.getElementById("refundrequire").value*1) < 0) {
		alert("전환할 예치금이 <%= FormatNumber(currentCash, 0) %> 보다 큽니다.\n<%= FormatNumber(currentCash, 0) %> 이하로 입력해 주세요.");
		document.getElementById("refundrequire").focus();
		document.getElementById("refundrequire").select();
		return;
	}

	if(confirm("무통장 환불하시겠습니까?") == true) {
		document.frmRefundByBank.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr height="30">
	<td style="padding-left:8px;"><img src="http://webadmin.10x10.co.kr/images/icon_arrow_link.gif"></td>
	<td style="padding-top:5px;"><b>무통장 환불</b></td>
</tr>
</table>
<form name="frmRefundByBank" method="post" action="cs_popRefundByBank_process.asp" style="margin:0px;">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="380" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">아이디</td>
  	<td bgcolor="#FFFFFF"><%= userid %></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">최근 주문번호</td>
  	<td bgcolor="#FFFFFF"><input type="text" class="text_ro" name="orderserial" value="<%= orderserial %>" readonly></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
  	<td bgcolor="#FFFFFF">
	  	<input class="text" type="text" size="20" name="rebankaccount" value="">
	  	<input class="csbutton" type="button" value="이전내역" onClick="popPreReturnAcct('<%= userid %>','frmRefundByBank','rebankaccount','rebankownername','rebankname');">
  	</td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
  	<td bgcolor="#FFFFFF">
  		<input class="text" type="text" size="20" name="rebankownername" value="">
  	</td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
  	<td bgcolor="#FFFFFF"><% DrawBankCombo "rebankname", "" %></td>
</tr>
<tr height="30">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">환불액</td>
  	<td bgcolor="#FFFFFF">
  		<input type="text" class="text" name="refundrequire" id="refundrequire" value="<%= currentCash %>" size="10"> 원 (Gift카드 잔액 : <%= FormatNumber(currentCash, 0) %> 원)
  	</td>
</tr>
</table>
</form>
<table class="a" width="390">
<tr height="30">
	<td align="right"><input type="button" value="환불하기" class="button" onClick="refundByBank()"></td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
