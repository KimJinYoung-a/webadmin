<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 매출등록
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcardcls.asp"-->
<%

dim i
dim shopid, cardPrice, cardNo, payMethod
dim cardNoValid, cardNoErrMsg, saveOrder
dim sqlStr

shopid = requestCheckVar(request("shopid"),32)
cardPrice = requestCheckVar(request("cardPrice"),32)
cardNo = requestCheckVar(request("cardNo"),32)
payMethod = requestCheckVar(request("payMethod"),32)
cardNoValid = requestCheckVar(request("cardNoValid"),32)
cardNoErrMsg = requestCheckVar(request("cardNoErrMsg"),32)
saveOrder = requestCheckVar(request("saveOrder"),32)

if (shopid = "") then
	shopid = "streetshop011"
end if

function CheckCardNo(cardNo)
	dim cardNoValid, cardNoErrMsg
	dim sqlStr

	cardNoValid = "N"
	cardNoErrMsg = ""

	if Len(cardNo) <> 16 then
		cardNoErrMsg = "잘못된 카드번호입니다."
		CheckCardNo = cardNoErrMsg
		exit function
	end if

	sqlStr = " select top 1 statusDiv, giftOrderSerial, designId "
	sqlStr = sqlStr + " from db_order.dbo.tbl_giftcard_offMasterCd where masterCardCode = '" & cardNo & "' "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		if rsget("statusDiv") <> "W" or Not IsNull(rsget("giftOrderSerial")) then
			cardNoErrMsg = "이미 사용중인 카드번호입니다."
		end if

		if IsNull(rsget("designId")) then
			cardNoErrMsg = "디자인 미등록 카드번호입니다."
		end if
	else
		cardNoErrMsg = "등록되지 않은 카드번호입니다."
	end if
	rsget.close

	CheckCardNo = cardNoErrMsg
end function


cardNoValid = "N"
if (cardNo = "") then
	cardNoErrMsg = "카드번호를 입력하세요."
elseif cardNoValid <> "Y" then
	cardNoErrMsg = CheckCardNo(cardNo)
	if cardNoErrMsg = "" then
		cardNoValid = "Y"
	end if
end if

Dim refIP : refIP = Request.ServerVariables ("REMOTE_ADDR")
dim errCode, errMsg
if (saveOrder = "Y") then
	if (shopid <> "") and (cardPrice <> "") and (cardNo <> "") and (cardNoValid = "Y") and (payMethod <> "") then
		sqlStr = " exec [db_order].[dbo].[usp_Ten_GiftCard_CheckSaveOrder] '" & shopid & "', " & cardPrice & ", '" & cardNo & "', '" & payMethod & "', '" & refIP & "' "
		''response.write sqlStr
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget

		errCode = ""
		errMsg = ""
		if Not rsget.Eof then
			errCode = rsget("errCode")
			errMsg = rsget("errMsg")
		end if
		rsget.close

		if errCode <> "0000" then
			response.write "<h1>에러 : " & rsget("errMsg") & "</h1>"
		else
			response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
		end if
		dbget.close : response.end
	else
		response.write "<h1>잘못된 접근입니다.</h1>"
		dbget.close : response.end
	end if
end if


'================================================================================
dim oOffShopCardPromotion

set oOffShopCardPromotion = new COffShopCardPromotion

oOffShopCardPromotion.FRectShopid = shopid
''oOffShopCardPromotion.FRectCardPrice = cardPrice

oOffShopCardPromotion.FCurrPage = 1
oOffShopCardPromotion.Fpagesize = 10

if (shopid <> "") then
	oOffShopCardPromotion.COffShopCardPromotionList
end if

%>
<script language="JavaScript" src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script>

function jsSetCardPrice(obj) {
	var frm = document.frm;
	if (jsCheckCardNo() == false) { return; }
	frm.cardPrice.value = obj.value.replace(",", "");
	jsSubmit();
}

function jsSetPayMethod(obj) {
	var frm = document.frm;
	if (jsCheckCardNo() == false) { return; }
	if (obj.id == "payMethodBtn0") {
		frm.payMethod.value = "7";
	} else if (obj.id == "payMethodBtn1") {
		frm.payMethod.value = "100";
	}
	jsSubmit();
}

function jsSubmit() {
	var frm = document.frm;
	frm.submit();
}

function jsSaveOrder() {
	var frm = document.frm;

	//var cardNo = document.getElementById("cardNo");
	//frm.cardNo.value = cardNo.value;

	if (jsCheckCardNo() == false) { return; }

	if (frm.shopid.value == "") {
		alert("매장이 선택되지 않았습니다.");
		return;
	}

	if (frm.cardPrice.value == "") {
		alert("금액이 선택되지 않았습니다.");
		return;
	}

	if (frm.cardNo.value == "") {
		alert("카드번호가 입력되지 않았습니다.");
		return;
	}

	if (frm.payMethod.value == "") {
		alert("결제방법이 선택되지 않았습니다.");
		return;
	}

	if (confirm("판매등록 하시겠습니까?") == true) {
		frm.saveOrder.value = "Y";
		frm.submit();
	}
}

function jsCheckCardNo() {
	var frm = document.frm;
	var cardNo = document.getElementById("cardNo");

	if (cardNo.value.length == 16) {
		if (frm.cardNo.value != cardNo.value) {
			frm.cardNo.value = cardNo.value;
			frm.cardNoValid.value = "";
			frm.cardNoErrMsg.value = "";
			return true;
		} else {
			return true;
		}
	} else if (cardNo.value.length == 0) {
		alert("먼저 카드번호를 입력하세요.");
		cardNo.focus();
		return false;
	} else {
		alert("잘못된 카드번호입니다.");
		cardNo.select();
		return false;
	}
}

$(document).ready(function() {
	var cardNo = document.getElementById("cardNo");
	cardNo.focus();
	<% if cardNoValid <> "Y" then %>
	//cardNo.select();
	<% end if %>
});

</script>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
	<tr>
		<td bgcolor="#EEEEEE" height="35" align="center"><h3>기프트카드 판매등록</h3></td>
	</tr>
</table>

<p />

<table width="75%" height="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class=a align="center">
	<tr>
		<td height="50" width="100" align="left"><h3>매장ID</h3></td>
		<td align="left"><h3><%= shopid %></h3></td>
	</tr>
	<tr>
		<td height="50" width="100" align="left"><h3>카드번호</h3></td>
		<td align="left">
			<h3>
				<input type="text" name="cardNo" value="<%= cardNo %>" size="32" id="cardNo" style="width:350px; height:50px; font-size:25px;" onKeyUp="if (window.event.keyCode == 13) { if (jsCheckCardNo() == true){ jsSubmit(); } }">
				<% if cardNoValid = "Y" then %>
				* 등록가능한 카드입니다.
				<% else %>
				<font color="red">* <%= cardNoErrMsg %></font>
				<% end if %>
			</h3>
		</td>
	</tr>
	<tr>
		<td height="50" width="100" align="left"><h3>금액</h3></td>
		<td align="left">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="10000", " background-color:#A6F4FF", "") %>" value="10,000" id="priceBtn0" onClick="jsSetCardPrice(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="20000", " background-color:#A6F4FF", "") %>" value="20,000" id="priceBtn1" onClick="jsSetCardPrice(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="30000", " background-color:#A6F4FF", "") %>" value="30,000" id="priceBtn2" onClick="jsSetCardPrice(this)">
			<p />
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="50000", " background-color:#A6F4FF", "") %>" value="50,000" id="priceBtn3" onClick="jsSetCardPrice(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="80000", " background-color:#A6F4FF", "") %>" value="80,000" id="priceBtn4" onClick="jsSetCardPrice(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="100000", " background-color:#A6F4FF", "") %>" value="100,000" id="priceBtn5" onClick="jsSetCardPrice(this)">
			<p />
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="200000", " background-color:#A6F4FF", "") %>" value="200,000" id="priceBtn6" onClick="jsSetCardPrice(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(cardPrice="300000", " background-color:#A6F4FF", "") %>" value="300,000" id="priceBtn7" onClick="jsSetCardPrice(this)">
		</td>
	</tr>
	<tr>
		<td height="50" width="100" align="left"><h3>프로모션</h3></td>
		<td align="left">
			<% if oOffShopCardPromotion.FResultCount > 0 then %>
			<% for i=0 to oOffShopCardPromotion.FResultcount-1 %>
			<% if (Left(Now,10) >= oOffShopCardPromotion.FItemList(i).FstartDate) and (Left(Now,10) <= oOffShopCardPromotion.FItemList(i).FendDate) then %>
			<h3><%= FormatNumber(oOffShopCardPromotion.FItemList(i).FcardPrice, 0) %> 원권 구매시,
				마일리지 <%= oOffShopCardPromotion.FItemList(i).FrateAmmount %><%= CHKIIF(oOffShopCardPromotion.FItemList(i).FrateGubun=1, "%", "포인트") %> 추가적립(~<%= oOffShopCardPromotion.FItemList(i).FendDate %>)<br /></h3>
			<% end if %>
			<% next %>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td height="50" width="100" align="left"><h3>결제방법</h3></td>
		<td align="left">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(payMethod="7", " background-color:#A6F4FF", "") %>" value="현금" id="payMethodBtn0" onClick="jsSetPayMethod(this)">
			<input type="button" style="width:150px; height:80px; font-size:25px;<%= CHKIIF(payMethod="100", " background-color:#A6F4FF", "") %>" value="신용카드" id="payMethodBtn1" onClick="jsSetPayMethod(this)">
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center" height="150">
			<input type="button" style="width:250px; height:100px; font-size:25px;" value="저장" onClick="jsSaveOrder()">
		</td>
	</tr>
</table>
<form name="frm" method="get">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="cardPrice" value="<%= cardPrice %>">
	<input type="hidden" name="cardNo" value="<%= cardNo %>">
	<input type="hidden" name="cardNoValid" value="<%= cardNoValid %>">
	<input type="hidden" name="cardNoErrMsg" value="<%= cardNoErrMsg %>">
	<input type="hidden" name="payMethod" value="<%= payMethod %>">
	<input type="hidden" name="saveOrder" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
