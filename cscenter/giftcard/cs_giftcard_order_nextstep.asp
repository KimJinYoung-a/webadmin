<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<%

dim giftorderserial
giftorderserial = RequestCheckVar(request("giftorderserial"),11)

'==============================================================================
dim oGiftOrder

set oGiftOrder = new cGiftCardOrder

if (giftorderserial <> "") then
	oGiftOrder.FRectGiftOrderSerial = giftorderserial

	oGiftOrder.getCSGiftcardOrderDetail
end if


dim IsAdminLogin
IsAdminLogin = C_ADMIN_AUTH



'==============================================================================
dim nextjumundiv, nextipkumdiv, prevjumundiv, previpkumdiv

if (oGiftOrder.FOneItem.Fjumundiv = "1") and (oGiftOrder.FOneItem.Fipkumdiv = "2") then
	'결제완료진행
    nextjumundiv = "3"
    nextipkumdiv = "4"

    if (oGiftOrder.FOneItem.FbookingYN = "N") then
	    nextjumundiv = "5"
	    nextipkumdiv = "8"
    end if
elseif (oGiftOrder.FOneItem.Fjumundiv = "3") then
	'전송완료진행
    nextjumundiv = "5"
    nextipkumdiv = "8"
elseif (oGiftOrder.FOneItem.Fjumundiv = "5") then
	'등록완료진행
    nextjumundiv = "7"
    nextipkumdiv = "8"
end if



%>
<script language='javascript'>
function SubmitForm() {
	<% if False and (IsAdminLogin <> True) then %>
		alert("시스템팀 전용메뉴입니다.");
		return;
	<% end if %>

    if ((frm.jumundiv.value == '1') && (frm.ipkumdiv.value != '2')) {
        alert('주문 접수 내역이 아닙니다.[시스템팀 문의]');
        return;
    }

    if (frm.jumundiv.value == '7') {
        alert('등록완료내역은 다음단계로 진행할 수 없습니다.');
        return;
    }

    if (frm.cancelyn.value!='N'){
        alert('취소된 내역은 다음 상태로 진행이 불가능 합니다.');
        return;
    }

    if (confirm("다음 상태로 진행 하시겠습니까?") == true) {
        frm.submit();
    }
}

function checkHPbyUserID() {
	if (frm.requserid.value == "") {
		alert("카드 사용자 ID 를 입력하세요.");
		return;
	}

	ifr.location.href = "cs_giftcard_order_nextstep_iframe.asp?requserid=" + frm.requserid.value;
}

function findUserIDbyHP() {
	if (frm.reqhp.value == "") {
		alert("카드 사용자 HP 를 입력하세요.");
		return;
	}

	ifr.location.href = "cs_giftcard_order_nextstep_iframe.asp?reqhp=" + frm.reqhp.value;
}

</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
<input type="hidden" name="mode" value="jumundivnextstep">
<input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
<input type="hidden" name="jumundiv" value="<%= oGiftOrder.FOneItem.Fjumundiv %>">
<input type="hidden" name="ipkumdiv" value="<%= oGiftOrder.FOneItem.Fipkumdiv %>">
<input type="hidden" name="userid" value="<%= oGiftOrder.FOneItem.Fuserid %>">
<input type="hidden" name="cancelyn" value="<%= oGiftOrder.FOneItem.Fcancelyn %>">
<input type="hidden" name="bookingYN" value="<%= oGiftOrder.FOneItem.FbookingYN %>">
<input type="hidden" name="reqhp" value="<%= oGiftOrder.FOneItem.Freqhp %>">
    <tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="160">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>다음 상태 진행</b>
				    </td>
				    <td align="right">
				    	<input type="button" value="다음 상태 진행" class="csbutton" onclick="javascript:SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
        <td width="50%" >현재상태</td>
        <td width="50%" >다음상태</td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td width="50%"><font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetJumunDivName %></font></td>
        <%
        prevjumundiv = oGiftOrder.FOneItem.Fjumundiv
        previpkumdiv = oGiftOrder.FOneItem.Fipkumdiv
        oGiftOrder.FOneItem.Fjumundiv = nextjumundiv
        oGiftOrder.FOneItem.Fipkumdiv = nextipkumdiv
        %>
        <td width="50%"><font color="<%= oGiftOrder.FOneItem.IpkumDivColor %>"><%= oGiftOrder.FOneItem.GetJumunDivName %></font></td>
        <%
        oGiftOrder.FOneItem.Fjumundiv = prevjumundiv
        oGiftOrder.FOneItem.Fipkumdiv = previpkumdiv
        %>
    </tr>
    <% if (nextjumundiv="7") and (oGiftOrder.FOneItem.Fcancelyn="N") then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
    	<td width="50%">카드사용자ID</td>
    	<td width="50%">
    		<input type="text" class="text" name="requserid" size="15" value="">
    		<input type="button" class="button" value="검증" onClick="checkHPbyUserID()">
    	</td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
    	<td width="50%">카드사용자HP</td>
    	<td width="50%">
    		<%= oGiftOrder.FOneItem.Freqhp %>
    		<input type="button" class="button" value="검색" onClick="findUserIDbyHP()">
    	</td>
    </tr>
    <% end if %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td colspan="2"><%= oGiftOrder.FOneItem.GetAccountdivName %> <font color="<%= oGiftOrder.FOneItem.CancelYnColor %>"><%= oGiftOrder.FOneItem.CancelYnName %></font></td>
    </tr>
    <% if (oGiftOrder.FOneItem.Fipkumdiv="2") and (oGiftOrder.FOneItem.Fcancelyn="N") then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="emailok" checked >입금확인메일발송</td>
        <td ><%= oGiftOrder.FOneItem.Fbuyemail %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="smsok" checked >입금확인SMS발송</td>
        <td><%= oGiftOrder.FOneItem.Fbuyhp %></td>
    </tr>
    <% end if %>
</table>
<iframe src="" name="ifr" scrolling="no" marginwidth="0" marginheight="0" frameborder="0" vspace=0" hspace="0" height="0" width="0"></iframe>
<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
