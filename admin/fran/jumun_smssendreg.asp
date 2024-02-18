<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : [OFF]오프_출고관리>>주문관리(물류) 문자 발송 페이지
' History : 2022.04.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderutf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim menupos, masteridx, ojumunmaster, hp, defaultsmstext, defaulttitle, smstext, title, paymentgroup
	masteridx 	= requestCheckvar(getNumeric(request("masteridx")),10)
    menupos 	= requestCheckvar(getNumeric(request("menupos")),10)
    title 	= requestCheckvar(request("title"),120)
    smstext 	= request("smstext")
	paymentgroup   = requestcheckvar(request("paymentgroup"),32)

if not(paymentgroup="ORDER" or paymentgroup="CARTOONBOX") then
	response.write "<script>alert('정상적인 경로가 아닙니다.');</script>"
    dbget.close() : response.end
end if

' 주문단위
if paymentgroup="ORDER" then
	set ojumunmaster = new COrderSheet
		ojumunmaster.FRectIdx = masteridx

		if masteridx<>"" then
			ojumunmaster.GetOneOrderSheetMaster
		end if

' 박스단위
elseif paymentgroup="CARTOONBOX" then
	set ojumunmaster = new CCartoonBox
		ojumunmaster.FRectMasterIdx = masteridx
		ojumunmaster.GetMasterOne
end if

if ojumunmaster.FtotalCount < 1 then
    response.write "<script type='text/javascript'>"
    response.write "    alert('해당되는 주문건이 없습니다.');"
    response.write "</script>"
    session.codePage = 949
    dbget.close() : response.end
end if

' 주문단위
if paymentgroup="ORDER" then
	defaulttitle="[텐바이텐] 홀세일 안내 입니다."
	defaultsmstext="10x10 홀세일 안내 입니다."&vbcrlf&"주문하신 오더넘버 "& ojumunmaster.FOneItem.Fbaljucode &" 의 출고준비가 완료되었습니다."&vbcrlf&"아래의 방법으로 결재를 진행해주세요."&vbcrlf&vbcrlf
	'defaultsmstext=defaultsmstext&"카드결재, 입금 중 선택 가능합니다"&vbcrlf
	defaultsmstext=defaultsmstext&"입금 완료 후 담당자에게 출고요청을 부탁드립니다."&vbcrlf
	defaultsmstext=defaultsmstext&"기업은행 277-039188-04-031"&vbcrlf

' 박스단위
elseif paymentgroup="CARTOONBOX" then
	defaulttitle="[텐바이텐] 홀세일 안내 입니다."
	defaultsmstext="10x10 홀세일 안내 입니다."&vbcrlf&"주문하신 오더넘버 "& ojumunmaster.FOneItem.Fidx &" 의 출고준비가 완료되었습니다."&vbcrlf&"아래의 방법으로 결재를 진행해주세요."&vbcrlf&vbcrlf
	'defaultsmstext=defaultsmstext&"카드결재, 입금 중 선택 가능합니다"&vbcrlf
	defaultsmstext=defaultsmstext&"입금 완료 후 담당자에게 출고요청을 부탁드립니다."&vbcrlf
	defaultsmstext=defaultsmstext&"기업은행 277-039188-04-031"&vbcrlf

end if

'[해외문자] 
'This is a 10x10 wholesale information. The shipping preparation for the order number SJ807491 has been completed. Please proceed with the payment using the method below.
'Bank: Industrial Bank of Korea
'Branch: Daehakro
'Address: 101, Dongsung-gil, Jongno-gu, Seoul, Republic of Korea
'USD Account No.:277-039188-04-031
'SWIFT code: IBKOKRSE (0032777)

hp = ojumunmaster.FOneItem.fmanager_hp
if (hp<>"" and smstext<>"") then
	if LenB(smstext) > 80 then
    	Call SendNormalLMS(hp, title, "", smstext)
    else
    	Call SendNormalSMS_LINK(hp, "", smstext)
    end if

	' 주문단위
	if paymentgroup="ORDER" then
    	call ordersheetsmssend(masteridx)

	' 박스단위
	elseif paymentgroup="CARTOONBOX" then
    	call cartoonboxsmssend(masteridx)

	end if

    response.write "<script type='text/javascript'>"
    response.write "    alert('전송되었습니다.');"
    session.codePage = 949
	response.write "    opener.location.reload();"
    response.write "    window.close();"
    response.write "</script>"
    dbget.close()	:	response.End
end if

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

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
	if (frm.hp.value.length<9){		// - 까지 체크하게 하지말것. 너무 불편하다고함.
		alert('휴드폰번호를 정확히 입력하세요.');
		return;
    }
	if (frm.smstext.value.length<1){
		alert('문자내용을 입력하세요.');
		return;
	}
    if (getByteLength(frm.smstext.value) > 2000) {
        alert('1000천 글자 이상을 입력할 수 없습니다.\n\n글자수를 줄여주세요');
        return;
    }

    if (getByteLength(frm.smstext.value) > 80) {
        var varconfirmMsg = 'LMS 를 이용해 문자를 전송합니다. 전송 하시겠습니까?';
    }else{
        var varconfirmMsg = '전송 하시겠습니까?';
    }

	var ret= confirm(varconfirmMsg);
	if(ret){
		frm.submit();
	}
}

function updateChar() {
	var length = calculate_msglen(document.getElementById("smstext").value);

	if (length <= 80) {
		document.getElementById("charlen").innerHTML = "(" + length + "/80)<br><br>SMS";

		document.getElementById("title").className = "text_ro";
		document.getElementById("title").readOnly = true;
		document.getElementById("title").value = "";
	} else {
		document.getElementById("charlen").innerHTML = "(" + length + "/2000)<br><br><font color='red'>LMS</font>";

		document.getElementById("title").className = "text";
		document.getElementById("title").readOnly = false;
	}
}

function calculate_msglen(message) {
	var nbytes = 0;

	for (i=0; i<message.length; i++) {
		var ch = message.charAt(i);

		if(escape(ch).length > 4) {
			nbytes += 2;  // 한글일때 2씩 더함
		} else if(ch == '\n') {
			if (message.charAt(i-1) != '\r') {
				nbytes += 1;  // Enter일때 1씩 더함
			}
		} else {
			nbytes += 1;  // 기타 문자들일때 1씩 더함
		}
	}

	return nbytes;
}

function getOnload(){
	updateChar();
}

window.onload = getOnload;

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		문자발송.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="/admin/fran/jumun_smssendreg.asp" style="margin:0px;" >
<input type="hidden" name="mode" value="send">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="paymentgroup" value="<%= paymentgroup %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>" align="center">공급받는자</td>
	<td bgcolor="#FFFFFF">
		<%
		' 주문단위
		if paymentgroup="ORDER" then
		%>
			<%= ojumunmaster.FOneItem.Fbaljuid %>
		<%
		' 박스단위
		elseif paymentgroup="CARTOONBOX" then
		%>
			<%= ojumunmaster.FOneItem.Fshopid %>
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">이전발송날짜</td>
	<td bgcolor="#FFFFFF">
		<%= ojumunmaster.FOneItem.fsmssenddate %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">휴대폰</td>
	<td bgcolor="#FFFFFF"><input type="text" id="hp" name="hp" class="text" value="<%= hp %>" size="15" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">제목</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input id="title" type="text" name="title" class="text" value="<%= defaulttitle %>" size="80" maxlength="120" readonly>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">
		문자내용<br>
		<div id="charlen"></div>
	</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea id="smstext" name="smstext" class="textarea" cols="80" rows="10" onKeyUp="updateChar()"><%= defaultsmstext %></textarea></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="4" align="center">
		<input type="button" class="button" value="SMS발송" onclick="SendSMS(frm);">
	</td>
</tr>
</table>
</form>

<%
set ojumunmaster=nothing
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
