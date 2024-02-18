<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
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

if (defaultMsg="") then defaultMsg="[텐바이텐]"

dim sqlstr
dim smstext1, smstext2, smstext3

if (reqhp<>"" and smstext<>"") then
	if defaultMsg <> "" then
		if checkNotValidHTML(defaultMsg) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if
	if smstext <> "" then
		if checkNotValidHTML(smstext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

    call SendOverLengthSMS(reqhp,"",smstext)

    response.write "<script>alert('전송되었습니다.');</script>"   
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
		alert('핸드폰번호를 정확히 입력하세요.');
		return;
    }
	if ((frm.smstext.value.length<1)||(frm.smstext.value=="[텐바이텐]")){
		alert('문자내용을 입력하세요.');
		return;
	}
    
    if (getByteLength(frm.smstext.value)>240){
        alert('SMS를 두번 이상으로 나누어 보낼 수 없습니다. 글자수를 줄여주세요');
        return;
    }
    
    if (getByteLength(frm.smstext.value)>80){
        var varconfirmMsg = '메세지를 2회 이상으로 나누어 전송 합니다. 전송 하시겠습니까?';
    }else{
        var varconfirmMsg = '전송 하시겠습니까?';
    }
    
	var ret= confirm(varconfirmMsg);
	
	if(ret){
		frm.submit();
	}
}

window.resizeTo('280','320')

</script>

<!-- 표 상단바 시작-->
<table width="230" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="10" valign="bottom">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td>
    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>SMS발송</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<table width="230" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/offshop/cscenter/action/pop_cs_sms_send.asp">
<input type="hidden" name="mode" value="send">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="orderno" class="text_ro" value="<%= orderno %>" size="20" maxlength="20" readonly></td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">핸드폰번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="reqhp" class="text" value="<%= reqhp %>" size="15" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">문자내용</td>
	<td bgcolor="#FFFFFF"><textarea name="smstext" class="textarea" cols="22" rows="5"><%= defaultMsg %></textarea></td>
</tr>
</form>
</table>

<!-- 표 하단바 시작-->
<table width="230" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
        <input type="button" class="button" value="SMS발송" onclick="SendSMS(frm);">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->