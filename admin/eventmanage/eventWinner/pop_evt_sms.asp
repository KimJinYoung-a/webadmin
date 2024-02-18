<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/pop_evt_sms.asp
' Description :  이벤트 당첨자 SMS 작성 페이지
' History : 2007.09.27 김정인
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body topmargin="0" >

<%

dim evtCode,arridx

evtCode =request("eC")
arridx = chkarray(request("arridx"))

dim appOne,evtName
set appOne = new ClsEvent
appOne.FECode = evtCode
appOne.fnGetEventCont
evtName = appOne.FEName
set appOne = nothing


dim oSms,arrSms
set oSms = new ClsEventEntry
oSms.FECode = evtCode
arrSms = oSms.fnGetSms

set oSms = nothing


dim MsgCont,replyNumber,regUser,regDate

if not isArray(arrSms) then
	MsgCont = evtName & " 이벤트에 당첨되셨습니다.공지사항을 확인해주세요."
	replyNumber = "1644-6030"
else
	MsgCont = arrSms(1,0)
	replyNumber = arrSms(2,0)
	regUser = arrSms(3,0)
	regDate = arrSms(4,0)
end if
%>

<script>

function GetByteLength(val){
 	var real_byte = val.length;
 	for (var ii=0; ii<val.length; ii++) {
  		var temp = val.substr(ii,1).charCodeAt(0);
  		if (temp > 127) { real_byte++; }
 	}

   return real_byte;
}

function fnTextLengthChk(){

	var bytes = document.getElementById('bytetxt');

	bytes.innerHTML= GetByteLength(document.msgfrm.msg.value);

	if(GetByteLength(document.msgfrm.msg.value) > 80){
		alert("내용이 제한길이를 초과하였습니다. \n80 Byte 까지만 작성가능합니다.");
		return false;
	}
	return true;
}

function fnSendSms(){
	if(document.msgfrm.msg.value.length<1){
		alert('내용을 입력하셔야 합니다');
		return false;
	}
	if(document.msgfrm.reNo.value.length<1){
		alert('회신번호를 입력하셔야 합니다');
		return false;
	}
	if(fnTextLengthChk()){
		document.msgfrm.mode.value='send';
		document.msgfrm.target='subframe';
		document.msgfrm.submit();
	}
}

function fnSmsSave(){
	if(document.msgfrm.msg.value.length<1){
		alert('내용을 입력하셔야 합니다');
		return false;
	}
	if(document.msgfrm.reNo.value.length<1){
		alert('회신번호를 입력하셔야 합니다');
		return false;
	}

	if(fnTextLengthChk()){

		document.msgfrm.mode.value='save';
		document.msgfrm.target='subframe';
		document.msgfrm.submit();
	}
}

</script>

<!-- 테이블 상단 검색바 시작 -->
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td width="10" background="/images/tbl_blue_round_04.gif"></td>
        <td align="right"></td>
        <td>&nbsp;</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 중간 내용시작-->
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top" style="padding : 0 0 10 0">
        <td width="10" background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<form name="msgfrm" method="post" action="event_SMS_Process.asp">
        		<input type="hidden" name="mode" value="">
        		<input type="hidden" name="eC" value="<%= evtCode %>">
        		<input type="hidden" name="arridx" value="<%= arridx %>">

        		<tr>
        			<td>
        				<textarea name="msg" cols="17" rows="10" style="overflow=hidden" onkeyUp="fnTextLengthChk();"><%= MsgCont %></textarea>
        			</td>
        			<td valign="top" align="left">
        				<b>회신번호</b><br><input type="text" size="13" name="reNo" value="<%= replyNumber %>">
        			</td>
        		</tr>
        		<tr>
        			<td><span id="bytetxt">1</span>/<b>80</b></td>
        		</tr>
        		</form>
        	</table>
        </td>

        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>

<!-- 하단 시작 -->
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">
        	<input type="button" class="button" value="보내기" onclick="fnSendSms();">&nbsp;
        	<input type="button" class="button" value="임시저장" onclick="fnSmsSave();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<iframe name="subframe" src="" frameborder="0" width="100" height="100"></iframe>

<Script language="javascript">
fnTextLengthChk();
</script>
</html>

<!--
SMS 최종 발송,수정자 : <%= regUser %>
SMS 최종 발송,수정일자 : <%= regDate %>
-->
<!-- #include virtual="/lib/db/dbclose.asp" -->