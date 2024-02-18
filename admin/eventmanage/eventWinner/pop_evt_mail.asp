<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/pop_evt_mail.asp
' Description :  이벤트 당첨자 메일 작성 페이지
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
arridx = request("arridx")
dim appOne,evtName
set appOne = new ClsEvent
appOne.FECode = evtCode
appOne.fnGetEventCont
evtName = appOne.FEName
set appOne = nothing

dim oMail,arrMail
set oMail = new ClsEventEntry
oMail.FECode = evtCode
arrMail = oMail.fnGetMail
set oMail = nothing

dim mailTitle,mailContents,replyName,replyMail,regUser,regDate

If isarray(arrMail) Then
	mailTitle = arrMail(1,0)
	mailContents = db2html(arrMail(2,0))
	replyName = arrMail(3,0)
	replyMail = arrMail(4,0)
	regUser  = arrMail(5,0)
	regDate  = arrMail(6,0)
Else
	mailTitle = evtName & " 이벤트 당첨자 공지 메일입니다"
	mailContents = ""
	replyName = "tenbyten"
	replyMail = "customer@10x10.co.kr"
End If

mailContents = replace(mailContents,"<br>",vbcrlf)

%>

<script>

// 메일 발송
function fnSendMail(){

	document.msgfrm.mode.value='send';
	document.msgfrm.action='event_Mail_process.asp';
	document.msgfrm.target='subframe';
	document.msgfrm.submit();

}
// 메일 임시저장
function fnMailSave(){
	document.msgfrm.mode.value='save';
	document.msgfrm.action='event_Mail_process.asp';
	document.msgfrm.target='subframe';
	document.msgfrm.submit();

}
// 메일 미리보기
function fnMailPreview(){
	var popPre = window.open('pop_evt_mail_preview.asp','mailpreview','width=650,height=600,status=no,toolbar=no,scrollbars=yes');
	document.msgfrm.target='mailpreview';
	document.msgfrm.action='pop_evt_mail_preview.asp';
	document.msgfrm.submit();
	popPre.focus();

}

</script>

<!-- 테이블 상단 검색바 시작 -->
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
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
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top" style="padding : 0 0 10 0">
        <td width="10" background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        		<form name="msgfrm" method="post" action="">
        		<input type="hidden" name="mode" value="">
        		<input type="hidden" name="eC" value="<%= evtCode %>">
        		<tr bgcolor="#FFFFFF">
        			<td align="center" width="120">보내는 이름 </td><td align="left"><input type="text" name="rpName" size="10" value="<%= replyName %>"></td>
        		</tr>
        		<tr bgcolor="#FFFFFF">
        			<td align="center" width="120">보내는 메일 주소</td><td align="left"><input type="text" name="rpMail" size="20" value="<%= replyMail %>"></td>
        		</tr>
        		<tr bgcolor="#FFFFFF">
        			<td align="center" width="120">메일 제목</td><td align="left"><input type="text" name="mlTitle" size="60" value="<%= mailTitle %>"></td>
        		</tr>
        		<tr bgcolor="#FFFFFF">
        			<td align="center" width="120">메일 내용</td><td align="left"><textarea name="mlCont" cols="70" rows="20"><%= mailContents %></textarea></td>
        		</tr>
        		</form>
        	</table>
        </td>

        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>


<!-- 하단 시작 -->
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">
        	<input type="button" class="button" value="보내기" onclick="fnSendMail();">&nbsp;
        	<input type="button" class="button" value="임시저장" onclick="fnMailSave();">&nbsp;
        	<input type="button" class="button" value="미리보기" onclick="fnMailPreview();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<iframe name="subframe" src="" frameborder="0" width="0" height="0"></iframe>

<Script language="javascript">
fnTextLengthChk();
</script>
</html>
<!--
Mail 최종 발송,수정자 : <%= regUser %>
Mail 최종 발송,수정일자 : <%= regDate %>
-->
<!-- #include virtual="/lib/db/dbclose.asp" -->