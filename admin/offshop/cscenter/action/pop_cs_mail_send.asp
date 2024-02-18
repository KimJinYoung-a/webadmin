<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim email, masteridx
	email 	= requestCheckVar(request("email"),128)
	masteridx 	= requestCheckVar(request("masteridx"),10)
%>

<script type='text/javascript'>

function SendCSMail(mailform){

	if (mailform.mailto.value.length<1){
		alert('메일주소를 입력하세요.');
		return;
	}
	if (mailform.title.value.length<1){
		alert('메일제목를 입력하세요.');
		return;
	}
	if (mailform.contents.value.length<1){
		alert('메일내용를 입력하세요.');
		return;
	}

	var ret= confirm('전송 하시겠습니까?');
	
	if(ret){
		mailform.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="10" valign="bottom">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td>
    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>메일발송</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="mailform" method="post" action="pop_cs_mail_send_process.asp">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">메일주소</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mailto" class="text" value="<%= email %>" size="30"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">메일제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" class="text" value="" size="80"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">메일내용</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" class="textarea" value="" cols="80" rows="19"></textarea></td>
</tr>    
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
        <input type="button" class="button" value="메일발송" onclick="SendCSMail(mailform);">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</form>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->