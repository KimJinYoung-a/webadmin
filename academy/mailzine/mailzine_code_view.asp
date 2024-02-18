<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_mailzinecls.asp"-->

<%

Dim omail,ix,idx

idx = RequestCheckvar(request("idx"),10)

set omail = new CUploadMaster
omail.MailzineDetail idx

%>
<table align="center" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td>
<input type="text" name="title" size="100" class="input" readonly value="<% = omail.Ftitle %>"><br>
<textarea name="mailcontents" rows="35" cols="115" class="input" readonly>
<html>
<head>
<title>핑거스 [theFingers] Membership Mail</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body bgcolor="#ffffff" text="#000000">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center">
		<img src="http://mailzine.10x10.co.kr/<% = omail.Fcode1 %>/<% = omail.Fimg1 %>" border="0" usemap="#ImgMap1"><br>
		<% if omail.Fimg2 <> "" then %>
		<img src="http://mailzine.10x10.co.kr/<% = omail.Fcode1 %>/<% = omail.Fimg2 %>" border="0" usemap="#ImgMap2">
		<% end if %>
		<% = omail.Fimgmap1 %>
		<% if omail.Fimg2 <> "" then %><% = omail.Fimgmap2 %><% end if %>
	</td>
</tr>
</table>
  <table width="600" border="0" cellspacing="0" cellpadding="5">
    <tr>
      <td width="17" valign="top"><font size="2" face="Verdana">1.</font></td>
      <td width="583"><font face="바탕" size="2" color="#000000">본 메일은 정보통신망 이용촉진
        및 정보보호 등에 관한 법률시행규칙에 의거<br>
        www.thefingers.co.kr (주)텐바이텐에 메일수신을 동의하셨기에 [광고]를<br>
        표시하지 않고 발송되는 발송전용메일입니다.</font></td>
    </tr>
    <tr>
      <td valign="top"><font size="2" face="Verdana">2.</font></td>
      <td><font face="바탕" size="2" color="#000000">수신을 원치 않으시는 분은 번거로우시겠지만 <a href="http://thefingers.co.kr/myfingers/membermodify.asp" target="_blank" onFocus='this.blur()'>
        홈페이지 &gt; MY Fingers &gt; 개인정보수정</a> 에서 <br>
        이메일 수신여부를 체크하여 주시기 바랍니다.</font></td>
    </tr>
  </table>
  <p><font face="바탕" size="2" color="#000000"><br>
    <br>
    </font></p>
</div>
</body>
</html>
</textarea>
	</td>
</tr>
</table>
<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
