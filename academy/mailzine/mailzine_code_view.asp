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
<title>�ΰŽ� [theFingers] Membership Mail</title>
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
      <td width="583"><font face="����" size="2" color="#000000">�� ������ ������Ÿ� �̿�����
        �� ������ȣ � ���� ���������Ģ�� �ǰ�<br>
        www.thefingers.co.kr (��)�ٹ����ٿ� ���ϼ����� �����ϼ̱⿡ [����]��<br>
        ǥ������ �ʰ� �߼۵Ǵ� �߼���������Դϴ�.</font></td>
    </tr>
    <tr>
      <td valign="top"><font size="2" face="Verdana">2.</font></td>
      <td><font face="����" size="2" color="#000000">������ ��ġ �����ô� ���� ���ŷο�ð����� <a href="http://thefingers.co.kr/myfingers/membermodify.asp" target="_blank" onFocus='this.blur()'>
        Ȩ������ &gt; MY Fingers &gt; ������������</a> ���� <br>
        �̸��� ���ſ��θ� üũ�Ͽ� �ֽñ� �ٶ��ϴ�.</font></td>
    </tr>
  </table>
  <p><font face="����" size="2" color="#000000"><br>
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
