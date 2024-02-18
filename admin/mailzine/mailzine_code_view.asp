<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->

<%
dim idx, code, omail ,yyyy1, mm1, dd1 , tmp , area
dim title,regdate,img1,img2,img3,img4,imgmap1,imgmap2,imgmap3,imgmap4,isusing,gubun

idx = request("idx")

set omail = new CMailzineList
	omail.frectidx = idx
	''omail.frectmailergubun = "EMS"
	
	'//idx ���� ������쿡�� ����
	if idx <> "" then
		omail.MailzineDetail()

		if omail.ftotalcount > 0 then
			title = omail.foneitem.ftitle
			regdate = omail.foneitem.fregdate
			img1 = omail.foneitem.fimg1
			img2 = omail.foneitem.fimg2
			img3 = omail.foneitem.fimg3
			img4 = omail.foneitem.fimg4
			imgmap1 = omail.foneitem.fimgmap1
			imgmap2 = omail.foneitem.fimgmap2
			imgmap3 = omail.foneitem.fimgmap3
			imgmap4 = omail.foneitem.fimgmap4
			isusing = omail.foneitem.fisusing
			gubun = omail.foneitem.fgubun
			area = omail.foneitem.farea
			'�ΰŽ�_������ : finger_all 
			'�ΰŽ�_������ : finger_metropolitan

			tmp = split(omail.foneitem.fregdate,".")
			yyyy1 = tmp(0)
			mm1 = tmp(1)
			dd1 = tmp(2)
			code = mm1 & dd1
		end if
	end if
%>
<font color="red">�� �ڵ� ����</font>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="FFFFFF">
	<td>
		<input type="text" name="title" size="100" class="input" readonly value="(����) <% = title %>"><br>
		<textarea name="mailcontents" rows="35" cols="115" class="input" readonly>
<html>
<head>
<title>10X10 [tenbyten] Membership Mail</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<meta name="viewport" content="width=700, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0" />
</head>

<body bgcolor="#ffffff" text="#000000">
<div align="center">
<table align="center" width="700" border="0" cellpadding="0" cellspacing="0">
<tr>
<td align="center" style="line-height:0; font-size:0;"><img src="<%=mailzine%>/<% = yyyy1 %>/<% = img1 %>" border="0" usemap="#ImgMap1"></td>
</tr>
<% If img2 <> "" Then %>
<tr>
<td align="center" style="line-height:0; font-size:0;"><img src="<%=mailzine%>/<% = yyyy1 %>/<% = img2 %>" border="0" usemap="#ImgMap2"></td>
</tr>
<% End If %>
<% If img3 <> "" Then %>
<tr>
<td align="center" style="line-height:0; font-size:0;"><img src="<%=mailzine%>/<% = yyyy1 %>/<% = img3 %>" border="0" usemap="#ImgMap3"></td>
</tr>
<% End If %>
<% If img4 <> "" Then %>
<tr>
<td align="center" style="line-height:0; font-size:0;"><img src="<%=mailzine%>/<% = yyyy1 %>/<% = img4 %>" border="0" usemap="#ImgMap4"></td>
</tr>
<% End If %>
<% = replace(imgmap1,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
<% If img2 <> "" Then %>
<% = replace(imgmap2,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
<% End If %>
<% If img3 <> "" Then %>
<% = replace(imgmap3,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
<% End If %>
<% If img4 <> "" Then %>
<% = replace(imgmap4,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
<% End If %>
</table>
<table align="center" width="670" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td style="padding-top:30px; text-align:center;">
	  <% if area = "finger_all" or area = "finger_metropolitan" then %>
	      <a href="https://www.facebook.com/thefingers.co.kr/" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_facebook.gif" alt="Facebook" style="border:none;"/></a>
	      <a href="https://www.instagram.com/thefingers.co.kr/" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_instargram.gif" alt="Instargram" style="border:none" /></a>
	      <a href="https://www.youtube.com/user/fingersacademy/" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_yutube.png" alt="yutube" style="border:none" /></a>
	  <% else %>
	      <a href="http://twitter.com/your10x10" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_twitter.gif" alt="Twitter" style="border:none" /></a>
	      <a href="http://www.facebook.com/your10x10" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_facebook.gif" alt="Facebook" style="border:none;"/></a>
	      <a href="http://www.instagram.com/your10x10/" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_instargram.gif" alt="Instargram" style="border:none" /></a>
	      <a href="https://www.pinterest.com/your10x10/" target="_blank"><img src="http://fiximage.10x10.co.kr/web2013/common/footer_sns_pinterest.gif" alt="pinterest" style="border:none" /></a>
	  <% end if %>
    </td>
  </tr>
  <tr>
	<td align="left" style="padding:25px 0 0 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;">* �� ������ ������Ÿ� �̿����� �� ������ȣ � ���� ���������Ģ�� �ǰ�&nbsp;<%= Year(Now) %>�� <%= month(Now) %>�� <%= day(Now) %>�� �������� ȸ������ ���ϼ��ſ� �����ϼ̱⿡ �߼۵Ǵ� �߼���������Դϴ�.<br />
	<% if area = "finger_all" or area = "finger_metropolitan" then %>
	* �� ������ �߽� �����̸� ȸ�� �� ������ ���� �� �����ϴ�.<br />
	<% end if %>
	* �� �̻� ������ ������ �����ø� ���Űź� ��ư�� Ŭ�����ּ���.<br />
	* ������������ �� �ý��� ���뿡 2~3�� �ݿ��ð��� �ҿ�� �� ������ �� �� ���� ��Ź�帳�ϴ�.<br />
	<a href='http://www.10x10.co.kr/member/mailzine/reject_mailzine.asp?M_ID=${TMS_M_EMAIL}' target='_blank' onFocus='this.blur()'>[���Űź�]</a> (To unsubscribe this e-mail, <a href='http://www.10x10.co.kr/member/mailzine/reject_mailzine.asp?M_ID=${TMS_M_EMAIL}' target='_blank' onFocus='this.blur()'>click HERE</a>)</td>
  </tr>
  <tr>
    <td align="center" style="padding:12px 0 20px 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;">
  	<% if area = "finger_all" or area = "finger_metropolitan" then %>
  	(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� ��ǥ�̻�:������<br>
  	����ڵ�Ϲ�ȣ : 211-87-00620 / ����Ǹž��Ű� : ��01-1968ȣ / ����������ȣ �� û�ҳ� ��ȣå���� : �̹���<br>
  	���ູ���� TEL : <strong>1644-1557</strong> / E-mail : <a href="mailto:customer@thefingers.co.kr" style="text-decoration:none; color:#333;">customer@thefingers.co.kr</a>
	<% else %>
  	(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� ��ǥ�̻�:������<br>
  	����ڵ�Ϲ�ȣ : 211-87-00620 / ����Ǹž��Ű� : ��01-1968ȣ / ����������ȣ �� û�ҳ� ��ȣå���� : �̹���<br>
  	���ູ���� TEL : <strong>1644-6030</strong> / E-mail : <a href="mailto:customer@10x10.co.kr" style="text-decoration:none; color:#333;">customer@10x10.co.kr</a>
	<% end if %>
    </td>
  </tr>
</table>
</div>
</body>
</html>
</textarea>
</td>
</tr>
</table>

<%
	set omail = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
