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
			
			tmp = split(omail.foneitem.fregdate,".")
			yyyy1 = tmp(0)
			mm1 = tmp(1)
			dd1 = tmp(2)
			code = mm1 & dd1			
		end if		
	end if

	'// ����Ʈ����
	Select Case area
		Case "ten_all", "ten_metropolitan"
			area = "10x10"
		Case "finger_all", "finger_metropolitan"
			area = "fingers"
	End Select

response.write "<html>"
response.write "<head>"
response.write "<title>10X10 [tenbyten] Membership Mail</title>"
response.write "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>"
response.write "</head>"

response.write "<body>"
response.write "<div align='center'>"
response.write "<table align='center' width=700 border=0 cellspacing=0 cellpadding=0>"
response.write "<tr>"
response.write "<td align='center'>"
response.write "<img src='"&mailzine&"/"& yyyy1 &"/"&img1 &"' border=0 usemap='#ImgMap1'>"
if img2 <> "" then
	response.write "<br><img src='"&mailzine&"/"& yyyy1 &"/"&img2 &"' border=0 usemap='#ImgMap2'>"
end if
if img3 <> "" then
	response.write "<br><img src='"&mailzine&"/"& yyyy1 &"/"&img3 &"' border=0 usemap='#ImgMap3'>"
end if
if img4 <> "" then
	response.write "<br><img src='"&mailzine&"/"& yyyy1 &"/"&img4 &"' border=0 usemap='#ImgMap4'>"
end if
%>
<% = replace(imgmap1,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
	<% if img2 <> "" then %>
	<% = replace(imgmap2,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
	<% end if %>
	<% if img3 <> "" then %>
	<% = replace(imgmap3,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
	<% end if %>
	<% if img4 <> "" then %>
	<% = replace(imgmap4,"target=" + Chr(34) + "_top" + Chr(34) ,"target=" + Chr(34) + "_blank" + Chr(34)) %>
	<% end if %>
<%
response.write "</td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align='center' style='padding:28px 0 28px 0'><a href='http://www.10x10.co.kr/member/join.asp' onFocus='blur()'><img src='http://fiximage.10x10.co.kr/web2009/mailing/bemail_join2.gif' border=0></a></td>"
response.write "</tr>"
response.write "<tr>"
response.write "<td align='center'><img src='http://fiximage.10x10.co.kr/web2009/member/memberadvtg.gif' width=690 height=299></td>"
response.write "</tr>"
'response.write "<tr>"
'response.write "<td style='padding-top:17px; padding-left:7px'><img src='http://fiximage.10x10.co.kr/web2009/mailing/bemail_copy.gif' width=457 height=30></td>"
'response.write "</tr>"
'response.write "<tr>"
'response.write "<td style='padding-top:13px; padding-left:7px'><a href='http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=" & area & "' onFocus='blur()' target='_blink'><img src='http://fiximage.10x10.co.kr/web2009/mailing/bemail_cancel_btn.gif' border=0></a></td>"
'response.write "</tr>"

'2015�� 1�� 27�� ����(Ʈ����,���,�ν�Ÿ ������ �߰�)-���¿�
response.write "<tr>"
response.write "  <td style='padding-top:30px; text-align:center;'>"
if area = "fingers" then
	response.write "      <a href='http://www.facebook.com/thefingers.co.kr/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_facebook.gif' alt='Facebook' style='border:none;'/></a>"
	response.write "      <a href='http://www.instagram.com/thefingers.co.kr/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_instargram.gif' alt='Instargram' style='border:none' /></a>"
	response.write "	  <a href='https://www.youtube.com/user/fingersacademy/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_yutube.png' alt='yutube' style='border:none' /></a>"
else
	response.write "      <a href='http://twitter.com/your10x10' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_twitter.gif' alt='Twitter' style='border:none' /></a>"
	response.write "      <a href='http://www.facebook.com/your10x10' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_facebook.gif' alt='Facebook' style='border:none;'/></a>"
	response.write "      <a href='http://www.instagram.com/your10x10/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_instargram.gif' alt='Instargram' style='border:none' /></a>"
	response.write "	  <a href='https://www.pinterest.com/your10x10/' target='_blank'><img src='http://fiximage.10x10.co.kr/web2013/common/footer_sns_pinterest.gif' alt='pinterest' style='border:none' /></a>"
end if
response.write "  </td>"
response.write "</tr>"
response.write "  <tr>"
response.write "    <td align='left' style='padding:25px 0 0 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;'>* �� ������ ������Ÿ� �̿����� �� ������ȣ � ���� ���������Ģ�� �ǰ�&nbsp;" & Year(Now) & "��" & month(Now) & "��" & day(Now) & "�� �������� ȸ������ ���ϼ��ſ� �����ϼ̱⿡ �߼۵Ǵ� �߼���������Դϴ�.<br />"
response.write "    * �� �̻� ������ ������ �����ø� ���Űź� ��ư�� Ŭ�����ּ���. <a href='http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=" & area & "' onFocus='blur()' target='_blink'>[��ȸ�� ���ϸ� ���Űź�]</a> (To unsubscribe this e-mail, <a href='http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=" & area & "' onFocus='blur()' target='_blink'>click HERE</a>)<br />"
response.write "    * ������������ �� �ý��� ���뿡 2~3�� �ݿ��ð��� �ҿ�� �� ������ �� �� ���� ��Ź�帳�ϴ�.</td>"
response.write "  </tr>"
response.write "<tr>"
response.write "  <td align='center' style='padding:25px 0 20px 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;'>"
if area = "fingers" then
response.write "	  (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� ��ǥ�̻�:������<br>"
response.write "      ����ڵ�Ϲ�ȣ : 211-87-00620 / ����Ǹž��Ű� : ��01-1968ȣ / ����������ȣ �� û�ҳ� ��ȣå���� : �̹���<br>"
response.write "      ���ູ���� TEL : <strong>1644-1557</strong> / E-mail : <a href='mailto:customer@thefingers.co.kr' style='text-decoration:none; color:#333;'>customer@thefingers.co.kr</a>"
else
response.write "	  (03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� ��ǥ�̻�:������<br>"
response.write "      ����ڵ�Ϲ�ȣ : 211-87-00620 / ����Ǹž��Ű� : ��01-1968ȣ / ����������ȣ �� û�ҳ� ��ȣå���� : �̹���<br>"
response.write "      ���ູ���� TEL : <strong>1644-6030</strong> / E-mail : <a href='mailto:customer@10x10.co.kr' style='text-decoration:none; color:#333;'>customer@10x10.co.kr</a>"
end if
response.write "  </td>"
response.write "</tr>"
response.write "</table>"
response.write "<br>"
response.write "</div>"
response.write "</body>"
response.write "</html>"
%>
<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
