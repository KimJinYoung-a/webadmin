<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
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

	'//idx 값이 있을경우에만 쿼리
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

	'// 사이트구분
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

'2015년 1월 27일 변경(트위터,페북,인스타 아이콘 추가)-유태욱
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
response.write "    <td align='left' style='padding:25px 0 0 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;'>* 본 메일은 정보통신망 이용촉진 및 정보보호 등에 관한 법률시행규칙에 의거&nbsp;" & Year(Now) & "년" & month(Now) & "월" & day(Now) & "일 기준으로 회원님의 메일수신에 동의하셨기에 발송되는 발송전용메일입니다.<br />"
response.write "    * 더 이상 수신을 원하지 않으시면 수신거부 버튼을 클릭해주세요. <a href='http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=" & area & "' onFocus='blur()' target='_blink'>[비회원 메일링 수신거부]</a> (To unsubscribe this e-mail, <a href='http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=" & area & "' onFocus='blur()' target='_blink'>click HERE</a>)<br />"
response.write "    * 개인정보변경 시 시스템 적용에 2~3일 반영시간이 소요될 수 있으니 이 점 양해 부탁드립니다.</td>"
response.write "  </tr>"
response.write "<tr>"
response.write "  <td align='center' style='padding:25px 0 20px 0; font-family:Dotum; color:#888; font-size:11px; line-height:16px;'>"
if area = "fingers" then
response.write "	  (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 대표이사:최은희<br>"
response.write "      사업자등록번호 : 211-87-00620 / 통신판매업신고 : 제01-1968호 / 개인정보보호 및 청소년 보호책임자 : 이문재<br>"
response.write "      고객행복센터 TEL : <strong>1644-1557</strong> / E-mail : <a href='mailto:customer@thefingers.co.kr' style='text-decoration:none; color:#333;'>customer@thefingers.co.kr</a>"
else
response.write "	  (03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 대표이사:최은희<br>"
response.write "      사업자등록번호 : 211-87-00620 / 통신판매업신고 : 제01-1968호 / 개인정보보호 및 청소년 보호책임자 : 이문재<br>"
response.write "      고객행복센터 TEL : <strong>1644-6030</strong> / E-mail : <a href='mailto:customer@10x10.co.kr' style='text-decoration:none; color:#333;'>customer@10x10.co.kr</a>"
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
