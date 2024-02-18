<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 온라인 1:1 게시판 문의 보기
' Hieditor : 2010.01.03 한용민 온라인 이동 수정/생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim mailcontent ,boardqna ,boarditem ,id, mode ,username, title, contents, regdate ,qaDiv
dim replyuser, replytitle, replycontents, replydate ,email, emailok, extsitename ,sql , shopid
	id = request("id")
	mode = request("mode")
	username = request("username")
	title = request("title")
	contents = db2html(request("contents"))
	regdate = request("regdate")
	replyuser = request("replyuser")
	replytitle = request("replytitle")
	replycontents = request("replycontents")
	replydate = request("replydate")
	qaDiv	= req("qaDiv","")	' 유형
	email = request("email")
	shopid = request("shopid")
	'emailok = "Y"
	extsitename = request("extsitename")

if (mode = "reply") then
    set boardqna = New CMyQNA
    set boarditem = new CMyQNAItem

    boarditem.id = id
    boarditem.replyuser = replyuser
    boarditem.replytitle = html2db(replytitle)
    boarditem.replycontents = html2db(replycontents)

    boardqna.reply(boarditem)

    '2007 리뉴얼부터 무조건 답변메일 발송
    'if (emailok = "Y") then
            mailcontent = "<html>"
            mailcontent = mailcontent + "<head>"
            mailcontent = mailcontent + "<title>QnA</title>"
            mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>"
            mailcontent = mailcontent + "<link href='http://www.10x10.co.kr/css/2007ten.css' rel='stylesheet' type='text/css'>"
            mailcontent = mailcontent + "</head>"
            mailcontent = mailcontent + "<body>"
            mailcontent = mailcontent + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td height='210' valign='bottom'>"
            mailcontent = mailcontent + "        <table width='100%' height='210'  border='0' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td width='402' align='left' valign='top'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_top.gif' width='402' height='170' border='0' usemap='#Map'></td>"
            mailcontent = mailcontent + "            <td rowspan='2' align='left' valign='top'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_top2.gif' width='198' height='210'></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td height='40' align='center' valign='top' class='black12px'>"
            mailcontent = mailcontent + "                <table width='100%' height='40' border='0' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td width='11' align='center'><img src='http://fiximage.10x10.co.kr/web2007/email/side_line.gif' width='11' height='40'></td>"
            mailcontent = mailcontent + "                    <td align='center' valign='top' class='black12px'>" + username + "님이 문의하신 1:1상담내용에 대한 답변메일입니다.</td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        </table>"
            mailcontent = mailcontent + "    </td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td align='center' valign='top' bgcolor='#FF6C00' style='padding:10 0 10 0'>"
            mailcontent = mailcontent + "        <table width='578'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td style='padding:0 0 7 16'><font color='#FFFFFF'>아래의 답변은 마이텐바이텐 <a href='http://www.10x10.co.kr/cscenter/qna/myqnalist.asp' target='_blank' class='link_title'><strong>1대1상담하기</strong></a>에서도 확인가능합니다.</font></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "		    <td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "		         <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m01.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        문의일시 : <span class='black12px'>" + regdate + "</span><br>"
            mailcontent = mailcontent + "                        <br> " + title + "<br><br>"
            mailcontent = mailcontent + "                        <br> " + contents + ""
            mailcontent = mailcontent + "                    </td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_round_down.gif' width='578' height='4'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "    		<td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "    		    <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m02.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        <b>답변일시 :</b>" + replydate + "<br>"
            mailcontent = mailcontent + "                         " + html2db(replytitle) + "<br><br>"
            mailcontent = mailcontent + "                         " + nl2br(db2html(replycontents)) +""
            mailcontent = mailcontent + "                    </td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='center' bgcolor='#FFFFFF' style='padding-bottom:8'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_b_n.gif' width='536' height='54'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_round_down.gif' width='578' height='4'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        </table>"
            mailcontent = mailcontent + "    </td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td><img src='http://fiximage.10x10.co.kr/web2007/email/bottom.jpg' width='600' height='134' border='0' usemap='#Map2'></td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "</table>"
            mailcontent = mailcontent + "<map name='Map'><area shape='rect' coords='2,2,160,54' href='http://www.10x10.co.kr' target='_blank' onFocus='this.blur();'></map>"
            mailcontent = mailcontent + "<map name='Map2'><area shape='rect' coords='389,33,495,57' href='http://www.10x10.co.kr/cscenter/csmain.asp' target='_blank' onFocus='this.blur();'></map>"
            mailcontent = mailcontent + "</body>"
            mailcontent = mailcontent + "</html>"


            call SendMail("customer@10x10.co.kr", email, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)

        response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
    'end if

    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"

elseif (mode = "firstreply") then

    set boardqna = New CMyQNA
    set boarditem = new CMyQNAItem

	boardqna.read id
	if (boardqna.results(0).replyuser<>"") then
		response.write "<script>alert('이미 답변이 된 내용입니다.');</script>"
		response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if

    boarditem.id = id
    boarditem.replyuser = replyuser
    boarditem.replytitle = html2db(replytitle)
    boarditem.replycontents = html2db(replycontents)

    boardqna.reply(boarditem)

    '2007 리뉴얼부터 무조건 답변메일 발송
    'if (emailok = "Y") then
            mailcontent = "<html>"
            mailcontent = mailcontent + "<head>"
            mailcontent = mailcontent + "<title>QnA</title>"
            mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>"
            mailcontent = mailcontent + "<link href='http://www.10x10.co.kr/css/2007ten.css' rel='stylesheet' type='text/css'>"
            mailcontent = mailcontent + "</head>"
            mailcontent = mailcontent + "<body>"
            mailcontent = mailcontent + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td height='210' valign='bottom'>"
            mailcontent = mailcontent + "        <table width='100%' height='210'  border='0' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td width='402' align='left' valign='top'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_top.gif' width='402' height='170' border='0' usemap='#Map'></td>"
            mailcontent = mailcontent + "            <td rowspan='2' align='left' valign='top'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_top2.gif' width='198' height='210'></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td height='40' align='center' valign='top' class='black12px'>"
            mailcontent = mailcontent + "                <table width='100%' height='40' border='0' cellpadding='0' cellspacing='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td width='11' align='center'><img src='http://fiximage.10x10.co.kr/web2007/email/side_line.gif' width='11' height='40'></td>"
            mailcontent = mailcontent + "                    <td align='center' valign='top' class='black12px'>" + username + "님이 문의하신 1:1상담내용에 대한 답변메일입니다.</td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        </table>"
            mailcontent = mailcontent + "    </td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td align='center' valign='top' bgcolor='#FF6C00' style='padding:10 0 10 0'>"
            mailcontent = mailcontent + "        <table width='578'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "            <td style='padding:0 0 7 16'><font color='#FFFFFF'>아래의 답변은 마이텐바이텐 <a href='http://www.10x10.co.kr/cscenter/qna/myqnalist.asp' target='_blank' class='link_title'><strong>1대1상담하기</strong></a>에서도 확인가능합니다.</font></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "		    <td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "		         <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m01.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        문의일시 : <span class='black12px'>" + regdate + "</span><br>"
            mailcontent = mailcontent + "                        " + title + "<br><br>"
            mailcontent = mailcontent + "                        " + contents + ""
            mailcontent = mailcontent + "                    </td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_round_down.gif' width='578' height='4'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "    		<td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "    		    <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m02.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        <b>답변일시 :</b>" + replydate + "<br>"
            mailcontent = mailcontent + "                         " + html2db(replytitle) + "<br><br>"
            mailcontent = mailcontent + "                         " + nl2br(db2html(replycontents)) +""
            mailcontent = mailcontent + "                    </td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='center' bgcolor='#FFFFFF' style='padding-bottom:8'><img src='http://fiximage.10x10.co.kr/web2007/email/qna_b_n.gif' width='536' height='54'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_round_down.gif' width='578' height='4'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                </table>"
            mailcontent = mailcontent + "            </td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        </table>"
            mailcontent = mailcontent + "    </td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "<tr>"
            mailcontent = mailcontent + "    <td><img src='http://fiximage.10x10.co.kr/web2007/email/bottom.jpg' width='600' height='134' border='0' usemap='#Map2'></td>"
            mailcontent = mailcontent + "</tr>"
            mailcontent = mailcontent + "</table>"
            mailcontent = mailcontent + "<map name='Map'><area shape='rect' coords='2,2,160,54' href='http://www.10x10.co.kr' target='_blank' onFocus='this.blur();'></map>"
            mailcontent = mailcontent + "<map name='Map2'><area shape='rect' coords='389,33,495,57' href='http://www.10x10.co.kr/cscenter/csmain.asp' target='_blank' onFocus='this.blur();'></map>"
            mailcontent = mailcontent + "</body>"
            mailcontent = mailcontent + "</html>"


            call SendMail("customer@10x10.co.kr", email, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)

        response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
    'end if

    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"

elseif  (mode = "del") then
	
    sql = "update [db_cs].[10x10].tbl_myqna " + VbCRlf
    sql = sql + " set isusing = 'N'" + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"
    'response.write sql
    'dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
	response.write "<script>location.replace('board/itemqna_list.asp')</script>"
end if

'' 답변,유형수정 2009년 4월 리뉴얼 버전
IF mode="REP" Or mode = "CHG" Then

	set boardqna = New CMyQNA

	boardqna.read ""	'' 초기값으로 세팅
	boardqna.results(0).id = id
	boardqna.results(0).qaDiv = qaDiv
	boardqna.results(0).replyuser = replyuser
	boardqna.results(0).replytitle = replytitle
	boardqna.results(0).replycontents = replycontents
    boardqna.BackProcData(mode)
    
    set boardqna = nothing

End If 

'//해당 매장 지정
if mode = "chshopid" then
	
	sql = "update db_cs.dbo.tbl_myqna set" + vbcrlf
	sql = sql & " shopid = '"& shopid &"'" + vbcrlf
	sql = sql & " where id = '"& id &"'"
	
	'response.write sql &"<br>"
	dbget.execute sql

	response.write "<script>alert('OK')</script>"
    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
end if

IF mode="CHG" Then
	response.write "<script>alert('수정되었습니다')</script>"
    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
End If 

'' 답변메일발송 2009년 4월 리뉴얼 버전
IF mode="REP" Then

	dim oMail
	dim MailHTML
	dim MailTypeNo

	set oMail = New MailCls

	oMail.MailType = 15 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles = "[텐바이텐]" & username & "님께서 문의하신 내용에 대한 답변입니다."  '"즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]"
	'oMail.SenderMail = "customer@10x10.co.kr"
	'oMail.SenderNm = "텐바이텐"

	oMail.AddrType = "string"
	oMail.ReceiverNm = username
	oMail.ReceiverMail = email

	MailHTML = oMail.getMailTemplate()

	IF MailHTML="" Then
		response.write "<script>alert('메일발송이 실패 하였습니다.')</script>"
    	response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	End IF

	MailHTML =replace(MailHTML,"[$USER_NAME$]",oMail.ReceiverNm)
	MailHTML =replace(MailHTML,"[$QUESTION_TIME$]",regdate)
	MailHTML =replace(MailHTML,"[$QUESTION_CONTENTS$]","<b>"& server.HTMLEncode(title) & "</b><br><br>"& nl2br(server.HTMLEncode(db2html(contents))))
	MailHTML =replace(MailHTML,"[$ANSWER_TIME$]",replydate)
	MailHTML =replace(MailHTML,"[$ANSWER_CONTENTS$]","<b>"& server.HTMLEncode(replytitle) &"</b><br><br>"& nl2br(server.HTMLEncode(db2html(replycontents))))
	MailHTML =replace(MailHTML,"[$ANSWER_NOTICE$]","")
	MailHTML =replace(MailHTML,"[$KEYVAL$]",MD5(id))

	oMail.MailConts = MailHTML
    oMail.MailerMailGubun = 1		' 메일러 자동메일 번호
    oMail.Send_TMSMailer()		'TMS메일러
	'oMail.Send_Mailer()
	'oMail.Send_CDO()
	'oMail.Send_CDONT()

	set oMail = nothing
	response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"

End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
