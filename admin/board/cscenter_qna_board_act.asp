<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->
<%

dim mailcontent

dim boardqna
dim boarditem
dim id, mode, replytitle, replycontents, replyuser
dim email, emailok, extsitename

id = request("id")
mode = request("mode")
replytitle = request("replytitle")
replycontents = request("replycontents")

email = request("email")
emailok = request("emailok")
extsitename = request("extsitename")

if (mode = "reply") then
        set boardqna = New CMyQNA
        set boarditem = new CMyQNAItem

        boarditem.id = id
        boarditem.replyuser = "10x10"
        boarditem.replytitle = html2db(replytitle)
        boarditem.replycontents = html2db(replycontents)

        boardqna.reply(boarditem)

        if (emailok = "Y") then
        		if extsitename="maxmovie" then
        			mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>맥스샾 답변메일 </TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_maxshop01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_maxshop03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://maxshop.maxmovie.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://maxshop.maxmovie.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                call sendmail("giftshop@10x10.co.kr", email, "[맥스샾] 문의하신 내용에 대한 답변입니다. ", mailcontent)
	            elseif extsitename="way2way" then
        			mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>웨이투웨이 답변메일 </TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_way2way01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_way2way03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://www.way2way.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://www.way2way.com/cscenter/csmain.asp' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                call sendmail("guide@way2way.com", email, "[웨이투웨이] 문의하신 내용에 대한 답변입니다. ", mailcontent)
        		else
	                mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]</TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://fiximage.10x10.co.kr/mail/customer_mail_01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://www.10x10.co.kr' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://www.10x10.co.kr/my10x10/myqnalist.asp' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                call sendmail("customer@10x10.co.kr", email, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)
            	end if

                response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
        end if

        response.write "<script>location.replace('cscenter_qna_board_reply.asp?id=" + id + "')</script>"

elseif (mode = "firstreply") then

	    set boardqna = New CMyQNA
        set boarditem = new CMyQNAItem

		boardqna.read id
		if (boardqna.results(0).replyuser<>"") then
			response.write "<script>alert('이미 답변이 된 내용입니다.');</script>"
			response.write "<script>location.replace('cscenter_qna_board_reply.asp?id=" + id + "')</script>"
			dbget.close()	:	response.End
		end if

        boarditem.id = id
        boarditem.replyuser = "10x10"
        boarditem.replytitle = html2db(replytitle)
        boarditem.replycontents = html2db(replycontents)

        boardqna.reply(boarditem)

        if (emailok = "Y") then
        		if extsitename="maxmovie" then
        			mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>맥스샾 답변메일 </TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_maxshop01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_maxshop03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://maxshop.maxmovie.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://maxshop.maxmovie.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                call sendmail("giftshop@10x10.co.kr", email, "[맥스샾] 문의하신 내용에 대한 답변입니다. ", mailcontent)
	            elseif extsitename="way2way" then
        			mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>웨이투웨이 답변메일 </TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_way2way01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_way2way03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://www.way2way.com' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://www.way2way.com/cscenter/csmain.asp' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                call sendmail("guide@way2way.com", email, "[웨이투웨이] 문의하신 내용에 대한 답변입니다. ", mailcontent)
        		else
	                mailcontent = "<HTML>"
	                mailcontent = mailcontent + "<HEAD>"
	                mailcontent = mailcontent + "<TITLE>즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]</TITLE>"
	                mailcontent = mailcontent + "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=euc-kr'>"
	                mailcontent = mailcontent + "</HEAD>"
	                mailcontent = mailcontent + "<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr>"
	                mailcontent = mailcontent + "<td align='center' valign='top'>"
	                mailcontent = mailcontent + "<TABLE WIDTH=600 BORDER=0 CELLPADDING=0 CELLSPACING=0>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://fiximage.10x10.co.kr/mail/customer_mail_01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	                mailcontent = mailcontent + "<TR>"
	                mailcontent = mailcontent + "<TD align='center' valign='top'>"
	                mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	                mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "</TD>"
	                mailcontent = mailcontent + "</TR>"
	                mailcontent = mailcontent + "<TR><TD><IMG SRC='http://partner.10x10.co.kr/admin/board/images/customer_mail_03.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	                mailcontent = mailcontent + "</TABLE>"
	                mailcontent = mailcontent + "</td>"
	                mailcontent = mailcontent + "</tr>"
	                mailcontent = mailcontent + "</table>"
	                mailcontent = mailcontent + "<map name='Map'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://www.10x10.co.kr' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "<map name='Map2'>"
	                mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://www.10x10.co.kr/my10x10/myqnalist.asp' target='_blank'>"
	                mailcontent = mailcontent + "</map>"
	                mailcontent = mailcontent + "</BODY>"
	                mailcontent = mailcontent + "</HTML>"

	                'mailcontent = "안녕하세요. 텐바이텐입니다.<br>"
	                'mailcontent = mailcontent + "답변 내용은 아래와 같습니다."
	                'mailcontent = mailcontent + "<hr>"
	                'mailcontent = mailcontent + db2html(replycontents)

	                call sendmail("customer@10x10.co.kr", email, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)
	            end if
                response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
        end if

        response.write "<script>location.replace('cscenter_qna_board_reply.asp?id=" + id + "')</script>"
elseif  (mode = "del") then

                dim sql

                sql = "update [db_cs].[10x10].tbl_myqna " + VbCRlf
                sql = sql + " set isusing = 'N'" + VbCRlf
                sql = sql + " where id = '" + Cstr(id) + "'"
                'response.write sql
                'dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
        response.write "<script>location.replace('board/itemqna_list.asp')</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->