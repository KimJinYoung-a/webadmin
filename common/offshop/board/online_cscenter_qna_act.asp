<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �¶��� 1:1 �Խ��� ���� ����
' Hieditor : 2010.01.03 �ѿ�� �¶��� �̵� ����/����
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
	qaDiv	= req("qaDiv","")	' ����
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

    '2007 ��������� ������ �亯���� �߼�
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
            mailcontent = mailcontent + "                    <td align='center' valign='top' class='black12px'>" + username + "���� �����Ͻ� 1:1��㳻�뿡 ���� �亯�����Դϴ�.</td>"
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
            mailcontent = mailcontent + "            <td style='padding:0 0 7 16'><font color='#FFFFFF'>�Ʒ��� �亯�� �����ٹ����� <a href='http://www.10x10.co.kr/cscenter/qna/myqnalist.asp' target='_blank' class='link_title'><strong>1��1����ϱ�</strong></a>������ Ȯ�ΰ����մϴ�.</font></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "		    <td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "		         <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m01.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        �����Ͻ� : <span class='black12px'>" + regdate + "</span><br>"
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
            mailcontent = mailcontent + "                        <b>�亯�Ͻ� :</b>" + replydate + "<br>"
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


            call SendMail("customer@10x10.co.kr", email, "��ſ��� ������ ���θ�, �ٹ����� [10X10=tenbyten]", mailcontent)

        response.write "<script>alert('�亯������ �߼۵Ǿ����ϴ�.')</script>"
    'end if

    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"

elseif (mode = "firstreply") then

    set boardqna = New CMyQNA
    set boarditem = new CMyQNAItem

	boardqna.read id
	if (boardqna.results(0).replyuser<>"") then
		response.write "<script>alert('�̹� �亯�� �� �����Դϴ�.');</script>"
		response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if

    boarditem.id = id
    boarditem.replyuser = replyuser
    boarditem.replytitle = html2db(replytitle)
    boarditem.replycontents = html2db(replycontents)

    boardqna.reply(boarditem)

    '2007 ��������� ������ �亯���� �߼�
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
            mailcontent = mailcontent + "                    <td align='center' valign='top' class='black12px'>" + username + "���� �����Ͻ� 1:1��㳻�뿡 ���� �亯�����Դϴ�.</td>"
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
            mailcontent = mailcontent + "            <td style='padding:0 0 7 16'><font color='#FFFFFF'>�Ʒ��� �亯�� �����ٹ����� <a href='http://www.10x10.co.kr/cscenter/qna/myqnalist.asp' target='_blank' class='link_title'><strong>1��1����ϱ�</strong></a>������ Ȯ�ΰ����մϴ�.</font></td>"
            mailcontent = mailcontent + "        </tr>"
            mailcontent = mailcontent + "        <tr>"
            mailcontent = mailcontent + "		    <td style='padding-bottom:10 '>"
            mailcontent = mailcontent + "		         <table width='100%'  border='0' cellspacing='0' cellpadding='0'>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td><img src='http://fiximage.10x10.co.kr/web2007/email/qna_m01.jpg' width='578' height='41'></td>"
            mailcontent = mailcontent + "                </tr>"
            mailcontent = mailcontent + "                <tr>"
            mailcontent = mailcontent + "                    <td align='left' bgcolor='#FFFFFF' style='padding:20 35 20 35'>"
            mailcontent = mailcontent + "                        �����Ͻ� : <span class='black12px'>" + regdate + "</span><br>"
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
            mailcontent = mailcontent + "                        <b>�亯�Ͻ� :</b>" + replydate + "<br>"
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


            call SendMail("customer@10x10.co.kr", email, "��ſ��� ������ ���θ�, �ٹ����� [10X10=tenbyten]", mailcontent)

        response.write "<script>alert('�亯������ �߼۵Ǿ����ϴ�.')</script>"
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

'' �亯,�������� 2009�� 4�� ������ ����
IF mode="REP" Or mode = "CHG" Then

	set boardqna = New CMyQNA

	boardqna.read ""	'' �ʱⰪ���� ����
	boardqna.results(0).id = id
	boardqna.results(0).qaDiv = qaDiv
	boardqna.results(0).replyuser = replyuser
	boardqna.results(0).replytitle = replytitle
	boardqna.results(0).replycontents = replycontents
    boardqna.BackProcData(mode)
    
    set boardqna = nothing

End If 

'//�ش� ���� ����
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
	response.write "<script>alert('�����Ǿ����ϴ�')</script>"
    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"
End If 

'' �亯���Ϲ߼� 2009�� 4�� ������ ����
IF mode="REP" Then

	dim oMail
	dim MailHTML
	dim MailTypeNo

	set oMail = New MailCls

	oMail.MailType = 15 '���� ������ ������ (mailLib2.asp ����)
	oMail.MailTitles = "[�ٹ�����]" & username & "�Բ��� �����Ͻ� ���뿡 ���� �亯�Դϴ�."  '"��ſ��� ������ ���θ�, �ٹ����� [10X10=tenbyten]"
	'oMail.SenderMail = "customer@10x10.co.kr"
	'oMail.SenderNm = "�ٹ�����"

	oMail.AddrType = "string"
	oMail.ReceiverNm = username
	oMail.ReceiverMail = email

	MailHTML = oMail.getMailTemplate()

	IF MailHTML="" Then
		response.write "<script>alert('���Ϲ߼��� ���� �Ͽ����ϴ�.')</script>"
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
    oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
    oMail.Send_TMSMailer()		'TMS���Ϸ�
	'oMail.Send_Mailer()
	'oMail.Send_CDO()
	'oMail.Send_CDONT()

	set oMail = nothing
	response.write "<script>alert('�亯������ �߼۵Ǿ����ϴ�.')</script>"
    response.write "<script>location.replace('online_cscenter_qna_reply.asp?id=" + id + "')</script>"

End if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
