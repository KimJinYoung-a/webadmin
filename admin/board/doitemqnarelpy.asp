<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 게시판관리>>[SITE]고객상품문의
' Hieditor : 최초 생성자 모름
'			 2017.05.19 한용민 수정(이메일발송수정. 고객이 선택한것과 상관없이 다 쏘게 되어 있었음)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim mailcontent, itemqna, boarditem, id, mode, replycontents, replyuser, usermail, emailok, extsitename 
dim page,makerid,notupbea,mifinish,research, sql
	id = request("id")
	mode = request("mode")
	replycontents = html2db(request("replycontents"))
	replyuser = session("ssBctId")
	usermail = request("usermail")
	emailok = request("emailok")
	extsitename = request("extsitename")
	page=request("page")
	makerid=request("makerid")
	notupbea=request("notupbea")
	mifinish=request("mifinish")
	research=request("research")

set itemqna = new CItemQna
	itemqna.FRectID = id
	itemqna.getOneItemQna

if (mode = "firstreply") then
	if Not IsNULL(itemqna.FOneItem.Freplydate) then
		response.write "<script type='text/javascript'>alert('이미 답변이 된 내용입니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if
end if

if (mode = "reply") or (mode = "firstreply") then
	if id="" then
		response.write "<script type='text/javascript'>alert('구분자가 없습니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('newitemqna_view.asp?id=" + id + "')</script>"
		dbget.close()	:	response.End
	end if

	sql = "update [db_cs].[dbo].tbl_my_item_qna " + VbCRlf
    sql = sql + " set replycontents = '" + replycontents + "'" + VbCRlf
    sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
    sql = sql + " , replydate = getdate()" + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"

    rsget.Open sql, dbget, 1

    'if (emailok = "Y") then
    '/답변 안된거만
    if IsNULL(itemqna.FOneItem.Freplydate) then
    	'/답변 이메일 수신여부가 Y 인거만
	    if itemqna.FOneItem.Femailok = "Y" then
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
	        mailcontent = mailcontent + "<TR><TD><IMG SRC='http://webadmin.10x10.co.kr/admin/board/images/customer_mail_01.gif' ALT='' WIDTH=600 HEIGHT=114 border='0' usemap='#Map'></TD></TR>"
	        mailcontent = mailcontent + "<TR>"
	        mailcontent = mailcontent + "<TD align='center' valign='top'>"
	        mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	        mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>[질문내용]</font></td></tr>"
	        mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(itemqna.FOneItem.Fcontents)) + "</font></td></tr>"
	        mailcontent = mailcontent + "</table><br>"
	        mailcontent = mailcontent + "<table width='100%' border='0' cellspacing='30' cellpadding='0'>"
	        mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>[답변내용]</font></td></tr>"
	        mailcontent = mailcontent + "<tr><td><font size='2' face='바탕'>" + nl2br(db2html(replycontents)) + "</font></td></tr>"
	        mailcontent = mailcontent + "</table>"
	        mailcontent = mailcontent + "</TD>"
	        mailcontent = mailcontent + "</TR>"
	        mailcontent = mailcontent + "<TR><TD><IMG SRC='http://webadmin.10x10.co.kr/admin/board/images/customer_mail_04.gif' ALT='' WIDTH=600 HEIGHT=89 border='0' usemap='#Map2'></TD></TR>"
	        mailcontent = mailcontent + "</TABLE>"
	        mailcontent = mailcontent + "</td>"
	        mailcontent = mailcontent + "</tr>"
	        mailcontent = mailcontent + "</table>"
	        mailcontent = mailcontent + "<map name='Map'>"
	        mailcontent = mailcontent + "<area shape='rect' coords='12,11,579,50' href='http://www.10x10.co.kr' target='_blank'>"
	        mailcontent = mailcontent + "</map>"
	        mailcontent = mailcontent + "<map name='Map2'>"
	        mailcontent = mailcontent + "<area shape='rect' coords='234,19,354,40' href='http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + CStr(itemqna.FOneItem.FItemID) + "' target='_blank'>"
	        mailcontent = mailcontent + "</map>"
	        mailcontent = mailcontent + "</BODY>"
	        mailcontent = mailcontent + "</HTML>"

	        call sendmail("customer@10x10.co.kr", usermail, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)
	        response.write "<script type='text/javascript'>alert('답변메일이 발송되었습니다.')</script>"
		end if
    end if

  response.write "<script type='text/javascript'>location.replace('newitemqna_view.asp?id=" + id + "&page=" + page + "&makerid=" + makerid + "&notupbea=" + notupbea + "&mifinish="+ mifinish + "&research=" + research + "')</script>"					

elseif  (mode = "del") then
    sql = "update [db_cs].[dbo].tbl_my_item_qna " + VbCRlf
    sql = sql + " set isusing = 'N'" + VbCRlf
    sql = sql + " , replyuser = '" + replyuser + "'" + VbCRlf
    sql = sql + " , replydate = getdate()" + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"

    rsget.Open sql, dbget, 1
    response.write "<script type='text/javascript'>location.replace('newitemqna_list.asp?page=" + page + "&makerid=" + makerid + "&notupbea=" + notupbea + "&mifinish="+ mifinish + "&research=" + research + "')</script>"
end if

set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->