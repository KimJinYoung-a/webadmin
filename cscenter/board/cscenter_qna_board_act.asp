<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 이상구 생성
'			2021.09.13 한용민 수정(알림톡 텐바이텐고객센터로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/util/scm_myalarmlib.asp" -->

<!-- #include virtual="/lib/email/mailLib2.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%

dim mailcontent

dim boardqna
dim boarditem
dim id, mode
dim username, title, contents, regdate
dim replyuser, replytitle, replycontents, replydate
dim email, emailok, extsitename
dim userphone, replyqadiv
dim orderserial, sitename, outmallorderserial, outmallitemid

dim sql
dim itemid
dim targetMakerID
dim chargeid
dim delupcheans

dim referer
referer = Request.ServerVariables("HTTP_REFERER")

id = request("id")
mode = request("mode")

username = request("username")
title = request("title")
contents = db2html(request("contents"))
regdate = request("regdate")

replyuser = request("replyuser")
replytitle = ReplaceBracket(request("replytitle"))
replycontents = ReplaceBracket(request("replycontents"))
replydate = request("replydate")

userphone = request("userphone")
replyqadiv = request("replyqadiv")

Dim qaDiv	: qaDiv	= req("qaDiv","")	' 유형

'// 공백문자, 쉼표 제거(2014-09-04, skyer9)
email = Replace(Replace(request("email"), ",", ""), " ", "")
'emailok = "Y"
extsitename = request("extsitename")

itemid = request("itemid")
orderserial = request("orderserial")
sitename = request("sitename")
targetMakerID = request("targetMakerID")
chargeid = request("chargeid")
delupcheans = request("delupcheans")

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

                On Error Resume Next
                call SendMail("customer@10x10.co.kr", email, "즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]", mailcontent)
                On Error Goto 0

            response.write "<script>alert('답변메일이 발송되었습니다.')</script>"
        'end if

        response.write "<script>location.replace('" + referer + "')</script>"

elseif (mode = "firstreply") then

	    set boardqna = New CMyQNA
        set boarditem = new CMyQNAItem

		boardqna.read id
		if (boardqna.results(0).replyuser<>"") then
			response.write "<script>alert('이미 답변이 된 내용입니다.');</script>"
			response.write "<script>location.replace('" + referer + "')</script>"
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

        response.write "<script>location.replace('" + referer + "')</script>"

elseif  (mode = "del") then
                sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
                sql = sql + " set dispyn = 'N', replyuser = '" & session("ssBctID") & "' " + VbCRlf
                sql = sql + " where id = '" + Cstr(id) + "' and replycontents is NULL"
                ''response.write sql
                ''dbget.close()	:	response.End
                rsget.Open sql, dbget, 1
        response.write "<script>location.replace('" + referer + "')</script>"
end if

'' 답변,유형수정 2009년 4월 리뉴얼 버전
IF mode="REP" Or mode = "CHG" Then

	if (replyqadiv = "") then
		replyqadiv = "01"
	end if
	set boardqna = New CMyQNA

	boardqna.read ""	'' 초기값으로 세팅

	boardqna.results(0).id = id
	boardqna.results(0).qaDiv = qaDiv
	boardqna.results(0).replyuser = replyuser
	boardqna.results(0).replytitle = replytitle
	boardqna.results(0).replycontents = replycontents
	boardqna.results(0).Freplyqadiv = replyqadiv

    boardqna.BackProcData(mode)

	if (mode = "REP") then
        sql = " update s "
        sql = sql + " set s.RPLY_CNTS = convert(varchar(4000), q.replycontents), s.TenStatus = 'S' "
        sql = sql + " from "
        sql = sql + " 	[db_cs].dbo.tbl_MyQna q "
        sql = sql + " 	join [db_temp].[dbo].[tbl_Sabannet_Detail] s on q.extQnaIdx = s.idx "
        sql = sql + " where "
        sql = sql + " 	1 = 1 "
        sql = sql + " 	and q.id = " & id
        sql = sql + " 	and q.replydate is not NULL "
        'response.write sql
        'dbget.close()	:	response.End
        'rsget.Open sql, dbget, 1
	end if

    set boardqna = nothing

End If

if mode = "CGHITEMID" then
	if Len(itemid) > 8 or Not IsNumeric(itemid) then
		'// 제휴몰 상품코드 -> 상품코드
		outmallitemid = itemid

		Call GetItemIdFromOutmallItemID(sitename, outmallitemid, itemid)

		if (itemid = 0) then
			response.write "잘못된 상품코드입니다[0]." & outmallitemid
			dbget.close()	:	response.End
		end if

		if (itemid = -1) then
			response.write "작업 이전입니다.[0]." & sitename
			dbget.close()	:	response.End
		end if
	end if

    sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
    sql = sql + " set itemid = " + CStr(itemid) + " " + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"
    'response.write sql
    'dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
end if

if mode = "CGHORDSERIAL" then
	if (Len(orderserial) <> 11) then
		'// 제휴몰 주문번호 -> 주문번호
		outmallorderserial = orderserial
		Call GetOrderserialWithOutmallOrderserial(outmallorderserial, orderserial)
		if (orderserial = "") then
			response.write "잘못된 주문번호입니다[0]." & outmallorderserial
			dbget.close()	:	response.End
		end if
	end if

	dim ojumun
	set ojumun = new COrderMaster
	ojumun.FPageSize = 1
	ojumun.FCurrPage = 1
	ojumun.FRectOrderSerial = orderserial
	ojumun.QuickSearchOrderList

	'' 과거 6개월 이전 내역 검색
	if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
		ojumun.FRectOldOrder = "on"
		ojumun.QuickSearchOrderList

		if (ojumun.FResultCount<1) then
			response.write "잘못된 주문번호입니다[1]." & orderserial
			dbget.close()	:	response.End
		end if
	end if

    if (delupcheans = "Y") then
        sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
        sql = sql + " set orderserial = '" + CStr(orderserial) + "', itemid = NULL, orderdetailidx = 0, upchereplycontents = NULL, upchereplydate = NULL, upchereplyuser = NULL, upcheviewdate = NULL, makerid = NULL " + VbCRlf
        sql = sql + " where id = '" + Cstr(id) + "'"
    else
        sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
        sql = sql + " set orderserial = '" + CStr(orderserial) + "' " + VbCRlf
        sql = sql + " where id = '" + Cstr(id) + "'"
    end if
    ''response.write sql
    ''dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
end if

if mode="setmakerid" then
    sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
    sql = sql + " set makerid = '" + CStr(targetMakerID) + "' " + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"
    'response.write sql
    'dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
end if

if mode="setchargeid" then
    sql = "update [db_cs].[dbo].tbl_myqna " + VbCRlf
    sql = sql + " set chargeid = '" + CStr(chargeid) + "' " + VbCRlf
    sql = sql + " where id = '" + Cstr(id) + "'"
    'response.write sql
    'dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
end if

IF mode="CHG" or mode = "CGHITEMID" or mode = "CGHORDSERIAL" or mode = "setmakerid" or mode = "setchargeid" Then
	response.write "<script>alert('수정되었습니다')</script>"
    response.write "<script>location.replace('" + referer + "')</script>"
End If

'// 답변등록 알림 SMS 전송, MY알림 등록
IF mode="REP" Then
	if (userphone <> "") and Left(userphone, 2) = "01" then
		''Call SendNormalSMS(userphone, "", "[텐바이텐] 1:1 상담게시판에 답변이 등록되었습니다.")
		'Call SendNormalSMS_LINK(userphone, "", "[텐바이텐] 1:1 상담게시판에 답변이 등록되었습니다.")
		dim fullText, failText, btnJson
		'fullText = "1:1 상담게시판에 답변이 등록되었습니다." & vbCrLf & vbCrLf &_
		'		"요청하신 문의에 대한 답변이 등록되었습니다." & vbCrLf & vbCrLf &_
		'		"▶주문번호 : " & chkIIF(orderserial="","해당없음",CStr(orderserial)) & vbCrLf &_
		'		"▶답변일자 : " & formatdate(now,"0000.00.00-00:00") & vbCrLf & vbCrLf &_
		'		"감사합니다."
        fullText = "[10x10] 1:1 문의 답변등록안내" & vbCrLf & vbCrLf
        fullText = fullText & "문의하신 내용에 대한 답변이 등록되었습니다." & vbCrLf
        fullText = fullText & "확인 후 더 궁금하신 사항이 있으시면 언제든 문의 주세요." & vbCrLf
        fullText = fullText & "감사합니다. :)"
        failText = "[텐바이텐]1:1 상담게시판에 답변이 등록되었습니다."
		btnJson = "{""button"":[{""name"":""내 문의답변 바로가기"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/q27o0K8Mjjb""}]}"
		'Call SendKakaoMsg_LINK(userphone,"1644-6030","C-0003",fullText,"SMS","",failText,btnJson)
        Call SendKakaoCSMsg_LINK("", userphone,"1644-6030","KC-0016",fullText,"SMS","",failText,btnJson,"","")
		response.write "<script>alert('답변등록 알림 문자가 전송되었습니다.')</script>"
	end if

	'// MY알림 등록
	set boardqna = New CMyQNA
	boardqna.read(id)
	if (boardqna.results(0).userid <> "") then
		dim myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL

		myalarmtitle = boardqna.code2name(boardqna.results(0).qadiv)
		if (myalarmtitle <> "") then
			myalarmtitle = "<1:1 상담/" & myalarmtitle & ">"
		else
			myalarmtitle = "<1:1 상담>"
		end if

		myalarmsubtitle = db2html(boardqna.results(0).title)
		if Len(myalarmsubtitle) > 20 then
			myalarmsubtitle = Left(boardqna.results(0).title, 20) & " ..."
		end if

		myalarmcontents = "고객님 문의에 대해 답변드립니다."
		myalarmwwwTargetURL = "/my10x10/qna/myqnalist.asp"

		Call MyAlarm_InsertMyAlarm_SCM(boardqna.results(0).userid, "005", myalarmtitle, myalarmsubtitle, myalarmcontents, myalarmwwwTargetURL)
	end if
	set boardqna = nothing

End if


'' 답변메일발송 2009년 4월 리뉴얼 버전
IF mode="REP" Then

	dim oMail
	dim MailHTML
	dim MailTypeNo

	set oMail = New MailCls

	oMail.MailType = 15 '메일 종류별 고정값 (mailLib2.asp 참고)
	oMail.MailTitles = "[텐바이텐] " & username & "님께서 문의하신 내용에 대한 답변입니다."  '"즐거움이 가득한 쇼핑몰, 텐바이텐 [10X10=tenbyten]"
	'oMail.SenderMail = "customer@10x10.co.kr"
	'oMail.SenderNm = "텐바이텐"

	oMail.AddrType = "string"
	oMail.ReceiverNm = username
	oMail.ReceiverMail = email

	MailHTML = oMail.getMailTemplate()

	IF MailHTML="" Then
		response.write "<script>alert('메일발송이 실패 하였습니다.')</script>"
    	response.write "<script>location.replace('" + referer + "')</script>"
		dbget.close()	:	response.End
	End IF

	MailHTML =replace(MailHTML,"[$USER_NAME$]",oMail.ReceiverNm)
	MailHTML =replace(MailHTML,"[$QUESTION_TIME$]",regdate)
	MailHTML =replace(MailHTML,"[$QUESTION_TITLE$]",server.HTMLEncode(title))
	MailHTML =replace(MailHTML,"[$QUESTION_CONTENTS$]", nl2br(server.HTMLEncode(db2html(contents))))
	If replydate = "" Then
		replydate = now()
	End If
	MailHTML =replace(MailHTML,"[$ANSWER_TIME$]",replydate)
	MailHTML =replace(MailHTML,"[$ANSWER_TITLE$]",server.HTMLEncode(replytitle))
	MailHTML =replace(MailHTML,"[$ANSWER_CONTENTS$]", nl2br(server.HTMLEncode(db2html(replycontents))))
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
    response.write "<script>location.replace('" + referer + "')</script>"

End if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
