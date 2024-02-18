<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ø¿«¡òﬁ¿ÃøÎπÆ¿«
' Hieditor : 2009.04.07 º≠µøºÆ ª˝º∫
'			 2011.05.03 «—øÎπŒ ºˆ¡§
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/offshopqnacls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<%
dim mailcontent ,boardqna ,boarditem ,idx, mode, replytitle, replycontents, replyuser, param, isNew
dim email, emailok, extsitename, usercell, cellok ,shopid, page, menupos ,SearchKey, SearchString
Dim brandname, regdate, shopname, itemname, itemid, replydate, contents, title, shopphone
dim sql, fileName, query1, fs, dirPath, objFile, mailheader, mailfooter, offshopid
	idx = request("idx")
	mode = request("mode")
	replytitle = request("replytitle")
	replycontents = request("replycontents")
	email = request("email")
	emailok = request("emailok")
	extsitename = request("extsitename")
	usercell = request("usercell")
	cellok	= request("cellok")
	menupos = Request("menupos")
	page = Request("page")
	shopid = Request("shopid")
	isNew = Request("isNew")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	regdate = Request("regdate")
	brandname = Request("brandname")
	itemname = Request("itemname")
	itemid = Request("itemid")
	contents = Request("contents")
	title = Request("title")
	replydate = FormatDate(now(),"0000-00-00")
	offshopid = Request("offshopid")
param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&shopid=" & shopid & "&isNew=" & isNew & "&menupos=" & menupos

if (mode = "reply") then
    set boardqna = New CMyQNA
    set boarditem = new CMyQNAItem

    boarditem.idx = idx
    boarditem.replyuser = "10x10"
    boarditem.replytitle = html2db(replytitle)
    boarditem.replycontents = html2db(replycontents)

    boardqna.reply(boarditem)

    if (emailok = "Y") then

		If offshopid<>"" Then
			query1 = " select top 1 shopname, shopphone from [db_shop].[dbo].tbl_shop_user"
			query1 = query1 & " where isusing='Y' "
			query1 = query1 & " and userid='" +  offshopid + "'"
			rsget.Open query1,dbget,1
		   if not rsget.EOF  then
				shopname = rsget("shopname")
				shopphone = rsget("shopphone")
		   end if
		   rsget.close
		End If

		mailcontent = "<html>" + vbcrlf
		mailcontent = mailcontent + "<head>" + vbcrlf
		mailcontent = mailcontent + "<title>∏≈¿Â πÆ¿« ¥‰∫Ø ∏ﬁ¿œ</title>" + vbcrlf
		mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>" + vbcrlf
		mailcontent = mailcontent + "<meta name='viewport' content='width=device-width, initial-scale=1.0'>" + vbcrlf
		mailcontent = mailcontent + "</head>" + vbcrlf
		mailcontent = mailcontent + "<body style='margin:0; padding:0;'>" + vbcrlf
		mailcontent = mailcontent + "<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto; background-color:#f7f7f7' background='#f7f7f7'>" + vbcrlf
		mailcontent = mailcontent + "	<tr>" + vbcrlf
		mailcontent = mailcontent + "		<td style='background-color:#f7f7f7; text-align:center;'>" + vbcrlf
		mailcontent = mailcontent + "			<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:750px; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "				<thead>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td><img src='http://mailzine.10x10.co.kr/2017/header.png' style='vertical-align:top; width:100%;' /></td>" + vbcrlf
		mailcontent = mailcontent + "					</tr>" + vbcrlf
		mailcontent = mailcontent + "				</thead>" + vbcrlf
		mailcontent = mailcontent + "				<tbody>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td style='margin-left:auto; margin-right:auto; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "							<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:58px 0 0 0; object-fit:contain; text-align:center;'><img src='http://mailzine.10x10.co.kr/2017/ico_qna.png' alt='QNA' /></td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:30px 0 0 0; font-size:47px; line-height:54px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; letter-spacing:-4px; color:#0d0d0d; text-align:center;'><strong style='color:#ff3131;'><span style='color:#ff3131;'>∏≈¿Âø° ∞¸«— πÆ¿«ªÁ«◊</span></strong>¿ª<br />æ»≥ªµÂ∏≥¥œ¥Ÿ.</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:20px 0 80px 0; font-size:23px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; letter-spacing:-1.3px; color:#ff001c; text-align:center;'>" + shopname + "</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "							</table>" + vbcrlf
		mailcontent = mailcontent + "						</td>" + vbcrlf
		mailcontent = mailcontent + "					</tr>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td style='margin-left:auto; margin-right:auto; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "							<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:0 29px 88px; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "										<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<th style='width:100%; padding:0 0 15px 3px; color:#ff001c; font-size:32px; line-height:36px; font-family:verdana, sans-serif, dotum, ""µ∏øÚ"", sans-serif; text-align:left;'>Q.<span style='margin:0; padding:0 0 0 7px; font-size:25px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; color:#0d0d0d; letter-spacing:-1px;'>πÆ¿««œΩ≈ ≥ªøÎ</span></th>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<td style='border-top:solid 2px #000;'>" + vbcrlf
		mailcontent = mailcontent + "													<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:20px 46px 20px 44px; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-size:21px; line-height:28px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left;'>" + title + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'><a href='http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" +CStr(itemid) + "' target='_blank' style='text-decoration:none; color:#0d0d0d;'>[" + CStr(brandname) + "]<br />" + CStr(itemname) + "</a></td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'>" + contents + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr><td style='width:100%; margin:0; padding:30px 0 40px 44px; border-bottom:solid 1px #eaeaea; color:#999999; font-size:20px; line-height:27px; font-weight:300; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; letter-spacing:-1px;'>" + regdate + "</td></tr>" + vbcrlf
		mailcontent = mailcontent + "													</table>" + vbcrlf
		mailcontent = mailcontent + "												</td>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "										</table>" + vbcrlf
		mailcontent = mailcontent + "									</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:0 29px; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "										<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<th style='width:100%; padding:0 0 15px 3px; color:#ff001c; font-size:32px; line-height:36px; font-family:verdana, sans-serif, dotum, ""µ∏øÚ"", sans-serif; text-align:left;'>A.<span style='margin:0; padding:0 0 0 7px; font-size:25px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; color:#0d0d0d; letter-spacing:-1px;'>πÆ¿«≥ªøÎø° ¥Î«— ¥‰∫Ø</span></th>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<td style='border-top:solid 2px #000;'>" + vbcrlf
		mailcontent = mailcontent + "													<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:20px 46px 20px 44px; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-size:21px; line-height:28px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left;'>" + html2db(replytitle) + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'>" + nl2br(db2html(replycontents)) +"<br /><br />≈ŸπŸ¿Ã≈Ÿ " + shopname + " <br />" + Cstr(shopphone) + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr><td style='width:100%; margin:0; padding:30px 0 40px 44px; border-bottom:solid 1px #eaeaea; color:#999999; font-size:20px; line-height:27px; font-weight:300; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; letter-spacing:-1px;'>" + replydate + "</td></tr>" + vbcrlf
		mailcontent = mailcontent + "													</table>" + vbcrlf
		mailcontent = mailcontent + "												</td>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "										</table>" + vbcrlf
		mailcontent = mailcontent + "									</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='padding:80px 0 72px 0; background-color:#fff; text-align:center;'><a href='http://www.10x10.co.kr' target='_blank' ><img src='http://mailzine.10x10.co.kr/2017/btn_go_tenten_red.png' alt='¥ı ∏π¿∫ ªÛ«∞ ∫∏∑Ø∞°±‚' style='border:0;' /></a></td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='height:90px; border-top:1px solid #f4f4f4; background-color:#fff; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; font-size:18px; line-height:1.39; letter-spacing:-1px; text-align:center; color:#808080;'>«◊ªÛ ∞Ì∞¥¥‘¿« ∞≥¿Œ ¡§∫∏∏¶ º“¡ﬂ«œ∞‘ ∫∏»£«œ∏Á,<br />≥°±Ó¡ˆ ±‚∫– ¡¡¿∫ ºÓ«Œ¿Ã µ… ºˆ ¿÷µµ∑œ √÷º±¿ª ¥Ÿ«œ∞⁄Ω¿¥œ¥Ÿ.</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "							</table>" + vbcrlf

        ' ∆ƒ¿œ¿ª ∫“∑ØøÕº≠ ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' «™≈Õ

		mailcontent=mailcontent&mailfooter

        call sendmail("customer@10x10.co.kr", email, "¡Ò∞≈øÚ¿Ã ∞°µÊ«— ºÓ«Œ∏Ù, ≈ŸπŸ¿Ã≈Ÿ [10X10=tenbyten]", mailcontent)

    	response.write "<script>alert('¥‰∫Ø∏ﬁ¿œ¿Ã πﬂº€µ«æ˙Ω¿¥œ¥Ÿ.')</script>"
    end if

	If cellok = "Y" Then
		Call SendNormalSMS(usercell,"","[≈ŸπŸ¿Ã≈Ÿ]Shop Q&A ø° ≥≤±‚Ω≈ ±€ø° ¥‰∫Ø¿Ã ¥ﬁ∑»Ω¿¥œ¥Ÿ.")
	End If

    response.write "<script>location.replace('offshop_qna_board_reply.asp?idx=" + idx + "&page=" & page & param & "')</script>"

elseif (mode = "firstreply") then

    set boardqna = New CMyQNA
    set boarditem = new CMyQNAItem

	boardqna.frectidx = idx
	boardqna.read
	if (boardqna.FItemList(0).replyuser<>"") then
		response.write "<script>alert('¿ÃπÃ ¥‰∫Ø¿Ã µ» ≥ªøÎ¿‘¥œ¥Ÿ.');</script>"
		response.write "<script>location.replace('offshop_qna_board_reply.asp?idx=" + idx + "&page=" & page & param & "')</script>"
		dbget.close()	:	response.End
	end if

    boarditem.idx = idx
    boarditem.replyuser = "10x10"
    boarditem.replytitle = html2db(replytitle)
    boarditem.replycontents = html2db(replycontents)

    boardqna.reply(boarditem)

    if (emailok = "Y") then
		
		If offshopid<>"" Then
			query1 = " select top 1 shopname, shopphone from [db_shop].[dbo].tbl_shop_user"
			query1 = query1 & " where isusing='Y' "
			query1 = query1 & " and userid='" +  offshopid + "'"
			rsget.Open query1,dbget,1
		   if not rsget.EOF  then
				shopname = rsget("shopname")
				shopphone = rsget("shopphone")
		   end if
		   rsget.close
		End If

		mailcontent = "<html>" + vbcrlf
		mailcontent = mailcontent + "<head>" + vbcrlf
		mailcontent = mailcontent + "<title>∏≈¿Â πÆ¿« ¥‰∫Ø ∏ﬁ¿œ</title>" + vbcrlf
		mailcontent = mailcontent + "<meta http-equiv='Content-Type' content='text/html; charset=euc-kr'>" + vbcrlf
		mailcontent = mailcontent + "<meta name='viewport' content='width=device-width, initial-scale=1.0'>" + vbcrlf
		mailcontent = mailcontent + "</head>" + vbcrlf
		mailcontent = mailcontent + "<body style='margin:0; padding:0;'>" + vbcrlf
		mailcontent = mailcontent + "<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto; background-color:#f7f7f7' background='#f7f7f7'>" + vbcrlf
		mailcontent = mailcontent + "	<tr>" + vbcrlf
		mailcontent = mailcontent + "		<td style='background-color:#f7f7f7; text-align:center;'>" + vbcrlf
		mailcontent = mailcontent + "			<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:750px; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "				<thead>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td><img src='http://mailzine.10x10.co.kr/2017/header.png' style='vertical-align:top; width:100%;' /></td>" + vbcrlf
		mailcontent = mailcontent + "					</tr>" + vbcrlf
		mailcontent = mailcontent + "				</thead>" + vbcrlf
		mailcontent = mailcontent + "				<tbody>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td style='margin-left:auto; margin-right:auto; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "							<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:58px 0 0 0; object-fit:contain; text-align:center;'><img src='http://mailzine.10x10.co.kr/2017/ico_qna.png' alt='QNA' /></td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:30px 0 0 0; font-size:47px; line-height:54px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; letter-spacing:-4px; color:#0d0d0d; text-align:center;'><strong style='color:#ff3131;'><span style='color:#ff3131;'>∏≈¿Âø° ∞¸«— πÆ¿«ªÁ«◊</span></strong>¿ª<br />æ»≥ªµÂ∏≥¥œ¥Ÿ.</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:20px 0 80px 0; font-size:23px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; letter-spacing:-1.3px; color:#ff001c; text-align:center;'>" + shopname + "</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "							</table>" + vbcrlf
		mailcontent = mailcontent + "						</td>" + vbcrlf
		mailcontent = mailcontent + "					</tr>" + vbcrlf
		mailcontent = mailcontent + "					<tr>" + vbcrlf
		mailcontent = mailcontent + "						<td style='margin-left:auto; margin-right:auto; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "							<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:0 29px 88px; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "										<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<th style='width:100%; padding:0 0 15px 3px; color:#ff001c; font-size:32px; line-height:36px; font-family:verdana, sans-serif, dotum, ""µ∏øÚ"", sans-serif; text-align:left;'>Q.<span style='margin:0; padding:0 0 0 7px; font-size:25px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; color:#0d0d0d; letter-spacing:-1px;'>πÆ¿««œΩ≈ ≥ªøÎ</span></th>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<td style='border-top:solid 2px #000;'>" + vbcrlf
		mailcontent = mailcontent + "													<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:20px 46px 20px 44px; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-size:21px; line-height:28px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left;'>" + title + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'><a href='http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" +CStr(itemid) + "' target='_blank' style='text-decoration:none; color:#0d0d0d;'>[" + CStr(brandname) + "]<br />" + CStr(itemname) + "</a></td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'>" + contents + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr><td style='width:100%; margin:0; padding:30px 0 40px 44px; border-bottom:solid 1px #eaeaea; color:#999999; font-size:20px; line-height:27px; font-weight:300; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; letter-spacing:-1px;'>" + regdate + "</td></tr>" + vbcrlf
		mailcontent = mailcontent + "													</table>" + vbcrlf
		mailcontent = mailcontent + "												</td>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "										</table>" + vbcrlf
		mailcontent = mailcontent + "									</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='margin:0; padding:0 29px; background-color:#fff;'>" + vbcrlf
		mailcontent = mailcontent + "										<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<th style='width:100%; padding:0 0 15px 3px; color:#ff001c; font-size:32px; line-height:36px; font-family:verdana, sans-serif, dotum, ""µ∏øÚ"", sans-serif; text-align:left;'>A.<span style='margin:0; padding:0 0 0 7px; font-size:25px; line-height:36px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; color:#0d0d0d; letter-spacing:-1px;'>πÆ¿«≥ªøÎø° ¥Î«— ¥‰∫Ø</span></th>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "											<tr>" + vbcrlf
		mailcontent = mailcontent + "												<td style='border-top:solid 2px #000;'>" + vbcrlf
		mailcontent = mailcontent + "													<table align='center' border='0' cellpadding='0' cellspacing='0' style='width:100%; margin-left:auto; margin-right:auto;'>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:20px 46px 20px 44px; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-size:21px; line-height:28px; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left;'>" + html2db(replytitle) + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr>" + vbcrlf
		mailcontent = mailcontent + "															<td style='width:100%; margin:0; padding:19px 46px 0 44px; font-size:20px; line-height:30px;font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; color:#0d0d0d; text-align:left; letter-spacing:-1px;'>" + nl2br(db2html(replycontents)) +"<br /><br />≈ŸπŸ¿Ã≈Ÿ " + shopname + " <br />" + Cstr(shopphone) + "</td>" + vbcrlf
		mailcontent = mailcontent + "														</tr>" + vbcrlf
		mailcontent = mailcontent + "														<tr><td style='width:100%; margin:0; padding:30px 0 40px 44px; border-bottom:solid 1px #eaeaea; color:#999999; font-size:20px; line-height:27px; font-weight:300; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; text-align:left; letter-spacing:-1px;'>" + replydate + "</td></tr>" + vbcrlf
		mailcontent = mailcontent + "													</table>" + vbcrlf
		mailcontent = mailcontent + "												</td>" + vbcrlf
		mailcontent = mailcontent + "											</tr>" + vbcrlf
		mailcontent = mailcontent + "										</table>" + vbcrlf
		mailcontent = mailcontent + "									</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='padding:80px 0 72px 0; background-color:#fff; text-align:center;'><a href='http://www.10x10.co.kr' target='_blank' ><img src='http://mailzine.10x10.co.kr/2017/btn_go_tenten_red.png' alt='¥ı ∏π¿∫ ªÛ«∞ ∫∏∑Ø∞°±‚' style='border:0;' /></a></td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "								<tr>" + vbcrlf
		mailcontent = mailcontent + "									<td style='height:90px; border-top:1px solid #f4f4f4; background-color:#fff; font-family:""∏º¿∫∞ÌµÒ"",""Malgun Gothic"",""µ∏øÚ"", dotum, sans-serif; font-size:18px; line-height:1.39; letter-spacing:-1px; text-align:center; color:#808080;'>«◊ªÛ ∞Ì∞¥¥‘¿« ∞≥¿Œ ¡§∫∏∏¶ º“¡ﬂ«œ∞‘ ∫∏»£«œ∏Á,<br />≥°±Ó¡ˆ ±‚∫– ¡¡¿∫ ºÓ«Œ¿Ã µ… ºˆ ¿÷µµ∑œ √÷º±¿ª ¥Ÿ«œ∞⁄Ω¿¥œ¥Ÿ.</td>" + vbcrlf
		mailcontent = mailcontent + "								</tr>" + vbcrlf
		mailcontent = mailcontent + "							</table>" + vbcrlf

        ' ∆ƒ¿œ¿ª ∫“∑ØøÕº≠ ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' «™≈Õ

		mailcontent=mailcontent&mailfooter

        call sendmail("customer@10x10.co.kr", email, "¡Ò∞≈øÚ¿Ã ∞°µÊ«— ºÓ«Œ∏Ù, ≈ŸπŸ¿Ã≈Ÿ [10X10=tenbyten]", mailcontent)
		response.write "<script>alert('¥‰∫Ø∏ﬁ¿œ¿Ã πﬂº€µ«æ˙Ω¿¥œ¥Ÿ.')</script>"
    end if

	If cellok = "Y" Then
		Call SendNormalSMS(usercell,"","[≈ŸπŸ¿Ã≈Ÿ]Shop Q&A ø° ≥≤±‚Ω≈ ±€ø° ¥‰∫Ø¿Ã ¥ﬁ∑»Ω¿¥œ¥Ÿ.")
	End If

    response.write "<script>location.replace('offshop_qna_board_reply.asp?idx=" + idx + "&page=" & page & param & "')</script>"

elseif  (mode = "del") then

    sql = "update [db_shop].[dbo].tbl_offshop_qna " + VbCRlf
    sql = sql + " set isusing = 'N'" + VbCRlf
    sql = sql + " where idx = '" + Cstr(idx) + "'"
    'response.write sql
    'dbget.close()	:	response.End
    rsget.Open sql, dbget, 1
	response.write "<script>location.replace('board/itemqna_list.asp?page=" & page & param & "')</script>"
end if
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->