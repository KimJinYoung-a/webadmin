<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Dim sellsite, itemid, outmallGoodNo
Dim strSql, gubun, paramData, retVal
gubun			= request("gubun")
sellsite		= request("sellsite")
itemid			= Trim(request("itemid"))
outmallGoodNo	= Trim(request("outmallGoodNo"))

If goodNoUpdateUser <> "Y" Then
	response.write "<script language='javascript'>alert('ºˆ¡§«“ ºˆ ¿÷¥¬ ±««—¿Ã æ¯Ω¿¥œ¥Ÿ.\n\n±Ë¡¯øµ πÆ¿«');window.close();</script>"
	response.end
End If

If gubun = "I" Then
	If sellsite = "interpark" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_interpark_reg_item where itemid="&itemid&")"
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		INSERT INTO [db_item].[dbo].tbl_interpark_reg_item (itemid, regdate, reguserid, interparkPrdno, mayiparkPrice)"
		strSql = strSql & " 		SELECT itemid, GETDATE(), 'system', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 		FROM db_item.dbo.tbl_item "
		strSql = strSql & " 		WHERE itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE [db_item].[dbo].tbl_interpark_reg_item "
		strSql = strSql & " 		SET interparkPrdno = '"& outmallGoodNo &"' "
		strSql = strSql & " 		WHERE itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=interpark&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/InterparkProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=interpark&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/InterparkProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "11st1010" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_11st_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_11st_regitem (itemid, regdate, reguserid, st11statCD, st11GoodNo)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"' "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE db_etcmall.dbo.tbl_11st_regitem "
		strSql = strSql & " 		SET st11GoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		,st11statCD = 7 "
		strSql = strSql & " 		WHERE itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=11st1010&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/11stProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=11st1010&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/11stProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=11st1010&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/11stProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "cjmall" Then
		strSql = ""
		strSql = strSql & " UPDATE db_item.dbo.tbl_cjmall_regitem "
		strSql = strSql & " SET cjmallPrdno = '"& outmallGoodNo &"' "
		strSql = strSql & " ,cjmallstatCD = 7 "
		strSql = strSql & " WHERE itemid = '"& itemid &"' "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=cjmall&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/cjmallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "coupang" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_coupang_regitem (itemid, regdate, reguserid, coupangstatCD, coupangGoodNo, coupangPrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.CoupangPrice = i.sellcash "
		strSql = strSql & " 		, R.CoupangGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.CoupangStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_coupang_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=coupang&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/coupangProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=coupang&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/coupangProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=coupang&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/coupangProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "ssg" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_ssg_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_ssg_regitem (itemid, regdate, reguserid, ssgstatCD, ssgGoodNo, ssgPrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.ssgPrice = i.sellcash "
		strSql = strSql & " 		, R.ssgGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.ssgStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_ssg_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=ssg&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=ssg&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/ssgProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "lotteon" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_lotteon_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_lotteon_regitem (itemid, regdate, reguserid, lotteonstatCD, lotteonGoodNo, lotteonPrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.lotteonPrice = i.sellcash "
		strSql = strSql & " 		, R.lotteonGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.lotteonStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_lotteon_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteon&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lotteonProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteon&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lotteonProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteon&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lotteonProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "WMP" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_wemake_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_wemake_regitem (itemid, regdate, reguserid, wemakestatCD, wemakeGoodNo, wemakePrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.wemakePrice = i.sellcash "
		strSql = strSql & " 		, R.wemakeGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.wemakeStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_wemake_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=WMP&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=WMP&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=WMP&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/wmpProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "gmarket1010" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_gmarket_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_gmarket_regitem (itemid, regdate, reguserid, gmarketstatCD, gmarketGoodNo, gmarketPrice, gmarketsellyn, APIadditem, APIaddgosi, APIaddopt)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash, 'Y', 'Y', 'Y', 'Y' "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.gmarketPrice = i.sellcash "
		strSql = strSql & " 		, R.gmarketGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.gmarketStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_gmarket_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=gmarket1010&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/GmarketProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=gmarket1010&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/GmarketProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "lotteimall" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_item.dbo.tbl_ltimall_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_item.dbo.tbl_ltimall_regitem (itemid, regdate, reguserid, ltimallstatCD, ltimallGoodNo, ltimallPrice, ltimallsellyn)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash, 'Y' "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.ltimallPrice = i.sellcash "
		strSql = strSql & " 		, R.ltimallGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.ltimallStatCd = 7 "
		strSql = strSql & " 		FROM db_item.dbo.tbl_ltimall_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteimall&action=DISPVIEW"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/LtimallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteimall&action=CHKSTOCK"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/LtimallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteimall&action=PRICE"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/LtimallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lotteimall&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/LtimallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "LFmall" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_lfmall_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_lfmall_regitem (itemid, regdate, reguserid, lfmallstatCD, lfmallGoodNo, lfmallPrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.lfmallPrice = i.sellcash "
		strSql = strSql & " 		, R.lfmallGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.lfmallStatCd = 7 "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_lfmall_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lfmall&action=CHKSTAT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lfmallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=lfmall&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/lfmallProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	ElseIf sellsite = "auction1010" Then
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_auction_regitem where itemid="&itemid&")"
		strSql = strSql & " BEGIN"& VbCRLF
		strSql = strSql & " 	INSERT INTO db_etcmall.dbo.tbl_auction_regitem (itemid, regdate, reguserid, auctionstatCD, auctionGoodNo, auctionPrice)"
		strSql = strSql & " 	SELECT itemid, GETDATE(), 'system', '7', '"& outmallGoodNo &"', sellcash "
		strSql = strSql & " 	FROM db_item.dbo.tbl_item "
		strSql = strSql & " 	WHERE itemid = '"& itemid &"' "
		strSql = strSql & " END "
		strSql = strSql & " ELSE " & VbCrlf
		strSql = strSql & "		BEGIN"& VbCRLF
		strSql = strSql & " 		UPDATE R "
		strSql = strSql & " 		SET R.auctionPrice = i.sellcash "
		strSql = strSql & " 		, R.auctionGoodNo = '"& outmallGoodNo &"' "
		strSql = strSql & " 		, R.auctionStatCd = 7 "
		strSql = strSql & " 		, R.APIadditem = 'Y' "
		strSql = strSql & " 		, R.APIaddopt = 'Y' "
		strSql = strSql & " 		, R.APIaddgosi = 'Y' "
		strSql = strSql & " 		FROM db_etcmall.dbo.tbl_auction_regitem R "
		strSql = strSql & " 		JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid "
		strSql = strSql & " 		WHERE R.itemid = '"& itemid &"' "
		strSql = strSql & " 	END "
		dbget.Execute strSql

		paramData = "redSsnKey=system&itemid="&itemid&"&mallid=auction1010&action=EDIT"
		retVal = SendReq("http://wapi.10x10.co.kr/outmall/proc/auctionProc.asp",paramData)
		rw retVal
		response.flush
		response.clear
	End If
	
	If session("ssBctID")="kjy8517" Then
		rw strSql
	End If

	response.write "<input type='button' value='√≥¿Ω¿∏∑Œ' onclick=location.replace('/admin/etc/outmall/popGoodNoMatch.asp');> "
	response.end
End If
%>
<script language='javascript'>
function frmsubmit(){
	var frm = document.frm;
	if(frm.sellsite.value == ''){
		alert('¡¶»ﬁ∏Ù¿ª º±≈√«œººø‰');
		frm.sellsite.focus(); 
		return;
	}
	if(frm.itemid.value == ''){
		alert('≈ŸπŸ¿Ã≈Ÿ ªÛ«∞ƒ⁄µÂ∏¶ ¿‘∑¬«œººø‰');
		frm.itemid.focus();
		return;
	}
	if(frm.outmallGoodNo.value == ''){
		alert('¡¶»ﬁ∏Ù ªÛ«∞ƒ⁄µÂ∏¶ ¿‘∑¬«œººø‰');
		frm.outmallGoodNo.focus();
		return;
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="gubun" value="I">
<tr bgcolor="#FFFFFF" height="30">
	<td>
		¡¶»ﬁ∏Ù:
		<select class="select" name="sellsite">
			<option></option>
			<option value="interpark">¿Œ≈Õ∆ƒ≈©</option>
			<option value="11st1010">11π¯∞°</option>
			<option value="cjmall">CJ∏Ù</option>
			<option value="coupang">ƒÌ∆Œ</option>
			<option value="ssg">SSG</option>
			<option value="lotteon">∑‘µ•On</option>
			<option value="WMP">WMP</option>
			<option value="lotteimall">∑‘µ•æ∆¿Ã∏Ù</option>
			<option value="LFmall">LFmall</option>
			<option value="auction1010">ø¡º«</option>
<!--
			<option value="gmarket1010">¡ˆ∏∂ƒœ(NEW)</option>
			<option value="ezwel">¿Ã¡ˆ¿£∆‰æÓ</option>
			<option value="hmall1010">Hmall</option>
-->
		</select><br />
		≈ŸπŸ¿Ã≈Ÿ ªÛ«∞ƒ⁄µÂ : <input type="text" class="text" name="itemid" ><br />
		¡¶»ﬁ∏Ù ªÛ«∞ƒ⁄µÂ : <input type="text" class="text" name="outmallGoodNo">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center">
<td>
	<input type="button" class="button" value="¿˙¿Â" onclick="frmsubmit();">&nbsp;&nbsp;
</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->