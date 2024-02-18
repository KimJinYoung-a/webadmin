<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 이벤트
' History : 2010.03.09 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
Dim mode , evt_state, eCategory, chkdisp ,igiftcnt , istatus ,shopid ,isracknum
Dim evt_code, evt_kind, evt_name, evt_startdate, evt_enddate, evt_prizedate , img_basic
Dim isgift ,israck ,isprize, partMDid, evt_forward, brand ,evt_using , opendate , closedate
Dim evt_comment, strAdd , strSql, tmpevt_code , selCM , issale
Dim addshopid

	igiftcnt = 0
	mode 	= requestCheckVar(Request.Form("mode"),25) '데이터 처리종류
	evt_code  	= requestCheckVar(Request.Form("evt_code"),10)	'이벤트코드
	evt_using 		= requestCheckVar(Request.Form("evt_using"),1)
	evt_kind 	= requestCheckVar(Request.Form("evt_kind"),4)			
	evt_name 	= requestCheckVar(Request.Form("evt_name"),100)
	evt_startdate		= requestCheckVar(Request.Form("evt_startdate"),10)
	evt_enddate 		= requestCheckVar(Request.Form("evt_enddate"),10)
	evt_prizedate 		= requestCheckVar(Request.Form("evt_prizedate"),10)
	evt_state 		= requestCheckVar(Request.Form("evt_state"),4)
	chkdisp 	= requestCheckVar(Request.Form("chkdisp"),2)
	eCategory 	= requestCheckVar(Request.Form("selC"),10)
	selCM = requestCheckVar(Request.Form("selCM"),10)
	evt_comment 	= requestCheckVar(Request.Form("evt_comment"),200)
	issale 		= requestCheckVar(Request.Form("issale"),1)
	isgift 		= requestCheckVar(Request.Form("isgift"),1)
	israck 		= requestCheckVar(Request.Form("israck"),1)
	isprize 		= requestCheckVar(Request.Form("isprize"),1)
	partMDid 	= requestCheckVar(Request.Form("partMDid"),32)
	evt_forward = html2db(Trim(Request.Form("evt_forward")))	    
    brand 		= requestCheckVar(Request.Form("brand"),32)
 	opendate = requestCheckVar(Request.Form("opendate"),30)
 	closedate =requestCheckVar(Request.Form("closedate"),30)
 	shopid		= requestCheckVar(Request("shopid"),32)
 	isracknum = requestCheckVar(Request.Form("isracknum"),10)    
	img_basic = requestCheckVar(Request.Form("img_basic"),256)
    
    addshopid = Request.Form("addshopid")

	if issale <> "Y" then
		issale = "N"
	end if	
	if isgift <> "Y" then
		isgift = "N"
	end if
	if isprize <> "Y" then
		isprize = "N"
	end if
		
	if israck <> "Y" then
		israck = "N"
	end if
	
'이벤트 신규 / 수정
if mode = "event_edit" then

	'//신규등록
	if evt_code = "" then
		if evt_forward <> "" then
			if checkNotValidHTML(evt_forward) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			dbget.close()	:	response.End
			end if
		end if

		'상태가 오픈일때 오픈일 등록
		opendate = "null"
		closedate = "null"
		
		IF evt_state = 7 THEN
			opendate = "getdate()"
		ELSEIF evt_state = 9 THEN
			closedate = "getdate()"
		END IF
	
		'트랜잭션 (1.master등록/2.disply등록)
		dbget.beginTrans
			'--1.master등록
			strSql = "INSERT INTO [db_shop].[dbo].[tbl_event_off] (evt_kind, evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_state, opendate, closedate, evt_lastupdate, adminid,shopid) "&_
				"		VALUES ("&evt_kind&",'"&html2db(evt_name)&"','"&evt_startdate&"','"&evt_enddate&"','"&evt_prizedate&"','"&evt_state&"',"&opendate&","&closedate&",getdate(),'"&session("ssBctId")&"','"&shopid&"')"
			
			'response.write strSql &"<br>"
			dbget.execute strSql
	
			strSql = "select SCOPE_IDENTITY()"
			rsget.Open strSql, dbget, 0
			tmpevt_code = rsget(0)
			rsget.Close
	
			'--2.disply등록
			strSql = ""
			IF chkdisp = "on" THEN
				strSql = "INSERT INTO [db_shop].[dbo].[tbl_event_off_display] (evt_code,evt_category,evt_cateMid,issale,isgift ,israck ,isprize, partMDid, evt_forward, brand, evt_comment,img_basic,isracknum) "&_
					" VALUES ("&tmpevt_code&",'"&eCategory&"','"&selCM&"','"&issale&"','"&isgift&"','"&israck&"','"&isprize&"','"&partMDid&"','"&html2db(evt_forward)&"','"&brand&"','"&html2db(evt_comment)&"','"&html2db(img_basic)&"','"&isracknum&"')"
				dbget.execute strSql
			END IF
		
	'//수정모드
	else
		if evt_forward <> "" then
			if checkNotValidHTML(evt_forward) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			dbget.close()	:	response.End
			end if
		end if

		strAdd = ""
	 	IF (evt_state = 7 and opendate ="" ) THEN 	'오픈처리일 설정
			strAdd = ", opendate = getdate() "
		ELSEIF (evt_state = 9 and closedate ="" ) THEN
			strAdd = ", closedate = getdate() "	'종료처리일 설정
		END IF
	
		'종료일 이전에 종료시 종료일 현재 날짜로 변경
		IF evt_state = 9 and  datediff("d",evt_enddate,date()) <0 THEN
				evt_enddate = date()
		END IF

		'트랜잭션 (1.master수정/2.disply수정)
		dbget.beginTrans
	
		'--1.master수정
		strSql = " UPDATE [db_shop].[dbo].[tbl_event_off] "&_
				 " SET  [evt_kind]="&evt_kind&",[evt_name]='"&html2db(evt_name)&"', [evt_startdate]='"&evt_startdate&"'"&_
				 " , [evt_enddate]='"&evt_enddate&"',[evt_prizedate]='"&evt_prizedate&"', [evt_state]="&evt_state&""&_
				 " , [evt_using] = '"&evt_using&"', evt_lastupdate = getdate(), adminid = '"&session("ssBctId")&"' "&strAdd&_
				 " , shopid = '"&shopid&"'" &_
				 " WHERE evt_code = "&evt_code
		dbget.execute strSql
	
		'--2.disply수정
		strSql = "SELECT evt_code FROM [db_shop].[dbo].[tbl_event_off_display] WHERE evt_code= "&evt_code
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			IF chkdisp = "on" THEN
				strSql = " UPDATE [db_shop].[dbo].[tbl_event_off_display] SET "&_
						" evt_category ='"&eCategory&"',evt_cateMid ='"&selCM&"', issale = '"&issale&"',isgift='"&isgift&"',israck='"&israck&"'"&_
						" ,isprize='"&isprize&"', partMDid='"&partMDid&"', evt_forward='"&evt_forward&"', brand='"&brand&"'"&_						
						" ,isracknum = '"&isracknum&"',evt_comment = '"&evt_comment&"',img_basic = '"&html2db(img_basic)&"' " &_
						" WHERE evt_code ="&evt_code
				
				'response.write strSql
				dbget.execute strSql
			ELSE
				strSql = " DELETE FROM  [db_shop].[dbo].[tbl_event_off_display]  WHERE  evt_code ="&evt_code
				dbget.execute strSql
			END IF
		ELSE
			IF chkdisp = "on" THEN
				strSql = "INSERT INTO [db_shop].[dbo].[tbl_event_off_display] "&_
						" (evt_code, evt_category,evt_cateMid,issale,isgift ,israck ,isprize, partMDid, evt_forward, brand,isracknum ,img_basic) "&_
						" VALUES ("&evt_code&",'"&eCategory&"','"&selCM&"','"&issale&"','"&isgift&"'"&_
						",'"&israck&"','"&isprize&"','"&partMDid&"','"&evt_forward&"','"&brand&"','"&isracknum&"','"&html2db(img_basic)&"'"&_
						") "
						
				'response.write strSql
				dbget.execute strSql
			END IF
		END IF
		rsget.close
	
		 '-이벤트 상태에 따른 할인,사은품,쿠폰 상태 변경---------------		
		IF (evt_state < 7) THEN  	'오픈전 상태 발급대기로 등록
			istatus = 0
		ELSEIF (evt_state <9) THEN
			istatus = 7
		ELSE
			istatus = evt_state
		END IF
		'--------------------------------------------------------------
	
		'--gift 확인
		Dim strgift	: strgift = ""
		IF isgift <> "Y" THEN 
			strgift = ", gift_using = 'N' "
		end if
		
		strSql =" SELECT count(gift_code) FROM [db_shop].[dbo].[tbl_gift_off] WHERE evt_code = "&evt_code&" AND gift_using ='Y' "
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			igiftcnt = rsget(0)
		END IF
		rsget.close

		if igiftcnt > 0 then
		strSql =" UPDATE [db_shop].[dbo].[tbl_gift_off] Set gift_name = '"&evt_name&"'" &_ 
				" ,gift_startdate ='"&evt_startdate&"', gift_enddate ='"&evt_enddate&"', gift_status= "	&istatus&strAdd&_
				" ,lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&strgift&_
				" WHERE evt_code = "&evt_code
		dbget.execute strSql
		end if	

		'-- sale 확인
		'Dim strSale	: strSale = ""
		'Dim arrSale,intSale

		'IF issale <> "Y" THEN strSale = ", sale_using = 0 "
		'strSql = " SELECT sale_code, sale_status FROM [db_shop].[dbo].[tbl_sale_off] WHERE evt_code = "&evt_code&" AND sale_using =1 "
		'rsget.Open strSql, dbget
		'IF not (rsget.EOF or rsget.BOF) THEN
		'	arrSale = rsget.getRows()
		'END IF
		'rsget.close

		'IF isarray(arrSale)  THEN
		'	For intSale = 0 To UBound(arrSale,2)
		'	'세일의 경우 오픈상태값 6, 종료상태값 8 이므로 상태값 조정 필요
		'	if (evt_state = 7 AND arrSale(1,intSale) >= 6) OR ( evt_state > 7 AND arrSale(1,intSale) >= 8 )  THEN		istatus = arrSale(1,intSale)
		'		strSql ="	UPDATE [db_shop].[dbo].[tbl_sale_off] Set sale_name = '"&evt_name&"', sale_startdate ='"&evt_startdate&"', sale_enddate ='"&evt_enddate&"', sale_status="	&istatus&strAdd&_
		'				"			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&strSale&_
		'				"		WHERE evt_code = "&evt_code&" and sale_code = "&arrSale(0,intSale)
		'		dbget.execute strSql
		'	Next
		'END IF
		
	end if

	IF Err.Number = 0 THEN
		dbget.CommitTrans		
		
		''' 추가 매장 관련. =================================
		if evt_code="" then evt_code=tmpevt_code
		
	    addshopid = Trim(addshopid)
	    if Left(addshopid,1)<>"," then addshopid = ","+ addshopid
	    addshopid = shopid + addshopid
	    
	    addshopid = "'" & replace(replace(addshopid," ",""),",","','") & "'"
	    ''rw addshopid
	    
	    strSql = " delete from db_shop.dbo.tbl_event_off_AssignedShop"&VbCRLF
        strSql = strSql & " where evt_code="&evt_code&VbCRLF
        strSql = strSql & " and AssignShopid not in ("&VbCRLF
        strSql = strSql & " 	select userid from db_shop.dbo.tbl_shop_user"&VbCRLF
        strSql = strSql & " 	where userid in ("&addshopid&")"&VbCRLF
        strSql = strSql & " )"

		'response.write strSql & "<br>"
        dbget.Execute strSql

        strSql = " insert into db_shop.dbo.tbl_event_off_AssignedShop"&VbCRLF
        strSql = strSql & " (evt_code,AssignShopid)"&VbCRLF
        strSql = strSql & " select "&evt_code&",U.userid from db_shop.dbo.tbl_shop_user U"&VbCRLF
        strSql = strSql & " 	left join db_shop.dbo.tbl_event_off_AssignedShop A"&VbCRLF
        strSql = strSql & " 	on A.evt_code="&evt_code&""&VbCRLF
        strSql = strSql & " 	and U.userid=A.AssignShopid"&VbCRLF
        strSql = strSql & " where U.userid in ("&addshopid&")"&VbCRLF
        strSql = strSql & " and A.evt_code is NULL "

		'response.write strSql & "<br>"
        dbget.Execute strSql
		'''=========================================================
		
		'사은품이벤트이나 사은품이 등록이 안된경우 경고처리
		IF 	(isgift = "Y" AND igiftcnt < 1) THEN	
			response.write "<script>"
			response.write "alert('저장되었습니다.\n\n사은품 등록이 필요합니다. 사은품을 등록해주세요');"			
			response.write "opener.location.reload(); self.close();"			
			response.write "</script>"			
		ELSE
			response.write "<script>"
			response.write "alert('OK');"			
			response.write "opener.location.reload(); self.close();"
			response.write "</script>"							
		END IF
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
			response.write "<script>alert('문제발생'); self.close();"			
	END IF		

END if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->