<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event_process.asp
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2007.02.12 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
'--------------------------------------------------------
' 변수선언 & 파라미터 값 받기
'--------------------------------------------------------
Dim eMode
Dim eCode, eKind, eManager, eScope, eName, eSdate, eEdate, ePdate, eState, eCategory, eChkDisp, eTag
Dim eSale, eGift, eCoupon, eComment, eBbs, eItemps, eApply, eLevel, eDId, eMId, eFwd, eISort, eIAddType, eBrand,eusing, eOnlyTen, eisblogurl
Dim eBImg, eIcon, eMImg, eGImg,eVType,eMHtml , eLinkType, eLinkURL, eBImg2010, eBImgMo, eDispCate, eDateView , eBImgMoToday ,eBImgMo2014 , eNamesub
Dim sPartnerid, eLinkCode, eCommentTitle,sOpenDate,sCloseDate, eItemListType
Dim strSql, tmpeCode , selCM
Dim eGCode, backUrl, strparm
Dim blnFull, blnIteminfo, sWorkTag , blnItemprice, blnWide
Dim eNameEng '영문 이벤트명 추가
Dim subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary
Dim etcitemban , etcitemid, evt_sortNo , CCode

eMode 	= requestCheckVar(Request.Form("imod"),2) '데이터 처리종류
eCode  	= requestCheckVar(Request.Form("eC"),10)	'이벤트코드
CCode	= requestCheckVar(Request.Form("cC"),10)	'아이템 복사를 위한 이벤트 코드

strparm = Request.Form("strparm")
 
	eusing 		= requestCheckVar(Request.Form("using"),1)
	eKind 		= requestCheckVar(Request.Form("eventkind"),4)
	eManager 	= requestCheckVar(Request.Form("eventmanager"),4)
	eScope 		= requestCheckVar(Request.Form("eventscope"),4)
	sPartnerid 	= requestCheckVar(Request.Form("selP"),32)
	IF eScope="" THEN eScope="2"
	IF eScope="2" THEN sPartnerid = ""
	eName 		= html2db(requestCheckVar(Request.Form("sEN"),120))
	eNamesub	= html2db(requestCheckVar(Request.Form("subsEN"),100)) ' 이벤트 서브카피
	eNameEng = html2db(requestCheckVar(Request.Form("sENEng"),120)) '영문이벤트명

	subcopyK = html2db(requestCheckVar(Request.Form("subcopyK"),500)) '서브카피 한글
	subcopyE = html2db(requestCheckVar(Request.Form("subcopyE"),500)) '서브카피 영문
	If  subcopyK = "한글" Then subcopyK = ""
	If  subcopyE = "영문" Then subcopyE = ""

	eSdate		= requestCheckVar(Request.Form("sSD"),10)
	eEdate 		= requestCheckVar(Request.Form("sED"),10)
	ePdate 		= requestCheckVar(Request.Form("sPD"),10)
	eState 		= requestCheckVar(Request.Form("eventstate"),4)

	evt_sortNo	= requestCheckVar(Request.Form("sortNo"),5)	'정렬번호(회차)
	if evt_sortNo="" then evt_sortNo="0"

	eChkDisp 	= requestCheckVar(Request.Form("chkDisp"),2)
	eCategory 	= requestCheckVar(Request.Form("selC"),10)
	selCM = requestCheckVar(Request.Form("selCM"),10)
	eDispCate = requestCheckVar(Request.Form("dispcate"),12)

	eLinkCode 	= requestCheckVar(Request.Form("eLC"),10)
	IF eLinkCode = "" THEN eLinkCode = 0
	eCommentTitle 	= html2db(requestCheckVar(Request.Form("eCT"),200))
	eTag			= html2db(requestCheckVar(Replace(Request.Form("eTag")," ",""),300))
	If Right(eTag,1) = "," Then
		eTag = Left(eTag,(Len(eTag)-1))
	End If

	eSale 		= requestCheckVar(Request.Form("chSale"),2)
	IF eSale ="" THEN	eSale = 0

	eGift 		= requestCheckVar(Request.Form("chGift"),2)
	IF eGift ="" THEN	eGift = 0

	eCoupon 	= requestCheckVar(Request.Form("chCoupon"),2)
	IF eCoupon ="" THEN		eCoupon = 0

	eOnlyTen	= requestCheckVar(Request.Form("chOnlyTen"),2)
	IF eOnlyTen ="" THEN	eOnlyTen = 0

	eComment 	= requestCheckVar(Request.Form("chComm"),2)
	IF eComment ="" THEN 	eComment = 0

	eBbs 		= requestCheckVar(Request.Form("chBbs"),2)
	IF eBbs ="" THEN	eBbs = 0

	eItemps 	= requestCheckVar(Request.Form("chItemps"),2)
	IF eItemps ="" THEN		eItemps = 0

	eisblogurl	= requestCheckVar(Request.Form("isblogurl"),2)
	IF eisblogurl ="" THEN		eisblogurl = 0

	eApply 		= requestCheckVar(Request.Form("chApply"),2)
	IF eApply ="" THEN	eApply = 0

	eOneplusone 		= requestCheckVar(Request.Form("chOneplusone"),2)
	IF eOneplusone ="" THEN	eOneplusone = 0

	eFreedelivery 		= requestCheckVar(Request.Form("chFreedelivery"),2)
	IF eFreedelivery ="" THEN	eFreedelivery = 0

	eBookingsell 		= requestCheckVar(Request.Form("chBookingsell"),2)
	IF eBookingsell ="" THEN	eBookingsell = 0

	eDiary 		= requestCheckVar(Request.Form("chDiary"),2)
	IF eDiary ="" THEN	eDiary = 0

	eLevel 		= requestCheckVar(Request.Form("eventlevel"),4)
	eDId 		= requestCheckVar(Request.Form("selDId"),32)
	eMId 		= requestCheckVar(Request.Form("selMId"),32)
	eFwd 		= html2db(Trim(Request.Form("tFwd")))
	eISort 		= requestCheckVar(Request.Form("itemsort"),4)
	sWorkTag	= requestCheckVar(Request.Form("sWorkTag"),32)

	eVType 		= Trim(Request.Form("eventview"))	'화면템플릿 종류
	IF eVType = "5" or eVType = "6" THEN
		eMHtml = html2db(Request.Form("tHtml5"))		'화면설정html 코드
	ELSE
		eMHtml = html2db(Request.Form("tHtml"))		'화면설정html 코드
		eMImg = Request.Form("main")
	END IF

	eLinkType = requestCheckVar(Request.Form("elType"),1)
	eLinkURL = requestCheckVar(Request.Form("elUrl"),128)

    eBImg 		= Request.Form("ban")
    eBImg2010	= Request.Form("ban2010")
    eBImgMo		= Request.Form("banMo")
    eBImgMoToday= Request.Form("banMoToday")
    eBImgMo2014 = Request.Form("banMoList") '//2014 모바일 리스트 이미지
    eBrand 		= Request.Form("ebrand")
    eIcon 		= Request.Form("icon")
    eGImg 		= Request.Form("gift")
	blnFull		= Request.Form("chkFull")
	blnWide		= Request.Form("chkWide")
  	blnIteminfo	= Request.Form("chkIteminfo")
	blnItemprice	= Request.Form("chkItemprice")
	eDateView	= Request.Form("dateview")

	etcitemid 		= Trim(requestCheckVar(Request.Form("etcitemid"),10)) '상품정보 상품코드
	etcitemban 		= Request.Form("etcitemban") '상품정보 상품이미지
	eItemListType	= Request.Form("itemlisttype")
	
  	IF blnFull = "" 	THEN blnFull = 1
  	IF blnWide = "" 	THEN blnWide = 0
  	IF blnIteminfo = "" THEN blnIteminfo = 0
  	IF blnItemprice = "" THEN blnItemprice = 0
  	IF eDateView = "" THEN eDateView = 0

'--------------------------------------------------------
' 데이터 처리
' I : 이벤트 개요등록, U: 개요수정, disply등록/수정
'--------------------------------------------------------
SELECT Case eMode
Case "I"
	'상태가 오픈일때 오픈일 등록
	sOpenDate = "null"
	sCloseDate = "null"
	IF eState = 7 THEN
		sOpenDate = "getdate()"
	ELSEIF eState = 9 THEN
		sCloseDate = "getdate()"
	END IF
	'트랜잭션 (1.master등록/2.disply등록)
	dbget.beginTrans
		'--1.master등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event] (evt_kind, evt_manager, evt_scope, partner_id,evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_level, evt_state, opendate, closedate, evt_lastupdate, adminid,evt_nameEng,evt_subcopyK,evt_subcopyE,evt_sortNo , evt_subname ) "&_
			"		VALUES ("&eKind&","&eManager&","&eScope&",'"&sPartnerid&"','"&eName&"','"&eSdate&"','"&eEdate&"','"&ePdate&"',"&eLevel&","&eState&","&sOpenDate&","&sCloseDate&",getdate(),'"&session("ssBctId")&"','"&eNameEng&"','"&subcopyK&"','"&subcopyE&"',"&evt_sortNo&" , '"& eNamesub &"')"
		dbget.execute strSql

		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_event] "		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
		strSql = "select SCOPE_IDENTITY()"

		rsget.Open strSql, dbget, 0
		tmpeCode = rsget(0)
		rsget.Close

		'--2.disply등록
		IF eChkDisp = "on" THEN
			strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display] (evt_code,evt_category,evt_cateMid,issale,isgift,iscoupon,iscomment,isbbs,isitemps,isapply, evt_itemsort,designerid, partMDid, evt_forward, brand, evt_comment, link_evtcode,evt_bannerlink,evt_LinkType,evt_tag,isOnlyTen,isGetBlogURL,workTag,isOneplusone,isFreedelivery,isbookingsell,evt_dispCate,evt_dateview,evt_itemlisttype, isDiary) "&_
				"		VALUES ("&tmpeCode&",'"&eCategory&"','"&selCM&"',"&eSale&","&eGift&","&eCoupon&","&eComment&","&eBbs&","&eItemps&","&eApply&","&eISort&",'"&eDId&"','"&eMId&"','"&eFwd&"','"&eBrand&"','"&eCommentTitle&"','"&eLinkCode&"','"&eLinkURL &"','"&eLinkType&"','"&eTag&"','"&eOnlyTen&"','"&eisblogurl&"','"&sWorkTag&"',"&eOneplusone&","&eFreedelivery&","&eBookingsell&",'"&eDispCate&"','"&eDateView&"','"&eItemListType&"','"&eDiary&"')"
			dbget.execute strSql
		END IF

	IF Err.Number = 0 THEN
		dbget.CommitTrans
		IF strparm = "" THEN strparm = "eventkind="&eKind
		IF 	(egift = 1 AND igiftcnt < 1) THEN	'사은품이벤트이나 사은품이 등록이 안된경우 경고처리
			Call sbAlertMsg ("저장되었습니다.\n\n사은품 등록이 필요합니다. 사은품을 등록해주세요",  "index.asp?menupos="&menupos&"&"&strparm, "self")
		ELSE
			Call sbAlertMsg ("저장되었습니다.",  "index.asp?menupos="&menupos&"&"&strparm, "self")
		END IF
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
	END IF

CASE "U"
	Dim strAdd : strAdd = ""
 	eGCode = split(Request.Form("selG"),",")
 	sOpenDate = requestCheckVar(Request.Form("eOD"),30)
 	sCloseDate =requestCheckVar(Request.Form("eCD"),30)

 	IF (eState = 7 and sOpenDate ="" ) THEN 	'오픈처리일 설정
		strAdd = ", [opendate] = getdate() "
	ELSEIF (eState = 9 and sCloseDate ="" ) THEN
		strAdd = ", [closedate] = getdate() "	'종료처리일 설정
	END IF

	'종료일 이전에 종료시 종료일 현재 날짜로 변경
	IF eState = 9 and  datediff("d",eEdate,date()) <0 THEN
			eEdate = date()
	END IF

	'트랜잭션 (1.master수정/2.disply수정)
	dbget.beginTrans

	'--1.master수정
	strSql = " UPDATE [db_event].[dbo].[tbl_event] "&_
			 "	SET  [evt_kind]="&eKind&", [evt_manager]="&eManager&", [evt_scope]="&eScope&", [partner_id]='"&sPartnerid&"',[evt_name]='"&eName&"'"&_
			 " 		, [evt_startdate]='"&eSdate&"', [evt_enddate]='"&eEdate&"',[evt_prizedate]='"&ePdate&"', [evt_level]="&eLevel&", [evt_state]="&eState&", [evt_using] = '"&eusing&"'"&_
			 "		, evt_lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&_
			 "		, evt_nameEng = '"&eNameEng&"' ,evt_subcopyK = '"&subcopyK&"' , evt_subcopyE = '"&subcopyE&"'"&strAdd&_
			 "		, evt_sortNo=" &evt_sortNo&_
			 "		, evt_subname='" & eNamesub &"'" &_
			 "  WHERE evt_code = "&eCode
	dbget.execute strSql

	'--2.disply수정
	strSql = "SELECT evt_code FROM [db_event].[dbo].[tbl_event_display] WHERE evt_code= "&eCode
	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		IF eChkDisp = "on" THEN
			strSql = "UPDATE [db_event].[dbo].[tbl_event_display] SET "&_
					" 	evt_category ='"&eCategory&"',evt_cateMid ='"&selCM&"',issale="&eSale&",isgift="&eGift&",iscoupon="&eCoupon&",iscomment="&eComment&","&_
					"	isbbs="&eBbs&",isitemps="&eItemps&",isapply="&eApply&", evt_itemsort="&eISort&", designerid='"&eDId&"', partMDid='"&eMId&"', evt_forward='"&eFwd&"',"&_
					" 	evt_bannerimg = '"&eBImg&"', evt_giftimg='"&eGImg&"', evt_template = "&eVType&", evt_mainimg = '"&eMImg&"', evt_html='"&eMHtml&"', brand='"&eBrand&"', evt_icon = '"&eIcon&"'"&_
					"	,evt_comment = '"&eCommentTitle&"', link_evtcode ="&eLinkCode&", evt_fullyn="&blnFull&", evt_wideyn="&blnWide&", evt_iteminfoyn= "&blnIteminfo&_
					"	,evt_bannerlink = '"&eLinkURL&"', evt_LinkType ='"&eLinkType &"', evt_tag = '" & eTag & "', evt_bannerimg2010 = '"&eBImg2010&"', isOnlyTen = '"&eOnlyTen&"' " &_
					"	,isGetBlogURL = '"&eisblogurl&"', workTag='"&sWorkTag&"', evt_itempriceyn='"&blnItemprice&"', evt_bannerimg_mo = '"&eBImgMo&"' , isOneplusone = "&eOneplusone&" ,isFreedelivery = "&eFreedelivery&" ,isBookingsell = "&ebookingsell&" " & _
					"	,etc_itemid = '"&etcitemid&"', etc_itemimg='"&etcitemban&"', evt_dispCate='"&eDispCate&"', evt_dateview='"&eDateView&"' , evt_todaybanner='"& eBImgMoToday &"' , evt_mo_listbanner='"& eBImgMo2014 &"', evt_itemlisttype='"&eItemListType&"', isDiary='"&eDiary&"' " & _
					"	WHERE evt_code ="&eCode
			
			'response.write strSql
			dbget.execute strSql
		ELSE
			strSql = " DELETE FROM  [db_event].[dbo].[tbl_event_display]  WHERE  evt_code ="&eCode
			dbget.execute strSql
		END IF
	ELSE
		IF eChkDisp = "on" THEN
			strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display] "&_
					" (evt_code, evt_category,evt_cateMid,issale,isgift,iscoupon,iscomment,isbbs,isitemps,isapply, evt_itemsort, designerid, partMDid, evt_forward, evt_bannerimg,evt_template,evt_mainimg,evt_html, brand, evt_icon,evt_bannerlink,evt_LinkType,evt_tag, evt_bannerimg2010, isOnlyTen, isGetBlogURL,workTag, evt_bannerimg_mo,isOneplusone,isFreedelivery,etc_itemid,etc_itemimg,isBookingsell,evt_dispCate,evt_dateview , evt_todaybanner , evt_mo_listbanner, evt_itemlisttype, isDiary) "&_
					" VALUES "&_
					"("&eCode&",'"&eCategory&"','"&selCM&"',"&eSale&","&eGift&","&eCoupon&","&eComment&","&eBbs&","&eItemps&","&eApply&","&eISort&",'"&eDId&"','"&eMId&"','"&eFwd&"','"&eBImg&"', "&eVType&", '"&eMImg&"', '"&eMHtml&"','"&eBrand&"','"&eIcon&"','"&eLinkURL&"','"&eLinkType&"','"&eTag&"','"&eBImg2010&"', '"&eOnlyTen&"','"&eisblogurl&"','"&sWorkTag&"','"&eBImgMo&"',"&eOneplusone&","&eFreedelivery&",'"&etcitemid&"','"&etcitemban&"',"&ebookingsell&",'"&eDispCate&"','"&eDateView&"' , '"& eBImgMoToday &"' , '"& eBImgMo2014 &"', '"&eItemListType&"','"&eDiary&"') "
			dbget.execute strSql
		END IF
	END IF
	rsget.close

	 '-이벤트 상태에 따른 할인,사은품,쿠폰 상태 변경---------------
	Dim istatus
		IF (eState < 7) THEN  	'오픈전 상태 발급대기로 등록
			istatus = 0
		ELSEIF (eState <9) THEN
			istatus = 7
		ELSE
			istatus = eState
		END IF
	'--------------------------------------------------------------

	'--gift 확인
	Dim strgift	: strgift = ""
	Dim igiftcnt : igiftcnt = 0
	Dim isAllGiftEvent : isAllGiftEvent = False
	
	IF egift = 0 THEN strgift = ", gift_using = 'N' "

		strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_using ='Y' "
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			igiftcnt = rsget(0)
		END IF
		rsget.close
        
        ''전체 사은 이벤트 인지 CHECK
        strSql =" SELECT count(gift_code) FROM [db_event].[dbo].[tbl_gift] WHERE evt_code = "&eCode&" AND gift_scope in (1,9) "
        rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			isAllGiftEvent = rsget(0)>0
		END IF
		rsget.close
        
        '전체사은/다이어리사은은 강제 종료되면 안됨.
        if (isAllGiftEvent) then
            strgift = ""
        end if
        
		if igiftcnt > 0 then
		strSql ="	UPDATE [db_event].[dbo].[tbl_gift] Set gift_name = '"&eName&"', makerid ='"&eBrand&"' ,gift_startdate ='"&eSdate&"', gift_enddate ='"&eEdate&"', gift_status= "	&istatus&strAdd&_
				"			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"', site_scope= "&eScope&", partner_id='"&sPartnerid&"' "&strgift&_
				"		WHERE evt_code = "&eCode
	    
	    if (istatus=0) then ''전체사은/다이어리사은은 강제 종료되면 안됨.
		    strSql = strSql&"  and gift_scope not in (1,9)"  
	    end if
	    
		dbget.execute strSql
		end if

	'-- sale 확인
	Dim strSale	: strSale = ""
	Dim arrSale,intSale

		IF eSale = 0 THEN strSale = ", sale_using = 0 "
		strSql = " SELECT sale_code, sale_status FROM [db_event].[dbo].[tbl_sale] WHERE evt_code = "&eCode&" AND sale_using =1 "
		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			arrSale = rsget.getRows()
		END IF
		rsget.close

		IF isarray(arrSale)  THEN
			For intSale = 0 To UBound(arrSale,2)
			'세일의 경우 오픈상태값 6, 종료상태값 8 이므로 상태값 조정 필요
			if (eState = 7 AND arrSale(1,intSale) >= 6) OR ( eState > 7 AND arrSale(1,intSale) >= 8 )  THEN		istatus = arrSale(1,intSale)
				strSql ="	UPDATE [db_event].[dbo].[tbl_sale] Set sale_name = '"&eName&"', sale_startdate ='"&eSdate&"', sale_enddate ='"&eEdate&"', sale_status="	&istatus&strAdd&_
						"			, lastupdate = getdate(), adminid = '"&session("ssBctId")&"'"&strSale&_
						"		WHERE evt_code = "&eCode&" and sale_code = "&arrSale(0,intSale)
				dbget.execute strSql
			Next
		END IF

	IF Err.Number = 0 THEN
		dbget.CommitTrans
		IF strparm = "" THEN strparm = "eventkind="&eKind
		IF 	(egift = 1 AND igiftcnt < 1) THEN	'사은품이벤트이나 사은품이 등록이 안된경우
			Call sbAlertMsg ("저장되었습니다.\n\n사은품 등록이 필요합니다. 사은품을 등록해주세요","index.asp?menupos="&menupos&"&"&strparm, "self")
		ELSE
			Call sbAlertMsg ("저장되었습니다.","index.asp?menupos="&menupos&"&"&strparm, "self")
		END IF
		dbget.close()	:	response.End
	ELSE
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
	END IF

CASE "gD"	'그룹삭제
	eGCode= Request("eGC")

	strSql = "UPDATE [db_event].[dbo].[tbl_eventitem_group] SET evtgroup_using ='N' " &_
				"	WHERE evtgroup_code = "&eGCode&" OR evtgroup_pcode ="&eGCode
	dbget.execute strSql

	IF Err.Number = 0 THEN
		Call sbAlertMsg ("삭제되었습니다.", "iframe_eventitem_group.asp?eC="&eCode&"&menupos="&menupos, "self")
	ELSE
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
	END IF
	dbget.close()	:	response.End

Case "IT" '아이템복사 2014-05-16 이종화

	Dim cnt , gcnt , tempi , tempii

	'//그룹개수
	strSql = "select count(*) as gcnt " + VbCrlf
	strSql = strSql + " from db_event.dbo.tbl_eventitem_group " + VbCrlf
	strSql = strSql + " where evtgroup_using = 'Y' and evt_code = " + Ccode

	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		gcnt = rsget("gcnt")
	END IF
	rsget.close

	If gcnt > 0 Then '// 그룹이 있을 경우
		dbget.beginTrans '//트렌젝션
		'//화면템플릿 업데이트
		strSql = "update db_event.dbo.tbl_event_display set " + VbCrlf
		strSql = strSql + " evt_template = (select evt_template from db_event.dbo.tbl_event_display where evt_code = "&CCode&") "  + VbCrlf
		strSql = strSql + " where evt_code= " + eCode 
		dbget.execute strSql

		IF Err.Number = 0 Then
			'//그룹 일단 다 복사
			strSql = " insert into db_event.dbo.tbl_eventitem_group " + VbCrlf 
			strSql = strSql + " (evt_code , evtgroup_desc , evtgroup_sort , evtgroup_img , evtgroup_link " + VbCrlf
			strSql = strSql + " , evtgroup_pcode , evtgroup_depth , evtgroup_using) " + VbCrlf
			strSql = strSql + " select '"& eCode &"', t.evtgroup_desc  , t.evtgroup_sort , t.evtgroup_img , t.evtgroup_link  " + VbCrlf
			strSql = strSql + " , t.evtgroup_pcode , t.evtgroup_depth , t.evtgroup_using " + VbCrlf
			strSql = strSql + " From db_event.dbo.tbl_eventitem_group as t " + VbCrlf
			strSql = strSql + " where t.evt_code = '"& CCode &"'" 

			dbget.execute strSql

			IF Err.Number = 0 Then
				'//후에 그룹코드 변경 업데이트
				strSql = " update b set " + VbCrlf
				strSql = strSql + " b.evtgroup_pcode = (select top 1 c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c where c.evt_code = '"& eCode &"' and c.evtgroup_depth = a.evtgroup_depth ) " + VbCrlf
				strSql = strSql + " from db_event.dbo.tbl_eventitem_group as a " + VbCrlf
				strSql = strSql + " inner join " + VbCrlf
				strSql = strSql + " db_event.dbo.tbl_eventitem_group as b " + VbCrlf
				strSql = strSql + " on a.evtgroup_code = b.evtgroup_pcode " + VbCrlf
				strSql = strSql + " where b.evt_code = '"& eCode &"'"

				dbget.execute strSql

				IF Err.Number = 0 Then
					'//상품 그룹복사 전체
					strSql = " insert into [db_event].[dbo].tbl_eventitem " + VbCrlf
					strSql = strSql + " (evt_code,itemid,evtgroup_code,evtitem_sort , evtitem_imgsize) " + VbCrlf
					strSql = strSql + " select '"& eCode &"', i.itemid, i.evtgroup_code ,50 , 200 " + VbCrlf
					strSql = strSql + " from [db_event].[dbo].tbl_eventitem i " + VbCrlf
					strSql = strSql + " where evt_code= '"& CCode &"' " + VbCrlf
					strSql = strSql + " and itemid not in " + VbCrlf
					strSql = strSql + " (select itemid from [db_event].[dbo].tbl_eventitem " + VbCrlf
					strSql = strSql + " where evt_code= '"& eCode &"' " + VbCrlf
					strSql = strSql + " ) "

					dbget.execute strSql
					
					IF Err.Number = 0 Then
						'//상품 그룹복사 - 그룹코드 전체 변경
						strSql = " update i Set " + VbCrlf
						strSql = strSql + " i.evtgroup_code =  " + VbCrlf
						strSql = strSql + " (select top 1 c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c  " + VbCrlf
						strSql = strSql + " 	where c.evt_code = '"& eCode &"'  " + VbCrlf
						strSql = strSql + " 	and c.evtgroup_depth = a.evtgroup_depth " + VbCrlf
						strSql = strSql + " ) " + VbCrlf
						strSql = strSql + " from [db_event].[dbo].tbl_eventitem as i " + VbCrlf
						strSql = strSql + " inner Join " + VbCrlf
						strSql = strSql + " db_event.dbo.tbl_eventitem_group as a " + VbCrlf
						strSql = strSql + " on i.evtgroup_code = a.evtgroup_code " + VbCrlf
						strSql = strSql + " where i.evt_code = '"& eCode &"'"
						dbget.execute strSql

						IF Err.Number = 0 Then
							dbget.CommitTrans					
							Response.write "<script>alert('상품이 복사 되었습니다.');</script>"
							Response.write "<script>parent.opener.location.reload();</script>"
							Response.write "<script>parent.self.close();</script>"
							dbget.close()	:	response.End
						Else
							dbget.RollBackTrans
							Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
						END IF
					Else
						dbget.RollBackTrans
						Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
					END IF
				Else
					dbget.RollBackTrans
					Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
				END IF
			Else 
				dbget.RollBackTrans
				Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
			END IF
		Else
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
		END IF

	Else '// 그룹이 없을경우 상품만 복사
		'//상품개수
		strSql = "select count(*) as cnt " + VbCrlf
		strSql = strSql + " from [db_event].[dbo].tbl_eventitem i "  + VbCrlf
		strSql = strSql + " where evt_code= " + CCode 
		strSql = strSql + " and itemid not in " + VbCrlf
		strSql = strSql + " (select itemid from [db_event].[dbo].tbl_eventitem " + VbCrlf
		strSql = strSql + " where evt_code= " + eCode 
		strSql = strSql + " ) " 

		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			cnt = rsget("cnt")
		END IF
		rsget.close

	'	Response.write cnt
	'	Response.end
		
		If cnt > 0 Then 
		dbget.beginTrans '//트렌젝션

			strSql = " insert into [db_event].[dbo].tbl_eventitem " + VbCrlf
			strSql = strSql + " (evt_code,itemid,evtgroup_code,evtitem_sort) " + VbCrlf
			strSql = strSql + " select " + CStr(eCode) + ", i.itemid, '0' ,50 " + VbCrlf
			strSql = strSql + " from [db_event].[dbo].tbl_eventitem i "  + VbCrlf
			strSql = strSql + " where evt_code= " + CCode 
			strSql = strSql + " and itemid not in " + VbCrlf
			strSql = strSql + " (select itemid from [db_event].[dbo].tbl_eventitem " + VbCrlf
			strSql = strSql + " where evt_code= " + eCode 
			strSql = strSql + " ) " 

			dbget.execute strSql

			IF Err.Number = 0 Then
				dbget.CommitTrans
				Response.write "<script>alert('상품이 복사 되었습니다.');</script>"
				Response.write "<script>parent.self.close();</script>"
				dbget.close()	:	response.End
			Else
				dbget.RollBackTrans
				Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
			END IF
		Else
			Call sbAlertMsg ("이미 상품이 복사 되었습니다.", "back", "")
		End If 
	
	End If 

CASE "IC"	'그룹삭제 2014-10-29 이종화
	
	'//그룹개수
	strSql = "select count(*) as gcnt " + VbCrlf
	strSql = strSql + " from db_event.dbo.tbl_eventitem_group " + VbCrlf
	strSql = strSql + " where evtgroup_using = 'Y' and evt_code = " + eCode

	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		gcnt = rsget("gcnt")
	END IF
	rsget.close

	If gcnt > 0 Then 
		dbget.beginTrans '//트렌젝션

		strSql = "delete from [db_event].[dbo].[tbl_eventitem_group] " &_
					"	WHERE evt_code= " & eCode 
		dbget.execute strSql
		IF Err.Number = 0 Then
			'//상품개수
			strSql = "select count(*) as cnt " + VbCrlf
			strSql = strSql + " from [db_event].[dbo].tbl_eventitem i "  + VbCrlf
			strSql = strSql + " where evt_code= " + eCode 

			rsget.Open strSql, dbget
			IF not (rsget.EOF or rsget.BOF) THEN
				cnt = rsget("cnt")
			END IF
			rsget.close

			If cnt > 0 Then '//그룹은 있는데 상품이 없는 경우 
				'//상품도 삭제
				strSql = "delete from [db_event].[dbo].[tbl_eventitem] " &_
						"	WHERE evt_code= " & eCode 
				dbget.execute strSql
				IF Err.Number = 0 Then
					dbget.CommitTrans		
					Response.write "<script>alert('삭제되었습니다.');</script>"
					Response.write "<script>parent.location.reload();</script>"
				Else
					dbget.RollBackTrans
					Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
				END If
			Else
				dbget.CommitTrans		
				Response.write "<script>alert('삭제되었습니다.');</script>"
				Response.write "<script>parent.location.reload();</script>"
			End If 
		Else
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
		END IF
		dbget.close()	:	response.End

	Else '//그룹은 없고 상품만 있는 경우

		'//상품개수
		strSql = "select count(*) as cnt " + VbCrlf
		strSql = strSql + " from [db_event].[dbo].tbl_eventitem i "  + VbCrlf
		strSql = strSql + " where evt_code= " + eCode 

		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			cnt = rsget("cnt")
		END IF
		rsget.close

		If cnt > 0 Then 
			dbget.beginTrans '//트렌젝션

			strSql = "delete from [db_event].[dbo].[tbl_eventitem] " &_
					"	WHERE evt_code= " & eCode 
			dbget.execute strSql
			IF Err.Number = 0 Then
				dbget.CommitTrans		
				Response.write "<script>alert('삭제되었습니다.');</script>"
				Response.write "<script>parent.location.reload();</script>"
			Else
				dbget.RollBackTrans
				Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
			END If

		Else 
			Call sbAlertMsg ("삭제할 상품이 없습니다.", "back", "")
		End If 
		dbget.close()	:	response.End
	End If

CASE Else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
END SELECT
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->