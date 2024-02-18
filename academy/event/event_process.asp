<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2010.09.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim eMode
Dim eCode, eKind, eManager, eScope, eName, eSdate, eEdate, ePdate, eState, eCategory, eChkDisp, eTag
Dim eSale, eGift, eCoupon, eComment, eBbs, eItemps, eApply, eLevel, eDId, eMId, eFwd, eISort, eIAddType, eBrand,eusing, eOnlyTen, eisblogurl
Dim eBImg, eIcon, eMImg, eGImg,eVType,eMHtml , eLinkType, eLinkURL, eBImg2010
Dim sPartnerid, eLinkCode, eCommentTitle,sOpenDate,sCloseDate
Dim strSql, tmpeCode , selCM
Dim eGCode, backUrl, strparm
Dim blnFull, blnIteminfo

eMode 	= requestCheckVar(Request.Form("imod"),2) '데이터 처리종류
eCode  	= requestCheckVar(Request.Form("eC"),10)	'이벤트코드
strparm = Request.Form("strparm")

	eusing 		= requestCheckVar(Request.Form("using"),1)
	eKind 		= requestCheckVar(Request.Form("eventkind"),4)
	eManager 	= requestCheckVar(Request.Form("eventmanager"),4)
	eScope 		= requestCheckVar(Request.Form("eventscope"),4)
	sPartnerid 	= requestCheckVar(Request.Form("selP"),32)
	IF eScope <> "3" THEN sPartnerid = ""
	eName 		= html2db(requestCheckVar(Request.Form("sEN"),100))
	eSdate		= requestCheckVar(Request.Form("sSD"),10)
	eEdate 		= requestCheckVar(Request.Form("sED"),10)
	ePdate 		= requestCheckVar(Request.Form("sPD"),10)
	eState 		= requestCheckVar(Request.Form("eventstate"),4)


	eChkDisp 	= requestCheckVar(Request.Form("chkDisp"),2)
	eCategory 	= requestCheckVar(Request.Form("selC"),10)
	selCM = requestCheckVar(Request.Form("selCM"),10)

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

	eLevel 		= requestCheckVar(Request.Form("eventlevel"),4)
	eDId 		= requestCheckVar(Request.Form("selDId"),32)
	eMId 		= requestCheckVar(Request.Form("selMId"),32)
	eFwd 		= html2db(Trim(Request.Form("tFwd")))
	eISort 		= requestCheckVar(Request.Form("itemsort"),4)

	eVType 		= Trim(Request.Form("eventview"))	'화면템플릿 종류
	IF eVType = "5" or eVType = "6" THEN
		eMHtml = html2db(Request.Form("tHtml5"))		'화면설정html 코드
	ELSE
		eMHtml = html2db(Request.Form("tHtml"))		'화면설정html 코드
		eMImg = Request.Form("main")
	END IF

	eLinkType = requestCheckVar(Request.Form("elType"),1)
	eLinkURL = requestCheckVar(Request.Form("elUrl"),100)

    eBImg 		= Request.Form("ban")
    eBImg2010	= Request.Form("ban2010")
    eBrand 		= Request.Form("ebrand")
    eIcon 		= Request.Form("icon")
    eGImg 		= Request.Form("gift")
	blnFull		= Request.Form("chkFull")
  	blnIteminfo	= Request.Form("chkIteminfo")
  	IF blnFull = "" 	THEN blnFull = 1
  	IF blnIteminfo = "" THEN blnIteminfo = 0

'--------------------------------------------------------
' 데이터 처리
' I : 이벤트 개요등록, U: 개요수정, disply등록/수정
'--------------------------------------------------------
SELECT Case eMode

CASE "gD"	'그룹삭제
	eGCode= Request("eGC")

	strSql = "UPDATE [db_academy].[dbo].[tbl_eventitem_group] SET evtgroup_using ='N' " &_
				"	WHERE evtgroup_code = "&eGCode&" OR evtgroup_pcode ="&eGCode
	
	'response.write strSql &"<br>"
	dbacademyget.execute strSql

	IF Err.Number = 0 THEN
		response.write "<script>"
		response.write "alert('OK');"
		response.write "location.href='iframe_eventitem_group.asp?eC="&eCode&"&menupos="&menupos&"'"
		response.write "</script>"		
	ELSE
		response.write "<script>"
		response.write "alert('데이터 처리에 문제가 발생하였습니다.[2]');"
		response.write "history.back();"
		response.write "</script>"		
	END IF
	dbget.close()	:	response.End
CASE Else
		response.write "<script>"
		response.write "alert('데이터 처리에 문제가 발생하였습니다.');"
		response.write "history.back();"
		response.write "</script>"	
END SELECT
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->