<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
 Response.AddHeader "Pragma","no-cache"
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"

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
Dim eMode, vChangeContents, vSCMChangeSQL
'vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf
Dim eCode, eKind, eManager, eScope, eName, eSdate, eEdate, ePdate, eState, eCategory, eChkDisp, eTag
Dim eSale, eGift, eCoupon, eComment, eBbs, eItemps, eApply, eLevel, eDId, eMId, eFwd, eFwdMo, eISort, eIAddType, eBrand,eusing, eOnlyTen, eisblogurl, eisNew,eDiary
Dim eBImg, eIcon, eMImg, eGImg,eVType,eMHtml , eLinkType, eLinkURL, eBImg2010, eBImgMo, eDispCate, eDateView , eBImgMoToday ,eBImgMo2014 , eNamesub,eVType_mo
Dim sPartnerid, eLinkCode, eCommentTitle,sOpenDate,sCloseDate, eItemListType, sImgregdate, eCCId, nocate
Dim strSql, tmpeCode , selCM
Dim eGCode, backUrl, strparm
Dim blnFull, blnIteminfo, sWorkTag , blnItemprice, blnWide
Dim eNameEng '영문 이벤트명 추가
Dim subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell
Dim etcitemban , etcitemid, evt_sortNo , CCode
Dim eDgId,edgid2,edgstat1,edgstat2, eMdId, ePsId, eDpId
Dim isWeb, isMobile, isApp
dim eMHtml_mo, eMImg_mo
Dim CmtType,eCmtMT,eCmtST,eIpsMT,eIpsST,eGfMT,eGfST,eBSMT,eBSST,chkeCmt,chkeIps,chkeGf, chkeBS
dim blnReqPublish,blnexec,blnexec_mo,eexecfile,eexecfile_mo
dim eSalePer , sgroup_m , sgroup_w, eSaleCPer
Dim evt_tagkind , evt_tagopt1 , etc_opt1 , etc_opt2 '// 모바일 & 앱 상품이벤트 추가
Dim eSlideYN_W , eSlideYN_M '//슬라이드사용 유무 추가
Dim tmpType, endlessview
Dim eType, isConfirm, videoLink, videoFullLink, eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end
Dim DFcolorCD, DFcolorCD2, DFcolorCDMo, DFcolorCDMo2, mdtheme, mdthememo, bntype, bntypemo
Dim comm_isusing, comm_text, freebie_img, comm_start, comm_end, gift_isusing, gift_text1, gift_img1, gift_text2
Dim gift_img2, gift_text3, gift_img3, usinginfo, using_text1, using_contents1, using_text2, using_contents2, using_text3, using_contents3, upback, title_pc, title_mo

comm_isusing = requestCheckVar(Request.Form("comm_isusing"),1)
comm_text = Request.Form("comm_text")
freebie_img = Request.Form("freebie_img")
comm_start = requestCheckVar(Request.Form("comm_start"),10)
comm_end = requestCheckVar(Request.Form("comm_end"),10)
gift_isusing = requestCheckVar(Request.Form("gift_isusing"),1)
gift_text1 = Request.Form("gift_text1")
gift_img1 = Request.Form("gift_img1")
gift_text2 = Request.Form("gift_text2")
gift_img2 = Request.Form("gift_img2")
gift_text3 = Request.Form("gift_text3")
gift_img3 = Request.Form("gift_img3")
usinginfo = requestCheckVar(Request.Form("usinginfo"),1)
using_text1 = Request.Form("using_text1")
using_contents1 = Request.Form("using_contents1")
using_text2 = Request.Form("using_text2")
using_contents2 = Request.Form("using_contents2")
using_text3 = Request.Form("using_text3")
using_contents3 = Request.Form("using_contents3")
upback = requestCheckVar(Request.Form("upback"),1)
title_pc = Request.Form("title_pc")
title_mo = Request.Form("title_mo")
endlessview = requestCheckVar(Request.Form("endlessview"),1)
videoLink = requestCheckVar(Request.Form("videoLink"),256)
eval_isusing = requestCheckVar(Request.Form("eval_isusing"),1)
eval_text = Request.Form("eval_text")
eval_freebie_img = Request.Form("eval_freebie_img")
eval_start = requestCheckVar(Request.Form("eval_start"),10)
eval_end = requestCheckVar(Request.Form("eval_end"),10)

if title_pc <> "" then
	if checkNotValidHTML(title_pc) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	title_pc = Replace(title_pc, "'", "")
end If
if title_mo <> "" then
	if checkNotValidHTML(title_mo) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	title_mo = Replace(title_mo, "'", "")
end if
if comm_text <> "" then
	if checkNotValidHTML(comm_text) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	comm_text = Replace(comm_text, "'", "")
end if
if freebie_img <> "" then
	if checkNotValidHTML(freebie_img) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	freebie_img = Replace(freebie_img, "'", "")
end If
if eval_text <> "" then
	if checkNotValidHTML(eval_text) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	eval_text = Replace(eval_text, "'", "")
end if
if eval_freebie_img <> "" then
	if checkNotValidHTML(eval_freebie_img) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	eval_freebie_img = Replace(eval_freebie_img, "'", "")
end if
if gift_text1 <> "" then
	if checkNotValidHTML(gift_text1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_text1 = Replace(gift_text1, "'", "")
end if
if gift_img1 <> "" then
	if checkNotValidHTML(gift_img1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_img1 = Replace(gift_img1, "'", "")
end if
if gift_text2 <> "" then
	if checkNotValidHTML(gift_text2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_text2 = Replace(gift_text2, "'", "")
end if
if gift_img2 <> "" then
	if checkNotValidHTML(gift_img2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_img2 = Replace(gift_img2, "'", "")
end if
if gift_text3 <> "" then
	if checkNotValidHTML(gift_text3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_text3 = Replace(gift_text3, "'", "")
end if
if gift_img3 <> "" then
	if checkNotValidHTML(gift_img3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	gift_img3 = Replace(gift_img3, "'", "")
end if
if using_text1 <> "" then
	if checkNotValidHTML(using_text1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_text1 = Replace(using_text1, "'", "")
end if
if using_contents1 <> "" then
	if checkNotValidHTML(using_contents1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_contents1 = Replace(using_contents1, "'", "")
end if
if using_text2 <> "" then
	if checkNotValidHTML(using_text2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_text2 = Replace(using_text2, "'", "")
end if
if using_contents2 <> "" then
	if checkNotValidHTML(using_contents2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_contents2 = Replace(using_contents2, "'", "")
end if
if using_text3 <> "" then
	if checkNotValidHTML(using_text3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_text3 = Replace(using_text3, "'", "")
end if
if using_contents3 <> "" then
	if checkNotValidHTML(using_contents3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
	using_contents3 = Replace(using_contents3, "'", "")
end if

eMode 	= requestCheckVar(Request.Form("imod"),2) '데이터 처리종류
eCode  	= requestCheckVar(Request.Form("eC"),10)	'이벤트코드
CCode	= requestCheckVar(Request.Form("cC"),10)	'아이템 복사를 위한 이벤트 코드
eSalePer = requestCheckVar(Request.Form("sSP"),8)
eSaleCPer = requestCheckVar(Request.Form("sCSP"),8)
strparm = Request.Form("strparm")
nocate = requestCheckVar(request.Form("nocate"),1)
If nocate="" Then nocate="N"

DFcolorCD  		= requestCheckVar(Request.Form("DFcolorCD"),3)
DFcolorCD2  	= requestCheckVar(Request.Form("DFcolorCD2"),3)
DFcolorCDMo  	= requestCheckVar(Request.Form("DFcolorCDMo"),3)
DFcolorCDMo2  	= requestCheckVar(Request.Form("DFcolorCDMo2"),3)
mdtheme			= requestCheckVar(Request.Form("mdtheme"),2)
mdthememo	 	= requestCheckVar(Request.Form("mdthememo"),2)
bntype	 		= requestCheckVar(Request.Form("bntype"),1)
bntypemo	 	= requestCheckVar(Request.Form("bntypemo"),1)
eType		= requestCheckVar(Request.Form("eventtype"),3) '// 이벤트 유형

  isWeb			= requestCheckVar(Request.Form("blnWeb"),1)
  if isWeb = "" then isWeb = 0
  isMobile	= requestCheckVar(Request.Form("blnMobile"),1)
  if isMobile = "" then isMobile = 0
  isApp			= requestCheckVar(Request.Form("blnApp"),1)
 if isApp = "" then isApp = 0
	eusing 		= requestCheckVar(Request.Form("using"),1)
	eKind 		= requestCheckVar(Request.Form("eventkind"),4)
	eManager 	= requestCheckVar(Request.Form("eventmanager"),4)
	if eManager ="" then eManager =1
	eScope 		= requestCheckVar(Request.Form("eventscope"),4)
	eScope 	 =2
	sPartnerid 	= requestCheckVar(Request.Form("selP"),32)
	IF eScope="2" THEN sPartnerid = ""
	eName 		= html2db(stripHTML(requestCheckVar(Request.Form("sEN"),120)))

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
	eDispCate = requestCheckVar(Request.Form("disp"),12)

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

	eisNew			=requestCheckVar(Request.Form("chNew"),2)
	IF eisNew ="" THEN	eisNew = 0

	ediary			= requestCheckVar(Request.Form("chDiary"),2)
	if 	ediary = "" then ediary =0

	eLevel 		= requestCheckVar(Request.Form("eventlevel"),4)

	eDgId 		= requestCheckVar(Request.Form("sDgId"),32)
	eDgId2 		= requestCheckVar(Request.Form("sDgId2"),32)
	if Request.Form("designerstatus")<>"" then
		edgstat1	= requestCheckVar(Request.Form("designerstatus")(1),2)
		edgstat2	= requestCheckVar(Request.Form("designerstatus")(2),2)
	end if

	eMdId 		= requestCheckVar(Request.Form("sMdId"),32)
	ePsId 		= requestCheckVar(Request.Form("sPsId"),32)
	eDpId 		= requestCheckVar(Request.Form("sDpId"),32)
	eCCId		= requestCheckVar(Request.Form("sCCId"),32)

	eFwd 		= html2db(Trim(Request.Form("tFwd")))
	eFwdMo		= html2db(Trim(Request.Form("tFwdMo")))
	eISort 		= requestCheckVar(Request.Form("itemsort"),4)
	sWorkTag	= requestCheckVar(Request.Form("sWorkTag"),32)

	eVType 		= requestCheckVar(Request.Form("eventview"),1)	'화면템플릿 종류
	eVType_mo      = requestCheckVar(Request.Form("eventview_mo"),1)

	If eType = "9" Then eVType=9
	If eType = "9" Then eVType_mo=9

	IF eVType = "5" or eVType = "6" THEN
		eMHtml = html2db(Request.Form("tHtml5"))		'화면설정html 코드
	ELSE
		eMHtml = html2db(Request.Form("tHtml"))		'화면설정html 코드
		eMImg = Request.Form("main")
	END IF

	If eSalePer <> "" Then
		if ((eKind = "1" or  ekind="23" ) and (eSale = "1" or eCoupon="1") and (eSalePer <> "" or eSalePer <> "0" )) then eName = eName &"|"& "~"&Cstr(eSalePer) &"%"
	Elseif eSaleCPer<>"" And eSalePer="" Then
		if ((eKind = "1" or  ekind="23" ) and (eSale = "1" or eCoupon="1") and (eSaleCPer <> "" or eSaleCPer <> "0" )) then eName = eName &"|"& "~"&Cstr(eSaleCPer) &"%"
    End If

    IF eVType_mo = "5" or eVType_mo = "6" THEN
		eMHtml_mo = html2db(Request.Form("tHtml5_mo"))
	ELSE
		eMHtml_mo = html2db(Request.Form("tHtml_mo"))
		eMImg_mo = Request.Form("main_mo")
	END If

	If eSalePer <> "" Then
		eSalePer=Replace(eSalePer,"~","")
		eSalePer=Replace(eSalePer,"%","")
	End If

	If eSaleCPer <> "" Then
		eSaleCPer=Replace(eSaleCPer,"~","")
		eSaleCPer=Replace(eSaleCPer,"%","")
	End If

	eSlideYN_W	= requestCheckVar(Request.Form("slide_w_flag"),1)	'슬라이드 사용/pc
	eSlideYN_M	= requestCheckVar(Request.Form("slide_m_flag"),1)	'슬라이드 사용/mo

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

	chkeCmt = requestCheckVar(Request.Form("chkeCmt"),1)
	chkeIps = requestCheckVar(Request.Form("chkeIps"),1)
	chkeGf  = requestCheckVar(Request.Form("chkeGf"),1)
	chkeBS  = requestCheckVar(Request.Form("chkeBS"),1)

	CmtType = requestCheckVar(Request.Form("rdCmt"),1)
	eCmtMT = requestCheckVar(Request.Form("eCmtMT"),200)
	eCmtST = html2db(Request.Form("eCmtST"))
	eIpsMT = requestCheckVar(Request.Form("eIpsMT"),200)
	eIpsST = html2db(Request.Form("eIpsST"))
	eGfMT = requestCheckVar(Request.Form("eGfMT"),200)
	eGfST = html2db(Request.Form("eGfST"))
	eBSMT = requestCheckVar(Request.Form("eBSMT"),200)
	eBSST = html2db(Request.Form("eBSST"))

	blnReqPublish = requestCheckVar(Request.Form("chkReqP"),1)

	blnexec     = requestCheckVar(Request.Form("rdoEF"),1)
	blnexec_mo  = requestCheckVar(Request.Form("rdoEF_mo"),1)

	sgroup_w	= requestCheckVar(Request.Form("sgroup_W"),1) '// 최상위 랜덤노출 - 웹
	sgroup_m	= requestCheckVar(Request.Form("sgroup_M"),1) '// 최상위 랜덤노출 - 모바일


	isConfirm	= requestCheckVar(Request.Form("blnCnfm"),1) '// 관리자 확인 (유형이 50-커스텀형일때)

	If sgroup_w = "" Then sgroup_w = 0
	If sgroup_m = "" Then sgroup_m = 0

  	IF blnFull = "" 	THEN blnFull = 1
  	IF blnWide = "" 	THEN blnWide = 0
  	IF blnIteminfo = "" THEN blnIteminfo = 0
  	IF blnItemprice = "" THEN blnItemprice = 0
  	IF eDateView = "" THEN eDateView = 0
    IF blnReqPublish = "" THEN blnReqPublish = 0
    IF blnexec = "" THEN blnexec = 0
    IF blnexec_mo = "" THEN blnexec_mo = 0
    IF eType = "" THEN eType = 10
    IF isConfirm = "" THEN isConfirm = 0

    if blnexec = "1" then
	    eexecfile   =  requestCheckVar(Request.Form("sEFP"),128)
    else
        eexecfile = ""
    end if
    if blnexec_mo = "1" then
	    eexecfile_mo=  requestCheckVar(Request.Form("sEFP_mo"),128)
    else
        eexecfile_mo = ""
    end If
    '//이벤트 모바일 & 앱 상품이벤트 추가
	evt_tagkind = requestCheckVar(Request.Form("ietag"),10)
	evt_tagopt1 = requestCheckVar(Request.Form("ietagval"),10)
	etc_opt1 = requestCheckVar(Request.Form("mcopy"),100)
	etc_opt2 = requestCheckVar(Request.Form("scopy"),200)

'	Response.write chkeCmt &"<br/>"
'	Response.write chkeIps &"<br/>"
'	Response.write chkeGf  &"<br/>"
'	Response.write chkeBS  &"<br/>"
'	Response.End

	'// 2016.2.16 신규추가 상품상세설명 동영상 추가 - 원승현
	'// 2016-12-13  iframe 없는 경우 - iframe 생성 삽입
	'// 아이템 동영상 값 정규식으로 src, width, height값 뽑아냄
	If Trim(videoLink) <> "" Then
		Dim itemvideo, RetStr, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType, dbsql
		itemvideo = videoLink
		itemvideo = itemvideo & "?rel=0"
		'// 2016-12-13 추가 iframe 없이 주소만 넘어 올경우
		If InStr(itemvideo ,"iframe") > 0 Then
		else
			'// 비디오 변환 및 기본형 (유투브인지 비메오인지)
			If InStr(itemvideo , "youtu.be")>0 Then
				itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://youtu.be/","https://www.youtube.com/embed/")) &""" frameborder=""0"" allowfullscreen></iframe>"
			ElseIf InStr(itemvideo, "vimeo")>0 Then
				itemvideo = "<iframe src="""& Trim(Replace(itemvideo,"https://vimeo.com/","https://player.vimeo.com/video/")) &""" frameborder=""0"" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>"
			End If
		End If

		itemvideo = Trim(Replace(itemvideo,"""","'"))
		'// iframe 이외의 코드는 잘라버림
		itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

		'// 비디오 타입지정(유투브인지 비메오인지)
		If InStr(itemvideo, "youtube")>0 Then
			videoType = "youtube"
		ElseIf InStr(itemvideo, "vimeo")>0 Then
			videoType = "vimeo"
		Else
			videoType = "etc"
		End If

		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True

		regEx.pattern = "<iframe [^<>]*>"
		Set Matches = regEx.execute(itemvideo)
		For Each Match In Matches
			VideoTempSrc =  Mid(Match.Value, InStrRev(Match.Value,"src='")+5)
			RetSrc = Left(VideoTempSrc, InStr(VideoTempSrc, "'")-1)

			VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
			RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)

			VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
			RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
		Next
		Set regEx = Nothing
		Set Matches = Nothing

		videoFullLink=chrbyte(html2db(itemvideo),255,"")
	End If

'--------------------------------------------------------
' 데이터 처리
' I : 이벤트 개요등록, U: 개요수정, disply등록/수정
'--------------------------------------------------------
SELECT Case eMode
Case "I"
	'상태가 오픈일때 오픈일 등록
	sOpenDate = "null"
	sCloseDate = "null"
	sImgregdate = "null"

	IF eState = 7 THEN
		sOpenDate = "getdate()"
	ELSEIF eState = 9 THEN
		sCloseDate = "getdate()"
	ELSEIF eState = 3 THEN
	    sImgregdate = "getdate()"
	END IF

	'트랜잭션 (1.master등록/2.disply등록)
	dbget.beginTrans
		'--1.master등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event] (evt_kind, evt_manager, evt_scope, partner_id,evt_name, evt_startdate, evt_enddate, evt_prizedate, evt_level, evt_state, opendate, closedate, evt_lastupdate, adminid,evt_nameEng,evt_subcopyK,evt_subcopyE,evt_sortNo , evt_subname, isWeb, isMobile, isApp ,evt_imgregdate, evt_type, isConfirm) "&vbCrlf&_
			"		VALUES ("&eKind&","&eManager&","&escope&",'"&sPartnerid&"','"&eName&"','"&eSdate&"','"&eEdate&"','"&ePdate&"',"&eLevel&","&eState&","&sOpenDate&","&sCloseDate&",getdate(),'"&session("ssBctId")&"','"&eNameEng&"','"&subcopyK&"','"&subcopyE&"',"&evt_sortNo&" , '"& eNamesub &"',"&isWeb&","&isMobile&","&isApp&","&sImgregdate&","&eType&","&isConfirm&")"
		dbget.execute strSql

	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		response.End
	END IF

		'strSql = "select SCOPE_IDENTITY() From [db_event].[dbo].[tbl_event] "		'/사용금지.전체 라인 몽땅 뿌려짐. '/2016.06.02 한용민
		strSql = "select SCOPE_IDENTITY()"

		rsget.Open strSql, dbget, 0
		eCode = rsget(0)
		rsget.Close

		'--2.disply등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display] "&_
				" (evt_code,evt_category,evt_cateMid , evt_dispCate, brand"&_
				"	,issale,isgift,iscoupon,isOnlyTen,isOneplusone,isFreedelivery,isbookingsell, isDiary,isNew,iscomment,isbbs,isitemps,isapply,isGetBlogURL "&_
				"	,evt_itemsort,designerid, partMDid, publisherid, developerid,workTag,evt_tag, link_evtcode, evt_comment "&_
				"	,evt_forward,evt_forward_mo,evt_fullyn, evt_wideyn, evt_iteminfoyn, evt_itempriceyn,evt_dateview,evt_itemlisttype,evt_bannerlink,evt_LinkType, isReqPublish ,  evt_sgroup_w  , evt_sgroup_m , evt_slide_w_flag , evt_slide_m_flag, codecheckerid "&_
				"	,designerid2, dsn_state1, dsn_state2, mdtheme, mdthememo, themecolor, themecolormo, textbgcolor, textbgcolormo, mdbntype, mdbntypemo, SalePer, SaleCPer, endlessview, videoLink, videoFullLink)" & vbCrlf&_
				" VALUES ("&eCode&",'"&eCategory&"','"&selCM&"','"&eDispCate&"','"&eBrand&"'"&_
				" ,"&eSale&","&eGift&","&eCoupon&",'"&eOnlyTen&"',"&eOneplusone&","&eFreedelivery&","&eBookingsell&","&ediary&","&eisNew&","&eComment&","&eBbs&","&eItemps&","&eApply&",'"&eisblogurl&"'"&_
				" ,"&eISort&",'"&eDgId&"','"&eMdId&"','"&ePsId&"','"&eDpId&"','"&sWorkTag&"','"&eTag&"','"&eLinkCode&"','"&eCommentTitle&"'"&_
				" ,'"&eFwd&"','"&eFwdMo&"',"&blnFull&","&blnWide&","&blnIteminfo&","&blnItemprice&",'"&eDateView&"','"&eItemListType&"','"&eLinkURL &"','"&eLinkType&"',"&blnReqPublish&" , "& sgroup_w &" , "& sgroup_m &" , '"& eSlideYN_W &"' , '"& eSlideYN_M &"', '"& eCCId &"'" &_
				" ,'"&edgid2&"','"&edgstat1&"','"&edgstat2&"','"&mdtheme&"','"&mdthememo&"','"&DFcolorCD&"','"&DFcolorCDMo&"','"&DFcolorCD2&"','"&DFcolorCDMo2&"','"&bntype&"','"&bntypemo&"','" & eSalePer &"','" & eSaleCPer &"','" & endlessview & "','" & videoLink & "','" & videoFullLink &"')"
		dbget.execute strSql

	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
		response.End
	END IF

		'--3.MD 등록 테마 정보 등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_md_theme] (evt_code, comm_isusing, comm_text, freebie_img, comm_start, comm_end, gift_isusing, gift_text1, gift_img1, gift_text2, gift_img2, gift_text3, gift_img3, usinginfo, using_text1, using_contents1, using_text2, using_contents2, using_text3, using_contents3, nocate, title_pc, title_mo, eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end) "&vbCrlf&_
			"		VALUES ("&eCode&",'"&comm_isusing&"','"&comm_text&"','"&freebie_img&"','"&eSdate&"','"&eEdate&"','"&gift_isusing&"','"&gift_text1&"','"&gift_img1&"','"&gift_text2&"','"&gift_img2&"','"&gift_text3&"','"&gift_img3&"','"&usinginfo&"','"&using_text1&"','"&using_contents1&"','"&using_text2&"' , '"& using_contents2 &"','"&using_text3&"','"&using_contents3&"','"&nocate&"','"&title_pc&"','"&title_mo&"','"&eval_isusing&"','"&eval_text&"','"&eval_freebie_img&"','"&eSdate&"','"&eEdate&"')"
		dbget.execute strSql

	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[3]", "back", "")
		response.End
	END IF

	vChangeContents = vChangeContents & "이벤트 INSERT " & vbCrLf
	vChangeContents = vChangeContents & "- 이벤트명 : evt_name = " & eName & ", evt_code = " & eCode & vbCrLf
	vChangeContents = vChangeContents & "- 종류 : evt_kind = " & eKind & vbCrLf
	vChangeContents = vChangeContents & "- 타입 : 할인issale = " & eSale & ", 사은품isgift = " & eGift & ", 쿠폰iscoupon = " & eCoupon & ", isOnlyTen = " & eOnlyTen & ","
	vChangeContents = vChangeContents & " isOneplusone = " & eOneplusone & ", 무료배송isFreedelivery = " & eFreedelivery & ", 예약판매isbookingsell = " & eBookingsell & ","
	vChangeContents = vChangeContents & " isDiary = " & ediary & ", 런칭isNew = " & eisNew & vbCrLf
	vChangeContents = vChangeContents & "- 기능 : 코멘트iscomment = " & eComment & ", 게시판isbbs = " & eBbs & ", 상품후기isitemps = " & eItemps & ", Blog URL isGetBlogURL = " & eisblogurl & vbCrLf
	vChangeContents = vChangeContents & "- 기간 : evt_startdate ~ evt_enddate = " & eSdate & " ~ " & eEdate & vbCrLf
	vChangeContents = vChangeContents & "- 당첨발표일 : evt_prizedate = " & ePdate & vbCrLf
	vChangeContents = vChangeContents & "- 상태 : evt_state = " & eState & vbCrLf
	vChangeContents = vChangeContents & "- 중요도 : evt_level = " & eLevel & vbCrLf
	vChangeContents = vChangeContents & "- 브랜드 : brand = " & eBrand & vbCrLf
	vChangeContents = vChangeContents & "- 상품정렬방법 : evt_itemsort = " & eISort & vbCrLf
	vChangeContents = vChangeContents & "- 담당자 : 기획자 = " & eMdId & ", 디자이너(P) = " & eDgId & ", 디자이너(M) = " & eDgId2 & ", 퍼블리셔 = " & ePsId & ", 개발자 = " & eDpId & ", 개발검수자 = " & eCCId & ", 퍼블리싱요청 = " & blnReqPublish & ", 디자이너작업구분 = " & sWorkTag & "" & vbCrLf
	vChangeContents = vChangeContents & "- 상품정보 : evt_iteminfoyn = " & blnIteminfo & ", 상품가격정보 evt_itempriceyn = " & blnItemprice & ", 이벤트기간노출여부 evt_dateview = " & eDateView & "" & vbCrLf

	'------텍스트 타이틀(모바일)--------------------------------------------
	if CmtType = "" then CmtType = 1
	IF chkeCmt <> "0" THEN '코멘트는 ttType =1, 테스터는 2
	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = "&CmtType&" )"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eCmtMT&"', subTitle = '"&eCmtST&"'"&vbCrlf
	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = "&CmtType&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
	strSql = strSql& " VALUES("&eCode&","&CmtType&",'"&eCmtMT&"','"&eCmtST&"')"
	 dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	END IF

 	IF chkeIps <> "0" THEN '상품후기는 ttType =3
	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 3 )"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eIpsMT&"', subTitle = '"&eIpsST&"'"&vbCrlf
	strSql = strSql&"	WHERE  evt_code = "&eCode&" and ttType = 3 "&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
	strSql = strSql&" VALUES("&eCode&",3,'"&eIpsMT&"','"&eIpsST&"')"
	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	END IF

	IF chkeGf <> "0" THEN '기프트는 ttType =4
	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 4 )"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eGfMT&"', subTitle = '"&eGfST&"'"&vbCrlf
	strSql = strSql&  "	WHERE  evt_code = "&eCode&" and ttType = 4"&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
	strSql = strSql&  " VALUES("&eCode&",4 ,'"&eGfMT&"','"&eGfST&"')"
	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	END IF

	IF chkeBS <> "0" THEN ' 예약판매는 ttType =5
	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 5)"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eBSMT&"', subTitle = '"&eBSST&"'"&vbCrlf
	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 5"&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
	strSql = strSql& " VALUES("&eCode&",5,'"&eBSMT&"','"&eBSST&"')"
	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	END If
  '---------------------------------------------

  '================ 이벤트 모바일 상품이벤트 =================
  '2015-11-04 이종화 추가
	If eKind = "13" And (isMobile Or isApp) Then '상품이벤트
	strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_mobile_addetc where evt_code = "&eCode&" )"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_mobile_addetc SET evt_tagkind = '"& evt_tagkind &"', evt_tagopt1 = '"& evt_tagopt1 &"' , etc_opt1 = '"& etc_opt1 &"' , etc_opt2 = '"& etc_opt2 &"'  "&vbCrlf
	strSql = strSql& "	WHERE  evt_code = "&eCode&" "&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_mobile_addetc (evt_code, evt_tagkind , evt_tagopt1 , etc_opt1 , etc_opt2 )"&vbCrlf
	strSql = strSql& " VALUES("&eCode&", '"& evt_tagkind &"','"& evt_tagopt1 &"','"& etc_opt1 &"','"& etc_opt2 &"')"
	dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("상품이벤트 옵션 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	End if
  '===========================================================
		dbget.CommitTrans

    	'### 수정 로그 저장(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)

		IF strparm = "" THEN strparm = "eventkind="&eKind
		IF 	(egift = 1 AND igiftcnt < 1) THEN	'사은품이벤트이나 사은품이 등록이 안된경우 경고처리
			Call sbAlertMsg ("저장되었습니다.\n\n사은품 등록이 필요합니다. 사은품을 등록해주세요",  "index.asp?menupos="&menupos&"&"&strparm, "self")
		ELSE
			Call sbAlertMsg ("저장되었습니다.",  "index.asp?menupos="&menupos&"&"&strparm, "self")
		END IF
		response.End
CASE "U"
	Dim strAdd : strAdd = ""
	dim strAdd1 : strAdd1 = ""
 	eGCode = split(Request.Form("selG"),",")
 	sOpenDate = requestCheckVar(Request.Form("eOD"),30)
 	sCloseDate =requestCheckVar(Request.Form("eCD"),30)
    sImgregdate=requestCheckVar(Request.Form("eIRD"),30)

 	IF (eState = 7 and sOpenDate ="" ) THEN 	'오픈처리일 설정
		strAdd = ", [opendate] = getdate() "
	ELSEIF (eState = 9 and sCloseDate ="" ) THEN
		strAdd = ", [closedate] = getdate() "	'종료처리일 설정
	END IF

	IF (eState = 3 and sImgregdate ="" ) THEN
		strAdd1 = ", [evt_imgregdate] = getdate() "	'이미지등록일 설정
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
			 "		, evt_nameEng = '"&eNameEng&"' ,evt_subcopyK = '"&subcopyK&"' , evt_subcopyE = '"&subcopyE&"'"&strAdd&strAdd1&_
			 "		, evt_sortNo=" &evt_sortNo&_
			 "		, evt_subname='" & eNamesub &"'" &_
			 "		, isWeb = "&isWeb&", isMobile="&isMobile&" , isApp="&isApp&_
			 "		, evt_type = "&eType&", isConfirm="&isConfirm&_
			 "  WHERE evt_code = "&eCode
	dbget.execute strSql

	'--2.disply수정
	strSql = "SELECT evt_code FROM [db_event].[dbo].[tbl_event_display] WHERE evt_code= "&eCode
	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
			strSql = "UPDATE [db_event].[dbo].[tbl_event_display] SET "&_
					" 	evt_category ='"&eCategory&"',evt_cateMid ='"&selCM&"', evt_dispCate='"&eDispCate&"', brand='"&eBrand&"' "&_
					"   ,issale="&eSale&",isgift="&eGift&",iscoupon="&eCoupon&", isOnlyTen = '"&eOnlyTen&"',isOneplusone = "&eOneplusone&" ,isFreedelivery = "&eFreedelivery&" ,isBookingsell = "&ebookingsell&" , isDiary='"&eDiary&"', isNew ='"&eisNew&"'"&_
					"	,iscomment="&eComment&",isbbs="&eBbs&",isitemps="&eItemps&",isapply="&eApply&",isGetBlogURL = '"&eisblogurl&"'"&_
					"	,evt_itemsort="&eISort&", designerid='"&eDgId&"', partMDid='"&eMdId&"', publisherid='"&ePsId&"', developerid='"&eDpId&"', codecheckerid = '"&eCCId&"'"&_
					"	,workTag='"&sWorkTag&"', evt_tag = '" & eTag & "' , link_evtcode ="&eLinkCode&", evt_forward='"&eFwd&"', evt_forward_mo='"&eFwdMo&"'"&_
					"	,evt_comment = '"&eCommentTitle&"', evt_fullyn="&blnFull&", evt_wideyn="&blnWide&", evt_iteminfoyn= "&blnIteminfo&",evt_itempriceyn='"&blnItemprice&"', evt_dateview='"&eDateView&"'"&_
					"	,etc_itemid = '"&etcitemid&"', etc_itemimg='"&etcitemban&"' ,evt_bannerlink = '"&eLinkURL&"', evt_LinkType ='"&eLinkType &"', evt_mo_listbanner='"& eBImgMo2014 &"', evt_itemlisttype='"&eItemListType&"'" &_
					"	,evt_bannerimg = '"&eBImg&"', evt_giftimg='"&eGImg&"', evt_template = '"&eVType&"', evt_mainimg = '"&eMImg&"', evt_html='"&eMHtml&"', evt_mainimg_mo = '"&eMImg_mo&"', evt_html_mo='"&eMHtml_mo&"', evt_icon = '"&eIcon&"', evt_bannerimg2010 = '"&eBImg2010&"'  " & _
					"	,evt_bannerimg_mo = '"&eBImgMo&"'   , evt_todaybanner='"& eBImgMoToday &"' , isReqPublish = "&blnReqPublish &", evt_template_mo = '"&eVType_mo & "'" & _
					"   ,evt_isExec = "&blnexec&", evt_execFile= '"&eexecfile&"',evt_isExec_mo = "&blnexec_mo&", evt_execFile_mo= '"&eexecfile_mo&"', evt_sgroup_w = "& sgroup_w &" , evt_sgroup_m = "& sgroup_m&_
					"   ,evt_slide_w_flag = '"& eSlideYN_W &"' , evt_slide_m_flag =  '"& eSlideYN_M &"'"&_
					"   ,designerid2 = '"& edgid2 &"' , dsn_state1 =  '"& edgstat1  &"' , dsn_state2 =  '"& edgstat2 &"'"&_
					"   ,mdtheme = '"& mdtheme &"' , mdthememo =  '"& mdthememo  &"' , themecolor =  '"& DFcolorCD &"' , themecolormo =  '"& DFcolorCDMo & "'"&_
					"	,textbgcolor =  '"& DFcolorCD2 & "' , textbgcolormo =  '" & DFcolorCDMo2 & "', mdbntype='" & bntype & "', mdbntypemo='" & bntypemo & "', SalePer='" & eSalePer & "', SaleCPer='" & eSaleCPer & "', endlessview='" & endlessview & "', videoLink='" & videoLink & "', videoFullLink='" & videoFullLink & "'"&_
					"	WHERE evt_code ="&eCode
			dbget.execute strSql
	END IF
	rsget.close

	'--3.MD 등록 테마 정보 수정
	strSql = "SELECT evt_code FROM [db_event].[dbo].[tbl_event_md_theme] WHERE evt_code= "&eCode
	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
			strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]"
			strSql = strSql & " set gift_isusing='" & gift_isusing & "'"
			strSql = strSql & " ,usinginfo='" & usinginfo & "'"
			If comm_isusing <> "" Then
			strSql = strSql & " ,comm_isusing='" & comm_isusing & "'"
			strSql = strSql & " ,comm_text='" & comm_text & "'"
			If freebie_img <> "" Then
			strSql = strSql & " ,freebie_img='" & freebie_img & "'"
			End If
			strSql = strSql & " ,comm_start='" & comm_start & "'"
			strSql = strSql & " ,comm_end='" & comm_end & "'"
			End If
			If gift_text1 <> "" Then
			strSql = strSql & " ,gift_text1='" & gift_text1 & "'"
			End If
			If gift_img1 <> "" Then
			strSql = strSql & " ,gift_img1='" & gift_img1 & "'"
			End If
			If gift_text2 <> "" Then
			strSql = strSql & " ,gift_text2='" & gift_text2 & "'"
			End If
			If gift_img2 <> "" Then
			strSql = strSql & " ,gift_img2='" & gift_img2 & "'"
			End If
			If gift_text3 <> "" Then
			strSql = strSql & " ,gift_text3='" & gift_text3 & "'"
			End If
			If gift_img3 <> "" Then
			strSql = strSql & " ,gift_img3='" & gift_img3 & "'"
			End If
			If using_text1 <> "" Then
			strSql = strSql & " ,using_text1='" & using_text1 & "'"
			strSql = strSql & " ,using_contents1='" & using_contents1 & "'"
			End If
			If using_text2 <> "" Then
			strSql = strSql & " ,using_text2='" & using_text2 & "'"
			strSql = strSql & " ,using_contents2='" & using_contents2 & "'"
			End If
			If using_text3 <> "" Then
			strSql = strSql & " ,using_text3='" & using_text3 & "'"
			strSql = strSql & " ,using_contents3='" & using_contents3 & "'"
			End If
			strSql = strSql & " ,nocate='" & nocate & "'"
			strSql = strSql & " ,title_pc='" & title_pc & "'"
			strSql = strSql & " ,title_mo='" & title_mo & "'"
			strSql = strSql & " ,eval_isusing='" & eval_isusing & "'"
			strSql = strSql & " ,eval_text='" & eval_text & "'"
			strSql = strSql & " ,eval_freebie_img='" & eval_freebie_img & "'"
			strSql = strSql & " ,eval_start='" & eval_start & "'"
			strSql = strSql & " ,eval_end='" & eval_end & "'"
			strSql = strSql & "	WHERE evt_code ="&eCode
			dbget.execute strSql

	Else
		'--3.MD 등록 테마 정보 등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_md_theme] (evt_code, comm_isusing, comm_text, freebie_img, comm_start, comm_end, gift_isusing, gift_text1, gift_img1, gift_text2, gift_img2, gift_text3, gift_img3, usinginfo, using_text1, using_contents1, using_text2, using_contents2, using_text3, using_contents3,nocate,title_pc,title_mo, eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end) "&vbCrlf&_
			"		VALUES ("&eCode&",'"&comm_isusing&"','"&comm_text&"','"&freebie_img&"','"&comm_start&"','"&comm_end&"','"&gift_isusing&"','"&gift_text1&"','"&gift_img1&"','"&gift_text2&"','"&gift_img2&"','"&gift_text3&"','"&gift_img3&"','"&usinginfo&"','"&using_text1&"','"&using_contents1&"','"&using_text2&"' , '"& using_contents2 &"','"&using_text3&"','"&using_contents3&"','"&nocate&"','"&title_pc&"','"&title_mo&"','"&eval_isusing&"','"&eval_text&"','"&eval_freebie_img&"','"&eSdate&"','"&eEdate&"')"
		dbget.execute strSql
	END IF
	rsget.close

   	'------텍스트 타이틀(모바일)--------------------------------------------
   	if CmtType = "" then CmtType = 1
 '  	If CmtType = 1 Then tmpType = 2 Else tmpType = 1 End If ??
	IF chkeCmt <> "0" THEN '코멘트는 ttType =1, 테스터는 2
    	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType in (1,2))"&vbCrlf
    	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eCmtMT&"', subTitle = '"&eCmtST&"' , isusing = 1 , ttType = '"& CmtType &"' "&vbCrlf
    	strSql = strSql& "	WHERE  evt_code = "&eCode&" and  ttType in (1,2) "&vbCrlf
    	strSql = strSql& " ELSE "&vbCrlf
    	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
    	strSql = strSql& " VALUES("&eCode&","&CmtType&",'"&eCmtMT&"','"&eCmtST&"')"

	ELSE
	    strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = "&CmtType&" )"&vbCrlf
	    strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET isusing = 0 "&vbCrlf
	 	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = "&CmtType
	end if

	 dbget.execute strSql

		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF


	vChangeContents = vChangeContents & "이벤트 UPDATE " & vbCrLf
	vChangeContents = vChangeContents & "- 이벤트명 : evt_name = " & eName & ", evt_code = " & eCode & vbCrLf
	vChangeContents = vChangeContents & "- 종류 : evt_kind = " & eKind & vbCrLf
	vChangeContents = vChangeContents & "- 타입 : 할인issale = " & eSale & ", 사은품isgift = " & eGift & ", 쿠폰iscoupon = " & eCoupon & ", isOnlyTen = " & eOnlyTen & ","
	vChangeContents = vChangeContents & " isOneplusone = " & eOneplusone & ", 무료배송isFreedelivery = " & eFreedelivery & ", 예약판매isbookingsell = " & eBookingsell & ","
	vChangeContents = vChangeContents & " isDiary = " & ediary & ", 런칭isNew = " & eisNew & vbCrLf
	vChangeContents = vChangeContents & "- 기능 : 코멘트iscomment = " & eComment & ", 게시판isbbs = " & eBbs & ", 상품후기isitemps = " & eItemps & ", Blog URL isGetBlogURL = " & eisblogurl & vbCrLf
	vChangeContents = vChangeContents & "- 기간 : evt_startdate ~ evt_enddate = " & eSdate & " ~ " & eEdate & vbCrLf
	vChangeContents = vChangeContents & "- 당첨발표일 : evt_prizedate = " & ePdate & vbCrLf
	vChangeContents = vChangeContents & "- 상태 : evt_state = " & eState & vbCrLf
	vChangeContents = vChangeContents & "- 중요도 : evt_level = " & eLevel & vbCrLf
	vChangeContents = vChangeContents & "- 브랜드 : brand = " & eBrand & vbCrLf
	vChangeContents = vChangeContents & "- 상품정렬방법 : evt_itemsort = " & eISort & vbCrLf
	vChangeContents = vChangeContents & "- 담당자 : 기획자 = " & eMdId & ", 디자이너(P) = " & eDgId & ", 디자이너(M) = " & eDgId2 & ", 퍼블리셔 = " & ePsId & ", 개발자 = " & eDpId & ", 개발검수자 = " & eCCId & ", 퍼블리싱요청 = " & blnReqPublish & ", 디자이너작업구분 = " & sWorkTag & "" & vbCrLf
	vChangeContents = vChangeContents & "- 상품정보 : evt_iteminfoyn = " & blnIteminfo & ", 상품가격정보 evt_itempriceyn = " & blnItemprice & ", 이벤트기간노출여부 evt_dateview = " & eDateView & "" & vbCrLf
	vChangeContents = vChangeContents & "- 연관이벤트코드 : link_evtcode = " & eLinkCode & vbCrLf
	vChangeContents = vChangeContents & "- 대표상품정보및배너 : 대표상품코드 = " & etcitemid & ", 대표상품이미지 = " & etcitemban & vbCrLf


 	IF chkeIps  <> "0" THEN '상품후기는 ttType =3
    	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 3 )"&vbCrlf
    	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eIpsMT&"', subTitle = '"&eIpsST&"' , isusing = 1 "&vbCrlf
    	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 3 "&vbCrlf
    	strSql = strSql& " ELSE "&vbCrlf
    	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
    	strSql = strSql& " VALUES("&eCode&",3,'"&eIpsMT&"','"&eIpsST&"')"
	ELSE
	    strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 3 )"&vbCrlf
	    strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET isusing = 0 "&vbCrlf
	 	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 3"
	END IF
	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
	    END IF


	IF chkeGf <> "0" THEN '기프트는 ttType =4
    	strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 4 )"&vbCrlf
    	strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eGfMT&"', subTitle = '"&eGfST&"', isusing = 1 "&vbCrlf
    	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 4"&vbCrlf
    	strSql = strSql& " ELSE "&vbCrlf
    	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
    	strSql = strSql& " VALUES("&eCode&",4 ,'"&eGfMT&"','"&eGfST&"')"
    ELSE
	    strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 4 )"&vbCrlf
	    strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET isusing = 0 "&vbCrlf
	 	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 4"
	END IF

	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF


	IF  chkeBS <> "0" THEN ' 예약판매는 ttType =5

	    strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 5)"&vbCrlf
	    strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET MainTitle = '"&eBSMT&"', subTitle = '"&eBSST&"', isusing = 1 "&vbCrlf
	 	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 5"&vbCrlf
	    strSql = strSql& " ELSE "&vbCrlf
	    strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_TextTitle (evt_code, ttType, MainTitle, subTitle)"&vbCrlf
		strSql = strSql& " VALUES("&eCode&",5,'"&eBSMT&"','"&eBSST&"')"
	ELSE
	    strSql = "IF EXISTS(SELECT ttCode FROM db_event.dbo.tbl_event_TextTitle where evt_code = "&eCode&" and ttType = 5)"&vbCrlf
	    strSql = strSql& " UPDATE db_event.dbo.tbl_event_TextTitle SET isusing = 0 "&vbCrlf
	 	strSql = strSql& "	WHERE  evt_code = "&eCode&" and ttType = 5"
	END IF
	  dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END If

	'================ 이벤트 모바일 상품이벤트 =================
	'2015-11-04 이종화 추가
	If eKind = "13" And (isMobile Or isApp) Then '상품이벤트
	strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_mobile_addetc where evt_code = "&eCode&" )"&vbCrlf
	strSql = strSql& " UPDATE db_event.dbo.tbl_event_mobile_addetc SET evt_tagkind = '"& evt_tagkind &"', evt_tagopt1 = '"& evt_tagopt1 &"' , etc_opt1 = '"& etc_opt1 &"' , etc_opt2 = '"& etc_opt2 &"'  "&vbCrlf
	strSql = strSql& "	WHERE  evt_code = "&eCode&" "&vbCrlf
	strSql = strSql& " ELSE "&vbCrlf
	strSql = strSql& " INSERT INTO db_event.dbo.tbl_event_mobile_addetc (evt_code, evt_tagkind , evt_tagopt1 , etc_opt1 , etc_opt2 )"&vbCrlf
	strSql = strSql& " VALUES("&eCode&", '"& evt_tagkind &"','"& evt_tagopt1 &"','"& etc_opt1 &"','"& etc_opt2 &"')"
	dbget.execute strSql
		IF Err.Number <> 0 THEN
			dbget.RollBackTrans
			Call sbAlertMsg ("상품이벤트 옵션 처리에 문제가 발생하였습니다.[1]", "back", "")
			response.End
		END IF
	End if
  '---------------------------------------------

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

    	'### 수정 로그 저장(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)

		IF strparm = "" THEN strparm = "eventkind="&eKind
		IF 	(egift = 1 AND igiftcnt < 1) THEN	'사은품이벤트이나 사은품이 등록이 안된경우
			If upback="Y" Then
				Response.write "<script>parent.TnReloadThisPage();</script>"
			Else
				Call sbAlertMsg ("저장되었습니다.\n\n사은품 등록이 필요합니다. 사은품을 등록해주세요","index.asp?menupos="&menupos&"&eC="&eCode&"&"&strparm, "self")
			End If
		Else
			If upback="Y" Then
				Response.write "<script>parent.TnReloadThisPage();</script>"
			Else
				Call sbAlertMsg ("저장되었습니다.","index.asp?menupos="&menupos&"&eC="&eCode&"&"&strparm, "self")
			End If
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
		%>
		<script type="text/javascript">

		</script>
		<%

		vChangeContents = vChangeContents & "- 이벤트(" & eCode & ") 그룹 삭제처리 evtgroup_using ='N' 그룹코드 = " & eGCode & vbCrLf
    	'### 수정 로그 저장(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & eGCode & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)


		Call sbAlertMsg ("삭제되었습니다.", "iframe_eventitem_group.asp?eC="&eCode&"&menupos="&menupos, "self")
	ELSE
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
	END IF
	dbget.close()	:	response.End

Case "IT" '아이템복사 2014-05-16 이종화

	Dim cnt , gcnt , tempi , tempii, eTemplate, eTemplate_mo

	'//그룹개수
	strSql = "select count(*) as gcnt " & VbCrlf
	strSql = strSql & " from db_event.dbo.tbl_eventitem_group " & VbCrlf
	strSql = strSql & " where evtgroup_using = 'Y' and evt_code = " & Ccode

	rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		gcnt = rsget("gcnt")
	END IF
	rsget.close

    '//화면템플릿 업데이트
	strSql = " select evt_Template, case  when (evt_kind = 25 or evt_kind = 19 or evt_kind = 26) then evt_Template else evt_Template_mo end as evt_template_mo  from  db_event.dbo.tbl_event_display as d inner join db_event.dbo.tbl_event as e on d.evt_code = e.evt_code where d.evt_code = "&CCode&""
		rsget.Open strSql, dbget
	IF not (rsget.EOF or rsget.BOF) THEN
		eTemplate = rsget("evt_Template")
		if eTemplate = "" or isNull(eTemplate) then eTemplate = "NULL"
		eTemplate_mo = rsget("evt_Template_mo")
		if eTemplate_mo = "" or isNull(eTemplate_mo) then eTemplate_mo = "NULL"
	END IF
	rsget.close

	If gcnt > 0 Then '// 그룹이 있을 경우
		dbget.beginTrans '//트렌젝션

		strSql = "update db_event.dbo.tbl_event_display set " & VbCrlf
		strSql = strSql &" evt_template =  "&eTemplate&"  , evt_template_mo=  "&eTemplate_mo &" where evt_code= " & eCode
		dbget.execute strSql

		IF Err.Number = 0 Then
			'//그룹 일단 다 복사
			strSql = " insert into db_event.dbo.tbl_eventitem_group " & VbCrlf
			strSql = strSql & " (evt_code , evtgroup_desc , evtgroup_sort , evtgroup_img , evtgroup_link " & VbCrlf
			strSql = strSql & " , evtgroup_pcode , evtgroup_depth , evtgroup_using, evtgroup_desc_mo, evtgroup_sort_mo, evtgroup_img_mo,evtgroup_link_mo, evtgroup_pcode_mo, evtgroup_depth_mo , evtgroup_isDisp, evtgroup_isDisp_mo) " & VbCrlf
			strSql = strSql & " select '"& eCode &"', t.evtgroup_desc  , t.evtgroup_sort , t.evtgroup_img , t.evtgroup_link  " & VbCrlf
			strSql = strSql & " , t.evtgroup_pcode , t.evtgroup_depth , t.evtgroup_using, isNull(t.evtgroup_desc_mo,evtgroup_desc), isNull(t.evtgroup_sort_mo,t.evtgroup_sort) " & VbCrlf
			strSql = strSql & " , isNull(t.evtgroup_img_mo,t.evtgroup_img) ,isNull(t.evtgroup_link_mo,t.evtgroup_link), isNull(t.evtgroup_pcode_mo,t.evtgroup_pcode), isNull(t.evtgroup_depth_mo,t.evtgroup_depth) , isNull(t.evtgroup_isDisp, 1) , isNull(t.evtgroup_isDisp_mo,1)" & VbCrlf
			strSql = strSql & " From db_event.dbo.tbl_eventitem_group as t " & VbCrlf
			strSql = strSql & " where t.evt_code = '"& CCode &"' and t.evtgroup_using ='Y' "

			dbget.execute strSql

			IF Err.Number = 0 Then
				'//후에 그룹코드 변경 업데이트
				strSql = " update b set " & VbCrlf
				strSql = strSql & " b.evtgroup_pcode = (select c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c where c.evt_code = b.evt_code and c.evtgroup_depth = a.evtgroup_depth and c.evtgroup_using ='Y' ) " & VbCrlf
				strSql = strSql & " from db_event.dbo.tbl_eventitem_group as a " & VbCrlf
				strSql = strSql & " inner join " & VbCrlf
				strSql = strSql & " db_event.dbo.tbl_eventitem_group as b " & VbCrlf
				strSql = strSql & " on a.evtgroup_code = b.evtgroup_pcode " & VbCrlf
				strSql = strSql & " where b.evt_code = '"& eCode &"' and b.evtgroup_using='Y' and a.evtgroup_using='Y' "
			    dbget.execute strSql

               '//모바일 그룹코드 변경 업데이트
                strSql = " update b set " & VbCrlf
				strSql = strSql & " b.evtgroup_pcode_mo = (select distinct c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c where c.evt_code =  b.evt_code and c.evtgroup_depth_mo =  isNull(a.evtgroup_depth_mo,a.evtgroup_depth)  and c.evtgroup_using ='Y') " & VbCrlf
				strSql = strSql & " from db_event.dbo.tbl_eventitem_group as a " & VbCrlf
				strSql = strSql & " inner join " & VbCrlf
				strSql = strSql & " db_event.dbo.tbl_eventitem_group as b " & VbCrlf
				strSql = strSql & " on a.evtgroup_code = b.evtgroup_pcode_mo " & VbCrlf
				strSql = strSql & " where b.evt_code = '"& eCode &"'  and b.evtgroup_using='Y' and a.evtgroup_using='Y' "
				 dbget.execute strSql

                strSql = " update g set " & VbCrlf
				strSql = strSql & "  evtgroup_code_mo =  (select min(evtgroup_code) from db_event.dbo.tbl_Eventitem_Group " & VbCrlf
                strSql = strSql & "        where evt_code = g.evt_code and evtgroup_depth_mo = g.evtgroup_depth_mo group by evtgroup_depth_mo) " & VbCrlf
				strSql = strSql & " from db_event.dbo.tbl_Eventitem_Group  as g " & VbCrlf
                strSql = strSql & " where evt_code =  '"& eCode &"' and evtgroup_using='Y'" & VbCrlf
  				dbget.execute strSql

				IF Err.Number = 0 Then
					'//상품 그룹복사 전체
					strSql = " insert into [db_event].[dbo].tbl_eventitem " & VbCrlf
					strSql = strSql & " (evt_code,itemid,evtgroup_code,evtitem_sort , evtitem_imgsize,evtitem_sort_mo, evtitem_isDisp, evtitem_isDisp_mo) " & VbCrlf
					strSql = strSql & " select '"& eCode &"', i.itemid, i.evtgroup_code ,i.evtitem_sort ,i.evtitem_imgsize, isNull(i.evtitem_sort_mo,i.evtitem_sort), isNull(i.evtitem_isDisp,1), isNull(i.evtitem_isDisp_mo,1) " & VbCrlf
					strSql = strSql & " from [db_event].[dbo].tbl_eventitem i " & VbCrlf
					strSql = strSql & " where evt_code= '"& CCode &"' and evtitem_isusing ='1' " & VbCrlf
					strSql = strSql & " and itemid not in " & VbCrlf
					strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
					strSql = strSql & " where evt_code= '"& eCode &"' and evtitem_isusing ='1' " & VbCrlf
					strSql = strSql & " ) "

					dbget.execute strSql

					IF Err.Number = 0 Then
						'//상품 그룹복사 - 그룹코드 전체 변경
						strSql = " update i Set " & VbCrlf
						strSql = strSql & " i.evtgroup_code =  " & VbCrlf
						strSql = strSql & " (select c.evtgroup_code from db_event.dbo.tbl_eventitem_group as c  " & VbCrlf
						strSql = strSql & " 	where c.evt_code = '"& eCode &"'  " & VbCrlf
						strSql = strSql & " 	and c.evtgroup_depth = a.evtgroup_depth  and c.evtgroup_using='Y' " & VbCrlf
						strSql = strSql & " ) " & VbCrlf
						strSql = strSql & " from [db_event].[dbo].tbl_eventitem as i " & VbCrlf
						strSql = strSql & " inner Join " & VbCrlf
						strSql = strSql & " db_event.dbo.tbl_eventitem_group as a " & VbCrlf
						strSql = strSql & " on i.evtgroup_code = a.evtgroup_code " & VbCrlf
						strSql = strSql & " where i.evt_code = '"& eCode &"' and a.evtgroup_using='Y' and i.evtitem_isusing ='1'"
						dbget.execute strSql

						IF Err.Number = 0 Then
							dbget.CommitTrans

							vChangeContents = vChangeContents & "- 이벤트 상품 복사. " & CCode & " 상품을 " & eCode & " 로 복사" & vbCrLf
					    	'### 수정 로그 저장(event)
					    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
					    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & CCode & "', '" & menupos & "', "
					    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
					    	dbget.execute(vSCMChangeSQL)

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
		strSql = "select count(*) as cnt " & VbCrlf
		strSql = strSql & " from [db_event].[dbo].tbl_eventitem i "  & VbCrlf
		strSql = strSql & " where evt_code= " & CCode
		strSql = strSql & " and itemid not in " & VbCrlf
		strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
		strSql = strSql & " where evt_code= " & eCode & " and evtitem_isusing ='1' "&VbCrlf
		strSql = strSql & " ) and evtitem_isusing ='1' "

		rsget.Open strSql, dbget
		IF not (rsget.EOF or rsget.BOF) THEN
			cnt = rsget("cnt")
		END IF
		rsget.close

	'	Response.write cnt
	'	Response.end

		If cnt > 0 Then
		dbget.beginTrans '//트렌젝션

			strSql = " insert into [db_event].[dbo].tbl_eventitem " & VbCrlf
			strSql = strSql & " (evt_code,itemid,evtgroup_code,evtitem_sort,evtitem_imgsize, evtitem_sort_mo) " & VbCrlf
			strSql = strSql & " select " & CStr(eCode) & ", i.itemid, '0' ,evtitem_sort,i.evtitem_imgsize, isNull(i.evtitem_sort_mo, i.evtitem_sort)  " & VbCrlf
			strSql = strSql & " from [db_event].[dbo].tbl_eventitem i "  & VbCrlf
			strSql = strSql & " where evt_code= " & CCode
			strSql = strSql & " and itemid not in " & VbCrlf
			strSql = strSql & " (select itemid from [db_event].[dbo].tbl_eventitem " & VbCrlf
			strSql = strSql & " where evt_code= " & eCode
			strSql = strSql & "  and evtitem_isusing ='1' )  and evtitem_isusing ='1' "

			dbget.execute strSql

			IF Err.Number = 0 Then
				dbget.CommitTrans

				vChangeContents = vChangeContents & "- 이벤트 상품 복사. " & CCode & " 상품을 " & eCode & " 로 복사" & vbCrLf
		    	'### 수정 로그 저장(event)
		    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, sub_idx, menupos, contents, refip) "
		    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & CCode & "', '" & menupos & "', "
		    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		    	dbget.execute(vSCMChangeSQL)

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
	strSql = "select count(*) as gcnt " & VbCrlf
	strSql = strSql & " from db_event.dbo.tbl_eventitem_group " & VbCrlf
	strSql = strSql & " where evtgroup_using = 'Y' and evt_code = " & eCode

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
			'//상품도 삭제
			strSql = "delete from [db_event].[dbo].[tbl_eventitem] " &_
					"	WHERE evt_code= " & eCode
			dbget.execute strSql
			IF Err.Number = 0 Then
				dbget.CommitTrans

				vChangeContents = vChangeContents & "- 이벤트(" & eCode & ") 그룹 모두 삭제 DELETE" & vbCrLf
		    	'### 수정 로그 저장(event)
		    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
		    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		    	dbget.execute(vSCMChangeSQL)

				Response.write "<script>alert('삭제되었습니다.');</script>"
				Response.write "<script>parent.location.reload();</script>"
			Else
				dbget.RollBackTrans
				Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
			END If
		Else
			dbget.RollBackTrans
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
		END IF
		dbget.close()	:	response.End

	Else '//그룹은 없고 상품만 있는 경우

		'//상품개수
		strSql = "select count(*) as cnt " & VbCrlf
		strSql = strSql & " from [db_event].[dbo].tbl_eventitem i "  & VbCrlf
		strSql = strSql & " where evt_code= " & eCode

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

				vChangeContents = vChangeContents & "- 이벤트(" & eCode & ") 그룹 모두 삭제 DELETE" & vbCrLf
		    	'### 수정 로그 저장(event)
		    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
		    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & eCode & "', '" & menupos & "', "
		    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		    	dbget.execute(vSCMChangeSQL)

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