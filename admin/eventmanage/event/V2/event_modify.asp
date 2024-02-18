<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event_modify.asp
' Description :  이벤트 개요 등록
' History : 2007.02.13 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
	response.write "<script type='text/javascript'>"
	response.write "	alert('사용불가 페이지');history.back();"
	response.write "</script>"
	response.End
Dim eCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, echkdisp, eusing, etag, eonlyten, eisblogurl
Dim ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,elktype,elkurl,ebimg,etemp,emimg,ehtml,ehtml5, eisort,eiaddtype, edgid,edgid2,edgstat1,edgstat2, emdid ,efwd,ebrand,eicon,ebimg2010
Dim selPartner,dopendate,dclosedate, sWorkTag, ebimgMo, eDispCate, eDateView , ebimgToday , ebimgMo2014 , enamesub,dImgregdate, eCCId, eCCNm
Dim intI
Dim arrGift, intg,blngift
Dim eFolder, backUrl 
dim gimg : gimg = ""
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim ecommenttitle, elinkcode
Dim strparm , sCateMid
Dim cEGroup, arrGroup,arrGroup_mo, intgroup,strG, blngroup,vYear, blngroup_mo
Dim blnFull, blnIteminfo ,blnitemprice, evt_sortNo, blnWide
Dim enameEng , subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary, eNew
Dim eEtcitemid , eEtcitemimg, eItemListType
Dim eitemid, etype, isConfirm
Dim isWeb, isMobile, isApp, eDpid, ePsid, eDpnm, ePsnm, eDgnm, eDgnm2, eMdnm
dim tHtml5_mo, tHtml_mo, main_mo,emimg_mo,ehtml_mo,ehtml5_mo , efwd_mo
Dim maxDepth,dispCate
Dim eCmtCd,eIpsCd,eGfCd,eBSCd, rdCmt, eCmtMT, eCmtST, eIpsMT, eIpsST, eGfMT, eGfST, eBSMT, eBSST
dim arrText,intT
dim blnReqPublish,blnExec,eExecFile ,blnExec_mo ,eExecFile_mo  , etemp_mo
dim eSalePer
dim blnWeb,blnMobile,blnApp
dim rdIps,rdGf ,rdBS 
Dim sgroup_w , sgroup_m
Dim arrItemAdd , intA, endlessView
Dim tmpietag , tmpietagval , tmpmcopy , tmpscopy
Dim slide_w_flag , slide_m_flag , evt_m_addimg_cnt

eCode = requestCheckVar(Request("eC"),10)
ekind = requestCheckVar(Request("eK"),10)
 
maxDepth = 2 '전시카테고리 2depth까지 보여준다
eItemListType = "1"
blnIteminfo = True
isConfirm = False
  
	'## 검색 #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),120)

	sCategory	= requestCheckVar(Request("selC"),10) 		'카테고리
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	sState		= requestCheckVar(Request("eventstate"),4)	'이벤트 상태
	sKind 		= requestCheckVar(Request("eventkind"),4)	'이벤트종류
	etype		= requestCheckVar(Request("eventtype"),4)	'이벤트유형
	edgid  		= requestCheckVar(Request("sDgId"),32)		'담당 디자이너
	edgid2  	= requestCheckVar(Request("sDgId2"),32)		'서브 디자이너
	emdid  		= requestCheckVar(Request("sMdId"),32)		'담당 MD
	epsid			= requestCheckVar(Request("sPsId"),32)		'담당 퍼블리셔
	edpid			= requestCheckVar(Request("sDpId"),32)		'담당 개발
	eCCId			= requestCheckVar(Request("sCCId"),32)		'담당 개발검수자
    
    edgnm  		= requestCheckVar(Request("sdgnm"),32)		'담당 디자이너
    edgnm2 		= requestCheckVar(Request("sdgnm2"),32)		'서브 디자이너
	emdnm  		= requestCheckVar(Request("smdnm"),32)		'담당 MD
	epsnm  		= requestCheckVar(Request("spsnm"),32)		'담당 퍼블리셔
	edpnm  		= requestCheckVar(Request("sdpnm"),32)		'담당 개발자
    
	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무

	eOneplusone	 	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery	= requestCheckVar(Request("chFreedelivery"),2)	'무료배송
	eBookingsell	= requestCheckVar(Request("chBookingsell"),2)	'예약판매
	eDiary		= requestCheckVar(Request("chDiary"),2)	'다이어리
	eNew		= requestCheckVar(Request("chNew"),2)
	dispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리

	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	
	strparm  = "isWeb="&blnWeb&"&isMobile="&blnMobile&"&isApp="&blnApp&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&sCCId="&eCCId&_
				"&sdgnm="&edgnm&"&smdnm="&emdnm&"&spsnm="&epsnm&"&sdpnm="&edpnm&"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&dispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
				 
	'#######################################

	IF eCode = "" THEN	'이벤트 코드값이 없을 경우 back
		call sbAlertMsg("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오", "back","")
	END IF

	eFolder = eCode
'--------------------------------------------------------
' 이벤트 데이터 가져오기
'--------------------------------------------------------
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ekind 		=	cEvtCont.FEKind
	eman 		=	cEvtCont.FEManager
	escope 		=	cEvtCont.FEScope
	selPartner	=	cEvtCont.FEPartnerID
	ename 		=	db2html(cEvtCont.FEName)
	enamesub	=	db2html(cEvtCont.FENamesub) '이벤트 타이틀 서브카피
	enameEng =	db2html(cEvtCont.FENameEng) '이벤트 영문 추가
	subcopyK =	db2html(cEvtCont.FsubcopyK) '서브카피 한글
	subcopyE =	db2html(cEvtCont.FsubcopyE) '서브카피 영문
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	IF datediff("d",now,eeday) <0 THEN estate = 9 '기간 초과시 종료표기
	eregdate	=	cEvtCont.FERegdate
	eusing		= 	cEvtCont.FEUsing
	evt_sortNo	= 	cEvtCont.FESortNo
	eitemid		=	cEvtCont.FEitemid
	isWeb		= cEvtCont.FIsWeb
	isMobile	= cEvtCont.FIsMobile
	isApp		= cEvtCont.FIsApp
	etype		= cEvtCont.FEType
	isConfirm	= cEvtCont.FIsConfirm
	
 
	
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay 
	tmp_cdl 		= cEvtCont.FECategory
	tmp_cdm			= cEvtCont.FECateMid
	DispCate		= cEvtCont.FEDispCate

	esale 			= cEvtCont.FESale
	egift 			= cEvtCont.FEGift
	ecoupon 		= cEvtCont.FECoupon
	ecomment 		= cEvtCont.FECommnet
	ebbs 			= cEvtCont.FEBbs
	eitemps 		= cEvtCont.FEItemps
	eapply 			= cEvtCont.FEApply
	elktype			= cEvtCont.FELinkType
	IF elktype="" Then elktype="E" '//링크타입 기본값 설정
	elkurl			= cEvtCont.FELinkURL
	ebimg 			= cEvtCont.FEBImg
	ebimg2010		= cEvtCont.FEBImg2010
	ebimgMo			= cEvtCont.FEBImgMobile
	ebimgToday		= cEvtCont.FEBImgMoToday
	ebimgMo2014		= cEvtCont.FEBImgMoListBanner '//2014 모바일 리스트 배너 추가
	gimg			= cEvtCont.FEGImg
	etemp			= cEvtCont.FETemp
	etemp_mo        = cEvtCont.FETemp_mo
	if isNull(etemp) then etemp = 1
	if isNull(etemp_mo) then etemp_mo = 1
	if etemp = 5 or etemp = 6  THEN	'수작업 이벤트 일 경우 처리
		ehtml5 		= db2html(cEvtCont.FEHtml) 
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml) 
	end if
	
	if etemp_mo = 5 or etemp_mo = 6  THEN	'수작업 이벤트 일 경우 처리 
		ehtml5_mo 	= db2html(cEvtCont.FEHtml_mo)
	else 
		emimg_mo 	= cEvtCont.FEMImg_mo
		ehtml_mo 	= db2html(cEvtCont.FEHtml_mo)
	end if
	
	eisort 			= cEvtCont.FEISort
	edgid 			= cEvtCont.FEDgId
	emdid 			= cEvtCont.FEMdId
	efwd 			= db2html(cEvtCont.FEFwd)
	efwd_mo 		= db2html(cEvtCont.FEFwdMo)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode       = cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
	dImgregdate     = cEvtCont.FEImgRegdate
 	blnFull			= cEvtCont.FEFullYN
 	blnWide			= cEvtCont.FEWideYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)
 	eonlyten		= cEvtCont.FSisOnlyTen
 	eisblogurl		= cEvtCont.FSisGetBlogURL
 	sWorkTag		= cEvtCont.FWorkTag

	blnitemprice	= cEvtCont.FEItempriceYN

	eOneplusone	=	cEvtCont.FEOneplusOne
	eFreedelivery	=	cEvtCont.FEFreedelivery
	eBookingsell	=	cEvtCont.FEBookingsell
	eDiary 			= cEvtCont.FSisDiary
	eNew			= cEvtCont.FSisNew
	eEtcitemid		=	cEvtCont.FEtcitemid
	eEtcitemimg		=	cEvtCont.FEtcitemimg
	eDateView		= cEvtCont.FEdateview
	eItemListType = cEvtCont.FEListType

	edgid 			= cEvtCont.FEDgId
	edgid2 			= cEvtCont.FEDgId2
  	emdid 			= cEvtCont.FEMdId 
	epsid			= cEvtCont.FEPsId
	edpid			= cEvtCont.FEDpId
	eCCid			= cEvtCont.FECCId
	
	edgnm 			= cEvtCont.FEDgName
	edgnm2 			= cEvtCont.FEDgName2
  	emdnm 			= cEvtCont.FEMdName 
	epsnm			= cEvtCont.FEPsName
	edpnm			= cEvtCont.FEDpName
	eCCNm			= cEvtCont.FECCName

	edgstat1		= cEvtCont.FEDgStat1
	edgstat2		= cEvtCont.FEDgStat2

	blnReqPublish   = cEvtCont.FisReqPublish
	blnExec         = cEvtCont.FEisExec      
    eExecFile       = cEvtCont.FEexecFile    
    blnExec_mo      = cEvtCont.FEisExec_mo   
    eExecFile_mo    = cEvtCont.FEexecFile_mo 

	arrText			= cEvtCont.fnGetEventTextTitle

	arrItemAdd		= cEvtCont.fnGetEventMobileItemEvent '//아이템 이벤트

	sgroup_w		= cEvtCont.FEsgroup_W '// 최상위 랜덤노출 웹
	sgroup_m		= cEvtCont.FEsgroup_M '// 최상위 랜덤노출 모바일

	slide_w_flag	= cEvtCont.FESlide_W_Flag '// 슬라이드 웹
	slide_m_flag	= cEvtCont.FESlide_M_Flag '// 슬라이드 모바일
	evt_m_addimg_cnt	= cEvtCont.FEvt_m_addimg_cnt '// 모바일 추가 이미지 카운트
	endlessView = cEvtCont.FendlessView

	set cEvtCont = nothing
	IF elinkcode = 0 THEN elinkcode = ""

	 set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode   
	 	cEGroup.FEChannel = "P"    
	  	arrGroup    = cEGroup.fnGetEventItemGroup
	  	 
	    cEGroup.FEChannel = "M"        
	    arrGroup_mo    = cEGroup.fnGetEventItemGroup     
	    
	  	vYear = cEGroup.FRegdate
	 set cEGroup = nothing
 
	 blngroup = False
	 blngroup_mo = False
	 IF isArray(arrGroup) THEN blngroup = True
	 IF isArray(arrGroup_mo) THEN blngroup_mo = True

	 If eItemListType = "" OR isNull(eItemListType) Then eItemListType = "1" End If
	
		IF isArray(arrText) THEN
		For intT = 0 To UBound(arrText,2)
			IF arrText(1,intT) = 1 or arrText(1,intT) = 2 THEN
				eCmtCd = arrText(0,intT)
				rdCmt  = arrText(1,intT)		
				eCmtMT = arrText(2,intT)
				eCmtST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 3 THEN
				eIpsCd = arrText(0,intT) 
				eIpsMT = arrText(2,intT)
				eIpsST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 4 THEN
				eGfCd = arrText(0,intT) 
				eGfMT = arrText(2,intT)
				eGfST = arrText(3,intT) 
			ELSEIF  arrText(1,intT) = 5 THEN
				eBSCd = arrText(0,intT) 
				eBSMT = arrText(2,intT)
				eBSST = arrText(3,intT) 
			END IF	
		Next
	END If
	
	'//상품이벤트 모바일 & 앱
	If ekind = "13" And (isMobile Or isApp) Then
		IF isArray(arrItemAdd) Then
			For intA = 0 To UBound(arrItemAdd,2)
				tmpietag	= arrItemAdd(0,intA)
				tmpietagval = arrItemAdd(1,intA)
				tmpmcopy	= arrItemAdd(2,intA)
				tmpscopy	= arrItemAdd(3,intA)
			Next 
		End If 
	End If 

	if eCmtST = "" then
	   eCmtST = "정성껏 코멘트를 남겨주신     분을 추첨하여           를 선물로 드립니다." 
    end if
	 
	if  eCmtMT ="" or eCmtST="" then
	    chkeCmt = 0
    end if
    if  eIpsMT ="" or eIpsST="" then
	    chkeIps = 0
    end if
    if  eGfMT =""   then
	    chkeGf = 0
    end if
    if eBSMT =""  then
	    chkeBS = 0
    end if
	if (ekind = 1 or ekind=23) and (eSale or ecoupon ) then
	    dim tmpename
	    tmpename = Split(ename,"|") 
	  			 
	  	if Ubound(tmpename)>0 then
		    ename = tmpename(0)
		    eSalePer = tmpename(1)
		 end if

    end if
	 
	if eisort = "" then eisort = 3 
   
  dim idepartmentid, sdepartmentname,clsMem
    '부서명 가져오기
set clsMem = new CTenByTenMember
	clsMem.Fuserid = emdid
	clsMem.fnGetDepartmentInfo
	idepartmentid		= clsMem.Fdepartment_id
 	sdepartmentname = clsMem.FdepartmentNameFull 
 set clsMem = nothing
%>
<style>
	select {font-size:12px; vertical-align:top;}
	input[type=button], input[type=text] {vertical-align:top;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" > 
//-- jsEvtSubmit : 이벤트 등록 --//
	function jsEvtSubmit(frm){ 
		if(frm.eventkind.value == "29"){
			if(frm.sPsId.value == ""){
				alert("퍼블리셔팀에 꼭 문의를 해서 담당자를 지정해주세요.!!");
				return false;
			}
			if(frm.sDpId.value == ""){
				alert("시스템개발팀에 꼭 문의를 해서 담당자를 지정해주세요.!!");
				return false;
			}
		}
		
	    //채널선택 여부 확인
	    if (!frm.blnWeb.checked&&!frm.blnMobile.checked&&!frm.blnApp.checked){
	        alert("채널을 선택해주세요");
	        frm.blnWeb.focus();
	        return false;
	    }

	  	//유형선택 여부 확인
	  	if(!frm.eventtype.value){
		  	alert("이벤트 유형을 선택해 주세요");
		  	frm.eventtype.focus();
		  	return false;
	  	}

	  //브랜드할인이면 이벤트명 조합생성
	  if(frm.eventkind.value=='16') {
	  	if(!frm.ebrand.value){
		  	alert("브랜드를 선택해 주세요");
		  	frm.ebrand.focus();
		  	return false;
	  	}
	  	if(!frm.sEDN.value){
		  	alert("이벤트명을 입력해주세요");
		  	frm.sEDN.focus();
		  	return false;
	  	}
	  	if(frm.sMDc.value<=0){
		  	alert("최대 할인율을 입력해주세요");
		  	frm.sMDc.focus();
		  	return false;
	  	} else {
	  		frm.sEN.value = frm.sEDN.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value;
	  		frm.sENEng.value = frm.sEDNEng.value + "|" + frm.sSDc.value + "|" + frm.sMDc.value; // 영문이벤트
	  	}
	  }

	  //상품이벤트인데 대표상품 코드가 0이거나 업을때-2017-04-24 유태욱 추가
	  if(frm.eventkind.value=='13') {
		if(frm.etcitemid.value == 0 || frm.etcitemid.value == "" || isNaN(frm.etcitemid.value)){
			alert("상품이벤트일경우 대표상품코드를 넣으셔야 합니다.");
			frm.etcitemid.focus();
			return false;
		}
	  }

//	if(!frm.eventscope.value) {
//		alert("이벤트 범위를 선택해주세요");
//		frm.chkEscope[0].focus();
//		return false;
//	}

  if(frm.blnMobile.checked || frm.blnApp.checked){
        if(!frm.subsEN.value){
            alert("Mobile/App 의 서브카피를 입력해주세요");
            frm.subsEN.focus();
            return false;
        }
    }

	  if(!frm.sEN.value){
	  	alert("이벤트명을 입력해주세요");
	  	if(frm.eventkind.options[frm.eventkind.selectedIndex].value == 4){
	  	 frm.selStatic.focus();
	  	}else{
	  	 frm.sEN.focus();
	  	}
	  	return false;
	  }

	  if(frm.sENEng.value.length > 120){
		alert("영문이벤트명은 120자까지만 가능합니다.다시 입력해주세요.");
	 	frm.sENEng.focus();
	  	return false;
	  }

	if (frm.selC.value == '110'){
		if (frm.selCM.value==''){
			alert('감성채널은 중카테고리를 선택해야만 합니다');
			frm.selCM.focus();
			return false;
		}

	}

  	  if(!frm.sSD.value || !frm.sED.value ){
	  	alert("이벤트 기간을 입력해주세요");
	  	frm.sSD.focus();
	  	return false;
	  }

	  if(frm.sSD.value > frm.sED.value){
	  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
	  	frm.sED.focus();
	  	return false;
	  }

	var nowDate = jsNowDate();

	 if(frm.eventstate.value==7){
	 	if(frm.eOD.value !=""){
	 		nowDate = '<%IF dopendate <> ""THEN%><%=FormatDate(dopendate,"0000-00-00")%><%END IF%>';
		}
 	 }


	 if(frm.eventstate.value < 7){
	 	if(frm.sSD.value < nowDate){
			alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
		  	frm.sSD.focus();
		  	return false;
		 }

  	 	if(frm.sED.value < jsNowDate()){
	  		alert("종료일이 현재날짜보다 빠르면 안됩니다. 종료일을 다시 선택해주세요 ");
	  		frm.sED.focus();
	  		return false;
	  	}
	}

	if((frm.chComm.checked||frm.chBbs.checked||frm.chItemps.checked||frm.isblogurl.checked)&&frm.sPD.value=="") {
  		alert("당첨자 발표일을 선택해주세요 ");
  		frm.sPD.focus();
  		return false;
	}

	if(frm.sDgId.value!="" && frm.designerstatus[0].value==""){
  		alert("디자이너(PC) 상태를 선택해주세요.");
  		frm.designerstatus[0].focus();
  		return false;
	}
	if(frm.sDgId2.value!="" && frm.designerstatus[1].value==""){
  		alert("디자이너(Mobile) 상태를 선택해주세요.");
  		frm.designerstatus[1].focus();
  		return false;
	}

//	    if(!frm.eCT.value){
//	  		if(GetByteLength(frm.eCT.value) > 200){
//	  			alert("comment title은 200자 이내로 작성해주세요");
//	  			frm.eCT.focus();
//	  			return false;
//	  		}
//	  	}


  		if(GetByteLength(frm.eTag.value) > 250){
  			alert("Tag는 250자 이내로 작성해주세요");
  			frm.eTag.focus();
  			return false;
  		}

        if(document.all.dvCmt.style.display ==""){
            if (!frm.chkeCmt.checked &&  (!frm.eCmtMT.value ||  !frm.eCmtST.value)){
                alert("코멘트 내용을 입력해 주시거나 사용안함을 체크해주세요");
                frm.eCmtMT.focus();
                return false;
            }
        }
        
          if(document.all.dvIps.style.display ==""){ 
           if (!frm.chkeIps.checked &&  (!frm.eIpsMT.value ||  !frm.eIpsST.value)){
                alert("상품후기 내용을 입력해 주시거나 사용안함을 체크해주세요");
                frm.eIpsMT.focus();
                return false;
            }
        }
        
        
          if(document.all.dvGf.style.display ==""){
            if (!frm.chkeGf.checked && !frm.eGfMT.value ){
                alert("사은품 내용을 입력해 주시거나 사용안함을 체크해주세요");
                frm.eGfMT.focus();
                return false;
            }
        }
        
          if(document.all.dvBS.style.display ==""){
            if (!frm.chkeBS.checked && !frm.eBSMT.value ){
                alert("예약판매 내용을 입력해 주시거나 사용안함을 체크해주세요");
                frm.eBSMT.focus();
                return false;
            }
        }
        
        
	}

	function jsNowDate(){
	var mydate=new Date()
		var year=mydate.getYear()
		    if (year < 1000)
		        year+=1900

		var day=mydate.getDay()
		var month=mydate.getMonth()+1
		    if (month<10)
		        month="0"+month

		var daym=mydate.getDate()
		    if (daym<10)
		        daym="0"+daym

		return year+"-"+month+"-"+ daym
	}

//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
 


	function jsChkSubj(chk){
		if(chk=='16') {
			//브랜드할인일경우에는 제목 대신 할인율 범위로 표시
			eNameTr_A.style.display = "none";
			eNameTr_C.style.display = "none";
			eNameTr_B.style.display = "";
			eNameTr_BL.style.display= "";
		}else if(chk=='13') {
			//상품이벤트
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
			eNameTr_BL.style.display= "none";
			itemevt.style.display = ""; // 상품이벤트
		} else {
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
			eNameTr_BL.style.display= "none";
		}
		
		if(chk=='22'){
			document.all.divDE.style.display = "";
		}else{
			document.all.divDE.style.display = "none";
		}
		
		if((chk=='1'|| chk=='23')  && (document.frmEvt.chSale.checked || document.frmEvt.chCoupon.checked)){ //쇼핑찬스 일때 할인율 표기
		     document.all.spSale.style.display = "";
		     if (document.frmEvt.chSale.checked) {
		        document.all.spSale.style.color = "red";
		      }else{
		        document.all.spSale.style.color = "green";
		      }
		}else{
		     document.frmEvt.sSP.value ="";
		     document.all.spSale.style.display = "none"; 
		}
	}
	 

//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(){
	  var winLast,eKind;
	  eKind = document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value;
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
	}


	//제휴몰 표기
	function jsSetPartner(){
		if(document.frmEvt.chkEscope[0].checked&&document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="1";
		} else if(document.frmEvt.chkEscope[0].checked) {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="2";
		} else if(document.frmEvt.chkEscope[1].checked) {
			document.all.spanP.style.display ="";
			document.frmEvt.eventscope.value="3";
		} else {
			document.all.spanP.style.display ="none";
			document.frmEvt.eventscope.value="";
		}
	}

	 
	
	function jsGetID(sType, iCid, sUserID){
		var openWorker = window.open('PopWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function jsDelID(sType){ 
		eval("document.frmEvt.s"+sType+"Id").value = "";
		eval("document.frmEvt.s"+sType+"Nm").value = ""; 
	}
	
	//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){ 
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsManageEventImage(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function jsManageEventImageNew(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/V2/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}
 
 	function jsAddGroup(eCode,gCode, smode, eChannel){ 
		var winG 
		var vYear = "<%=vYear%>";  
		winG = window.open('pop_eventitem_group.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel,'popG','width=800, height=800,scrollbars=yes,resizable=yes');
		winG.focus();
	}
	
	function jsAddProcGroup(eCode, smode, sModeType,eChannel)
	{ 
	    document.frmG.target="ifrmProc";
	    document.frmG.mode.value = smode;
	    document.frmG.eCh.value = eChannel;
	    document.frmG.eMT.value = sModeType
	    document.frmG.submit();
	}

	function jsGroupImg(eCode,gCode,eChannel){
		var vYear = "<%=vYear%>";
		var winG = window.open('pop_eventitem_groupImage.asp?yr='+vYear+'&eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel,'popG','width=700, height=600,scrollbars=yes,resizable=yes');
		winG.focus();
	}

	function jsChangeFrm(iVal,sType){  
		if (sType =="P"){
		     $("div[id^='divFrm']").hide();  
    		$("#divGM").hide(); 
			$("#w_slide").hide();
    		
    		if(iVal == 3 || iVal == 7){  
    			$("#divGM").show(); 
    			$("#divFrm3").show();
				$("#w_slide").show();
    		}else if(iVal == 5 || iVal == 6 ){
    			//iframG.location.href = "about;blank"; 
    			$("#divFrm5").show(); 
    		}else{
    			//iframG.location.href = "about;blank"; 
    			$("#divFrm1").show();
				$("#w_slide").show();
    		} 
    	}else if (sType=="M") {
    	    $("div[id^='divMFrm']").hide();  
    	    $("#divGM_mo").hide();
			$("#m_slide").hide();
    		if(iVal == 3 || iVal == 7){  
    			$("#divGM_mo").show();  
    			$("#divMFrm3").show();
				$("#m_slide").show();
    		}else if(iVal == 5 || iVal == 6 ){ 
    			$("#divMFrm5").show();
    		}else{  
    			$("#divMFrm1").show();
				$("#m_slide").show();;
    		} 
    	}else if (sType=="DG1") {
			if(iVal==""){
				document.frmEvt.designerstatus[0].value = "";
			} else {
				document.frmEvt.designerstatus[0].value = "20";
			}
    	}else if (sType=="DG2") {
			if(iVal==""){
				document.frmEvt.designerstatus[1].value = "";
			} else {
				document.frmEvt.designerstatus[1].value = "20";
			}
    	}
	}
	
	//모바일 텍스트타일
	function jsChkTitle(sType){  
		if(sType=="g") { 
		 	if (document.frmEvt.chGift.checked){
				document.all.dvGf.style.display ="";
		 	}else{
		 		document.all.dvGf.style.display ="none";
			}	 
		}else if (sType=="i"){
			if (document.frmEvt.chItemps.checked){
				document.all.dvIps.style.display ="";
		 	}else{
		 		document.all.dvIps.style.display ="none";
			}	  
		}else if (sType=="b"){
			if (document.frmEvt.chBookingsell.checked){
				document.all.dvBS.style.display ="";
		 	}else{
		 		document.all.dvBS.style.display ="none";
			}	  
		}else if (sType=="c"){	
			if (document.frmEvt.chComm.checked){
				document.all.dvCmt.style.display ="";
		 	}else{
		 		document.all.dvCmt.style.display ="none";
			}	 
		}
	}
	
	function popRegItem(eCode, gCode,eChannel){
	var wImgView;
	wImgView = window.open('eventitem_regist.asp?eC='+eCode+'&selG='+gCode+'&eCh='+eChannel,'pImg','width=1400,height=800,scrollbars=yes,resizable=yes');
	wImgView.focus();
}

	function jsAddByte(obj){ 
	var realText = obj.value; 
	 var textBit = '';
	 var textLen = 0;
	 
	 for (var i = 0 ; i < realText.length ; i++) {
	  textBit = realText.charAt(i); 
	  if(escape(textBit).length > 4) {
	   textLen = textLen + 2;
	  } else {
	   textLen = textLen + 1;
	  }
	  
	  if (textLen >= 70){
	    realText = realText.substr(0,i);
	    obj.value = realText;
	    break;
	  }
	  
	 }
	
    document.frmEvt.subSize.value = textLen;  

	}
	
	// 블로그URL태그 검사(코멘트가 체크가 되어있어야 가능)
	function jsChkBlogEnable() {
		if($('#isblogurl').prop('checked') == true) {
			if($('#chComm').prop('checked') == false) {
				alert("블로그URL기능은 코멘트가 있어야만 사용가능합니다. 코멘트여부를 선택해주세요.");
				$('#isblogurl').prop('checked',false);
			}
		}
	}
	// 상품복사 리스트팝업
	function jsItemcopylist(){
		var winLast,eKind;
		winLast = window.open('pop_event_itemlist.asp?menupos=<%=menupos%>&eC=<%=eCode%>','pLast','width=800,height=600, scrollbars=yes')
		winLast.focus();
	}
	
	
	function jsChkChannel(sChannel){ 
	    if (sChannel =="P"){
	        if(document.frmEvt.blnWeb.checked){
	            document.all.divPC1.style.display="";
	            document.all.divPC2.style.display="";
	        }else{
	            document.all.divPC1.style.display="none";
	            document.all.divPC2.style.display="none";
	        }
	    }
	    if (sChannel =="M" || sChannel =="A"){
	        if(document.frmEvt.blnMobile.checked || document.frmEvt.blnApp.checked){
	            document.all.divMA1.style.display="";
	            document.all.divMA2.style.display=""; 
	        }else{
	            document.all.divMA1.style.display="none";
	            document.all.divMA2.style.display="none"; 
	        }
	    }
	}
	 
	function jsChkSale(){
	    var frm = document.frmEvt; 
	    if( (frm.eventkind.options[frm.eventkind.selectedIndex].value == 1 || frm.eventkind.options[frm.eventkind.selectedIndex].value == 23)   && (frm.chSale.checked|| frm.chCoupon.checked)){ 
	        document.all.spSale.style.display = "";
	         if (frm.chSale.checked) {
		        document.all.spSale.style.color = "red";
		      }else{
		        document.all.spSale.style.color = "green";
		      }
	    }else{
	       frm.sSP.value ="";
	        document.all.spSale.style.display = "none"; 
	    }
	}
  
   function jsPubHelp(){ 
	   	var winPop = window.open("pop_publishing.asp","popHelp","width=500,height=500,scrollbars=yes,resizable=yes");
		winPop.focus();
	}    
	
    function jsChkMBReq(){ 
	    if(document.frmEvt.chkMB.checked){  
	         document.frmEvt.sWorkTag.value = "★★" + document.frmEvt.sWorkTag.value; 
	    }else{
	          document.frmEvt.sWorkTag.value =  document.frmEvt.sWorkTag.value.replace("★★", "");
	    }
	}
		// 상품 초기화
	function jsItemclear(){
		var frm = document.frmitemclear;

		if(confirm("상품 초기화를 하시겠습니까?\n\n상품 초기화후 데이터 복구가 불가능 합니다.")){
			frm.target = "FrameCKP";
			//frm.target = "blank";
			frm.action = "/admin/eventmanage/event/event_process.asp";
			frm.submit();
		}
	}
	      
		function popCommentXLS(ecd) {
		 var wCmtXls = window.open('/admin/eventmanage/event/pop_event_Comment_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	//2015.05.19 유태욱(푸드파이터 이벤트용으로 임시 생성-이벤트 종료후 삭제예정)
	function popCommentXLS2(ecd) {
		 var wCmtXls = window.open('/admin/eventmanage/event/pop_event_Comment_xls_2.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	function popBBSXLS(ecd) {
		 var wBBSXls = window.open('/admin/eventmanage/event/pop_event_board_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wBBSXls.focus();
	}
	   
	   
	function jsCmtStyle(sName){  
	    if (eval("document.frmEvt.chk"+sName).checked){ 
	         eval("document.frmEvt."+sName+"MT").value = ""; 
	         eval("document.frmEvt."+sName+"MT").className = "textarea_ro";
	         eval("document.frmEvt."+sName+"MT").disabled  = true;
	        if (sName =="eCmt" || sName == "eIps" ) {
	         eval("document.frmEvt."+sName+"ST").value = "";
	         eval("document.frmEvt."+sName+"ST").className = "textarea_ro";
	         eval("document.frmEvt."+sName+"ST").disabled  = true; 
	        }
	    }else{
	         eval("document.frmEvt."+sName+"MT").className = "textarea"; 
	         eval("document.frmEvt."+sName+"MT").disabled  = false;
	         if (sName =="eCmt" || sName == "eIps" ) {
             eval("document.frmEvt."+sName+"ST").disabled  = false; 
             eval("document.frmEvt."+sName+"ST").className = "textarea";
            }
	    } 
	}

	function jstagchk(v){
		var taglength = document.frmEvt.ietag.length;
		for ( i = 1 ; i<=taglength ; i++ )
		{
			if (v == "1" || v == "2" )
			{
				document.frmEvt.ietagval.readOnly = false;
			}else{
				document.frmEvt.ietagval.readOnly = true;
			}
		}
	}
	
	//슬라이드 체크
	function slidechk(d){
		if (d == "w"){
			if($('input:checkbox[name=slide_w_flag]').is(':checked'))
			{
				var winpop = window.open('/admin/eventmanage/event/v2/template/slide/pop_pcweb_slide.asp?eC=<%=eCode%>','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
				winpop.focus();
			}else{ alert("PCWEB 슬라이드사용 체크 확인");}
		}else{
			if($('input:checkbox[name=slide_m_flag]').is(':checked'))
			{
				var winpop = window.open('/admin/eventmanage/event/v2/template/slide/pop_mobile_slide.asp?eC=<%=eCode%>','winpop','width=1200,height=850,scrollbars=yes,resizable=yes');
				winpop.focus();
			}else{ alert("MOBILE 슬라이드사용 체크 확인");}
		}
	}
	
	//미리보기
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes");
	    }
	}

	function popmoaddimg(){
		var winPopaddimg;
		winPopaddimg = window.open('/admin/eventmanage/event/v2/template/addbanner/pop_mobile_addbanner.asp?eC=<%=eCode%>','pCal','width=1450,height=800,scrollbars=yes,resizable=yes');
		winPopaddimg.focus();
	}

	// 이벤트 상품 최대 할인율 접수
	function fnGetMaxSalevalue() {
		var evtcd = document.frmEvt.eC.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetEvtMaxItemSalePer.asp",
			data: "eC="+evtcd,
			cache: false,
			success: function(message) {
				if(message) {
					document.frmEvt.sSP.value=message;
				} else {
					alert("이벤트의 상품이 없거나 할인중인 상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}
</script>
<form name="frmitemclear" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="IC">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<form name="frmG" method="post" action="eventgroup_process.asp">
  <input type="hidden" name="menupos" value="<%=menupos%>">  
  <input type="hidden" name="eC" value="<%=eCode%>">
  <input type="hidden" name="mode" value="">
  <input type="hidden" name="eCh" value="">
  <input type="hidden" name="eMT" value="">
</form>

<form name="frmEvt" method="post" action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="strparm" value="<%=strparm%>">  
<input type="hidden" name="banMoList" value="<%=ebimgMo2014%>">
<input type="hidden" name="icon" value="<%=eicon%>"> 
<input type="hidden" name="gift" value="<%=gimg%>"> 
<input type="hidden" name="etcitemban" value="<%=eEtcitemimg%>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트코드</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td>
							<%=eCode%>
							<input type="button" value="상품 복사" onclick="jsItemcopylist();" class="button"/>
							<input type='button' value='상품초기화' onclick='jsItemclear();' class='button' /> 
						</td>
		   				<td>
						<%
							'이벤트 종류에 따른 프론트링크 페이지 선택
							Select Case ekind
								Case "7"		'위클리코디
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "13"		'상품 이벤트
								    Response.Write "<td> 미리보기:" 
									Response.Write "<a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & eitemid & "' target='_blank'>PC-Web</a>"
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/category/category_itemPrd.asp?itemid=" & eitemid & "','M');"">Mobile</a>" 
									Response.Write"</td>"
								Case "14"		'소풍가는길
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "16"		'브랜드 할인행사
									Response.Write "<td><a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & ebrand & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>미리보기</a></td>"
								Case "22"		'DAY&(데이앤드)
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "26"		'모바일
									Response.Write "<td><a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case Else		'쇼핑찬스 및 기타
								    Response.Write "<td> 미리보기:" 
									Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'> PC-Web</a>" 
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "','M');"">Mobile</a>" 
									Response.Write "&nbsp;<a href= ""javascript:jsOpen('" & vmobileUrl & "/apps/appCom/wish/web2014/event/eventmain.asp?eventid=" & eCode & "','M');"">App</a>"
								  
									Response.Write"</td>"
							End Select

							'//인스타그램 전용 버튼 (마케팅만일단)
							If session("ssBctId") = "motions" Or session("ssBctId") = "greenteenz" Or session("ssBctId") = "bjh2546" Or session("ssBctId") = "djjung" Or session("ssBctId") = "ksy92630" Or session("ssBctId") = "ppono2" Or session("ssBctId") = "thensi7"  Then
								Response.write "<td><a href=""/admin/etc/only_sys/10x10instagram.asp?eventid="&eCode&""" target='_blank'>10x10instagram</a></td>"
							End If 
						%>
		   				</td>
		   				<td align="right">
		   				<% If sKind = "2" Then %>
		   					<input type="button" value="한마디List" onClick="window.open('/admin/eventmanage/oneline/?eC=<%=eCode%>&esday=<%=esday%>','oneline','width=600,height=500,scrollbars=yes');">
		   					<img src="/images/icon_excel_reply.gif" alt="코멘트 참여자 Excel다운로드" onClick="location.href='/admin/eventmanage/oneline/oneline_excel.asp?eC=<%=eCode%>&esday=<%=esday%>';" style="cursor:pointer" align="absmiddle">
		   				<% Else %>
		   					<img src="/images/icon_excel_reply.gif" alt="코멘트 참여자 Excel다운로드" onClick="popCommentXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   					<img src="/images/icon_excel_bbs.gif" alt="게시판 참여자 Excel다운로드" onClick="popBBSXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   				<% End If %>
		   					<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle" title ="xls 다운로드 회원기반">
							<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드 비회원"  title ="xls 다운로드 비회원" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_guest_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
							<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드 Lite버전"  title ="xls 다운로드 Lite버전" onClick="window.open('/admin/eventmanage/event/pop_event_votelist_lite_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
		   				</td>
		   			</tr>
		   			</table>
		   		</td>
		   	</tr>
		    <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>사용유무</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="radio" name="using" value="Y" <%IF eusing="Y" THEN%>checked<%END IF%>>Yes <input type="radio" name="using" value="N" <%IF eusing="N" THEN%>checked<%END IF%>>No
		   		</td>
		   	</tr> 
			<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>채널</B></td>
		   		<td bgcolor="#FFFFFF">
		   			 <label><input type="checkbox" name="blnWeb" value="1" <%IF isWeb THEN%>checked<%END IF%> onClick="jsChkChannel('P');"> PC-Web</label>
		   			 <label><input type="checkbox" name="blnMobile" value="1" <%IF  isMobile  THEN%>checked<%END IF%> onClick="jsChkChannel('M');"> Mobile</label>
		   			 <label><input type="checkbox" name="blnApp" value="1"  <%IF  isApp  THEN%>checked<%END IF%> onClick="jsChkChannel('A');"> APP</label>
		   		</td>
		   	</tr>
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>종류</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventkind",ekind,False,"onChange=javascript:jsChkSubj(this.value);"%>
		   			<% If ekind = "29" Then %>
		   			<strong> ※ <font color="blue" size="3">개발 및 코딩 작업이 있는 경우</font> <font color="red" size="3">반드시 작업자를 지정해야합니다.</font></strong>
		   			<% End If %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 타입</td>
		   		<td bgcolor="#FFFFFF">  
		   		<input type="checkbox" name="chSale" <%IF esale   THEN%>checked<%END IF%> value="1"  onClick="jsChkSale();">할인
		   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('g');">사은품
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon   THEN%>checked<%END IF%> value="1" onClick="jsChkSale();">쿠폰
		   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten   THEN%>checked<%END IF%> value="1">Only-TenByTen
		   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone  THEN%>checked<%END IF%> value="1">1+1
				<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery   THEN%>checked<%END IF%> value="1">무료배송
				<input type="checkbox" name="chBookingsell" <%IF eBookingsell  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('b');">예약판매
				<input type="checkbox" name="chDiary" <%IF eDiary  THEN%>checked<%END IF%> value="1">DiaryStory
				<input type="checkbox" name="chNew" <%IF eNew  THEN%>checked<%END IF%> value="1">런칭
		   		</td>
			</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기능</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chComm" <%IF ecomment   THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('c');">코멘트
		   		<input type="checkbox" name="chBbs" <%IF ebbs   THEN%>checked<%END IF%> value="1" >게시판
		   		<input type="checkbox" name="chItemps" <%IF eitemps  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('i');">상품후기
		   		<input type="checkbox" name="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
		   		<!--<input type="checkbox" name="chApply" <%IF eapply = 1 THEN%>checked<%END IF%> value="1" >응모-->
		   		</td>
		   	</tr>
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">이벤트 유형</td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventtype",etype,True,""%>
		   			<a href="#" onclick="window.open('popEventType.html','popViewEvtType','width=550,height=480, scrollbars=yes');return false;" style="margin-left:10px;color:#A38;">[이벤트 유형보기]</a>
		   			<span id="lyrEvtConfirm" style="<%=chkIIF(etype="50","","display:none;")%>margin-left:10px;">
		   			<% if isConfirm then %>
		   				<input type="hidden" name="blnCnfm" value="1">
		   				<font color="#AA2244">※ 이벤트 유형이 승인되었습니다.</font>
		   			<% else %>
		   				<label><input type="checkbox" name="blnCnfm" value="1" <%=chkIIF(session("ssAdminLsn")<="3","","readonly")%>> 이벤트 유형 승인</label>
		   			<% end if %>
		   			</span>
		   		</td>
		   	</tr>
		   <!--	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>주체</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventmanager",eman,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>범위</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="hidden" name="eventscope" value="2">
		   			<label><input type="checkbox" name="chkEscope" checked onclick="jsSetPartner()"> 10x10</label>
		   			<label><input type="checkbox" name="chkEscope" onclick="jsSetPartner()"> 제휴몰</label>
		   			<span id="spanP" style="display:none;">
		   			<select name="selP">
		   				<option value="">--제휴몰 전체--</option>
		   				<% sbOptPartner selPartner%>
		   			</select>
		   			</span>
		   		</td>
		   	</tr>-->
		   <tr id="eNameTr_A" style="display:<% if ekind="16"  then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>"> 
		   			<span id="spSale" style="display:<%if not ((ekind="1" or ekind="23") and (esale or ecoupon )) then %>none<%end if%>;<%if esale then%>color:red;<%else%>color:green;<%end if%>">
		   			    <b> 할인율: </b></font><input type="text" name="sSP" value="<%=eSalePer%>" size="10" class="text" />(예:40%~)
		   			    <input type="button" class="button" value="최대값 가져오기" onclick="fnGetMaxSalevalue()" />
		   			</span>
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>영문 이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:<% if ekind<>"16"  then Response.Write "none" %>;">
		   	<%
		   		'// 브랜드할인이면 제목을 할인율로 표시
		   		dim arrEname
				arrEname = Split(ename,"|")
				if Ubound(arrEname)<2 then
					arrEname = ename & "|0|0"
					arrEname = Split(arrEname,"|")
				end if

				If enameEng <> "" then
					Dim arrEnameEng
					arrEnameEng = Split(enameEng,"|")
					if Ubound(arrEnameEng)<2 then
						arrEnameEng = enameEng & "|0|0"
						arrEnameEng = Split(arrEnameEng,"|")
					end If
				End If
		   	%>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명<br>및 할인율</B></td>
		   		<td bgcolor="#FFFFFF">
					이벤트명: <input type="text" name="sEDN" size="60" maxlength="60" value="<%=arrEname(0)%>"><br>
					<% If enameEng <> "" Then %>
		   			영문이벤트명: <input type="text" name="sEDNEng" size="60" maxlength="60" value="<%=arrEnameEng(0)%>"><br>
					<% End If %>
		   			할인율: 최저 <input type="text" name="sSDc" size="4" value="<%=arrEname(1)%>" style="text-align:right;">% ~
		   			최고 <input type="text" name="sMDc" size="4" value="<%=arrEname(2)%>" style="text-align:right;">%<br>
		   			<font color=gray>※브랜드 스트리트에 보여질 할인율입니다. 실제로 상품에는 적용되지 않으니 상품에는 따로 할인을 적용해주세요.
		   		<br>이벤트 링크는 브랜드 스트리트로 연결되니 반드시 상세내용에 브랜드를 선택해주세요.</font> 
		   		</td>
		   	</tr>
		   	<tr>
		   		<td rowspan="2" align="center" bgcolor="<%= adminColor("tabletop") %>"><B>기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<%IF estate = 9 THEN%>
		   			시작일 : <%=esday%><input type="hidden" name="sSD" size="10" value="<%=esday%>">
		   			~ 종료일 : <%=eeday%> <input type="hidden" name="sED" value="<%=eeday%>" size="10" >
		   		<%ELSE%>
		   			시작일 : <input type="text" id="termSdt" name="sSD" size="10" value="<%=esday%>" />
							<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
		   			~ 종료일 : <input type="text" id="termEdt" name="sED" size="10" value="<%=eeday%>" />
							<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "termSdt", trigger    : "ChkStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(eeday<>"",", max: " & replace(eeday,"-",""),"")%>
						});
						var CAL_End = new Calendar({
							inputField : "termEdt", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(eeday<>"",", min: " & replace(esday,"-",""),"")%>
						});
					</script>
		   		<%END IF%>
				&nbsp;&nbsp;<input type="checkbox" name="endlessview"  value="Y"  <%IF endlessView="Y" THEN%>checked<%END IF%>> <a title="상시노출 설정시 기간이 지난 이벤트도 이벤트 종료 안내 레이어 안뜨도록 설정">상시노출</a>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			당첨 발표일 : <input type="text" id="priceDt" name="sPD" size="10" value="<%=epday%>" />
					<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkPrc_trigger" onclick="return false;" />
					(당첨자가 있는 경우에만 등록)
					<script type="text/javascript">
						var CAL_Prcdt = new Calendar({
							inputField : "priceDt", trigger    : "ChkPrc_trigger",
							onSelect: function() {
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
		   		</td>
		   	</tr>
		  	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>상태</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%   sbGetOptStatusCodeSort "eventstate",estate,false,"" 
		   				''if ekind="22" then
		   				''	'//데이앤드는 디자인파트만 사용해서 기존대로
		   				''	sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''else
		   				''	sbGetOptStatusCodeAuth "eventstate",estate,"M",""
		   				''end if
		   			%>
		   			<input type="hidden" name="eOD" value="<%=dopendate%>">
		   			<input type="hidden" name="eCD" value="<%=dclosedate%>">
		   			<input type="hidden" name="eIRD" value="<%=dImgregdate%>">
		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  오픈처리일 : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  종료처리일 : <%=dclosedate%>  </span><%END IF%>
		   			<%IF not isnull(dImgregdate) THEN%>/ <span style="padding-left:10px;">  이미지등록요청일 : <%=dImgregdate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>중요도</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
		</table>  
		<div id="divDE" style="display:none;"> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>"><b>정렬번호</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="0" size="6" maxlength="5" style="text-align:right;" />
		   			(※숫자가 클수록 우선표시 됩니다. / Day&:회차)
		   		</td>
		   	</tr> 
		</table>
		</div>
	</td>
</tr>
<tr>
	<td > 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">관리 카테고리</td>
		   		<td bgcolor="#FFFFFF">
		   		<%'DrawSelectBoxCategoryOnlyLarge "selCategory", ecategory,"" %>
		   		<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
		   		</td>
		   	</tr>
		   		<tr>
		   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">전시 카테고리</td>
		   		<td bgcolor="#FFFFFF">
		   		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		   		</td>
		   	</tr>
		   <tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
		   		<td bgcolor="#FFFFFF">
		   			<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
		   		</td>
		   	</tr>
		    
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">상품정렬방법</td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
		   		</td>
		   	</tr>
		   	<tr>    
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">퍼블리싱</td>
		   		<td bgcolor="#FFFFFF"><input type="checkbox" name="chkReqP" value="1" <%if blnReqPublish then%>checked<%end if%>>  퍼블리싱 요청 <img src="/images/admin_help.gif" style="cursor:hand;" onClick="jsPubHelp();"></td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당자</td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" class="a" cellpadding="1">
	   				<tr>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;" width="96">기획자</td>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;">
	   						<input type="hidden" name="sMdId" value="<%=emdid%>">
	   						<input type="name" name="sMdNm" value="<%=eMDnm%>"class="text_ro" readonly size="10">
	   						<input type="button" class="button" value="선택" onClick="jsGetID('Md','<%=idepartmentid%>','<%=emdid%>');">
	   						<input type="button" value="&times"  class="button" onClick="jsDelID('Md');" title="담당자 지우기" />
	   					</td>
	   				</tr>
	   				<tr>
	   					<td>디자이너(PC)</td>
	   					<td>
			   			    <%sbGetDesignerid "sDgId",edgid,"onchange=""jsChangeFrm(this.value,'DG1');"""%>
			   			    <%sbGetOptEventCodeValue "designerstatus",edgstat1,True,""%>
	   					</td>
	   				</tr>
	   				<tr>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;">디자이너(Mobile)</td>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;">
			   			    <%sbGetDesignerid "sDgId2",edgid2,"onchange=""jsChangeFrm(this.value,'DG2');"""%>
			   			    <%sbGetOptEventCodeValue "designerstatus",edgstat2,True,""%>
	   					</td>
	   				</tr>
	   				<tr>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;">퍼블리셔</td>
	   					<td style="border-bottom:1px dashed <%= adminColor("tablebg") %>;">
			   			    <input type="hidden" name="sPsId" value="<%=epsid%>">
			   			    <input type="name" name="sPsNm" value="<%=epsnm%>"class="text_ro" readonly size="10">
			   			    <input type="button" class="button" value="선택"  onClick="jsGetID('Ps','157','<%=epsid%>');">
			   			    <input type="button" value="&times"  class="button" onClick="jsDelID('Ps');" title="담당자 지우기" />
	   					</td>
	   				</tr>
	   				<tr>
	   					<td>개발자</td>
	   					<td>
			   			    <input type="hidden" name="sDpId" value="<%=edpid%>">
			   			    <input type="name" name="sDpNm" value="<%=edpnm%>" class="text_ro" readonly size="10">
			   			    <input type="button" class="button" value="선택" onClick="jsGetID('Dp','130','<%=edpid%>');">
			   			    <input type="button" value="&times"  class="button" onClick="jsDelID('Dp');" title="담당자 지우기" />
	   					</td>
	   				</tr>
	   				<tr>
	   					<td>개발검수자</td>
	   					<td>
			   			    <input type="hidden" name="sCCId" value="">
			   			    <input type="name" name="sCCNm" value="" class="text_ro" readonly size="10">
			   			    <input type="button" class="button" value="선택" onClick="jsGetID('CC','130','<%=edpid%>');">
			   			    <input type="button" value="&times"  class="button" onClick="jsDelID('CC');" title="담당자 지우기" />
	   					</td>
	   				</tr>
		   			</table>
		   		</td>
		   	</tr>  
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">디자이너 작업구분</td>
		   		<td bgcolor="#FFFFFF"><input type="text" name="sWorkTag" size="20" maxlength="16" class="text" value="<%= sWorkTag %>">
		   		    <input type="checkbox" name="chkMB"  onClick="jsChkMBReq();" <%if left(sWorkTag,4) ="★★" then%>checked<%end if%>> 모바일배너 요청시 체크   
		   		    </td>
		   	</tr> 
		 <!--삭제  2015.02.05
		 	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
		   		<td bgcolor="#FFFFFF">
		   			(200자 이내)		   			<Br>
		   			<textarea name="eCT" rows="2" style="width:100%;"></textarea>
		   		</td>
		   	</tr>-->
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
		   		<td bgcolor="#FFFFFF">
		   			(250자 이내)		   			<Br>
		   			<textarea name="eTag" rows="2" style="width:100%;"><%=etag%></textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">연관 이벤트코드</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="text" name="eLC" size="6" maxlength="10" value="<%=elinkcode%>"> 
		   		</td>
		   	</tr> 
		</table> 
	</td>
</tr> 
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="0">
			<tr>
				<td width="50%" valign="top">
				    <div id="divPC1" style="display:<%if not isweb then%>none<%end if%>;">
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tR>
							<td bgcolor="#FAECC5" align="center" colspan="2"><b>PC-WEB</b></td>
						</tr>
						<tr>
							<td align="center" bgcolor="#FAECC5">작업전달사항</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
							</td>
						</tr>
						<tr> 
							<td align="center" bgcolor="#FAECC5"><b>서브카피</B></td>
					   		<td bgcolor="#FFFFFF">  
					   			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					   			<tr> 
					   				<td width="48%" style="padding-right:3px;"><textarea name="subcopyK" style="width:100%; height:80px;" onclick="if(this.value=='한글')this.value='';" onblur="if(this.value=='')this.value='한글';" value="<%=subcopyK%>"><%=chkiif(subcopyK="","한글",subcopyK)%></textarea></td>
					   				<td width="48%"><textarea name="subcopyE" style="width:100%; height:80px;" onclick="if(this.value=='영문')this.value='';" onblur="if(this.value=='')this.value='영문';" value="<%=subcopyE%>"><%=chkiif(subcopyE="","영문",subcopyE)%></textarea></td>
					   			</tr> 
					   			</table>
					   		</td>
						</tr>
					 
						<tr>
					   		<td width="100" align="center"  bgcolor="#FAECC5">화면구성</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="radio" name="chkFull" value="0" <%IF not blnFull and not blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> 기본형&nbsp;&nbsp;
					   			<input type="radio" name="chkFull" value="1" <%IF  blnFull and not blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> 풀단&nbsp;&nbsp;
					   			<input type="radio" name="chkWide" value="1" <%IF blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkFull[0].checked=false;chkFull[1].checked=false;"> 와이드 
					   		</td>
						</tr> 
						<tr>
					   		<td align="center" bgcolor="#FAECC5">상품정보</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkIteminfo"  value="1"  <%IF blnIteminfo THEN%>checked<%END IF%>>  사용함
					   		</td>
					  	</tr>
						<tr>
					   		<td align="center"   bgcolor="#FAECC5">상품 가격정보<br/><font color="#BB8866">[쿠폰및 할인가<br/>노출여부]</font></td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkItemprice"  value="1"  <%IF blnitemprice THEN%>checked<%END IF%>> 노출안함
					   		</td>
					  	</tr>
						<tr>
					   		<td align="center"  bgcolor="#FAECC5">이벤트 기간<br/>노출여부</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="dateview"  value="1"  <%IF eDateView THEN%>checked<%END IF%>>  노출안함
					   		</td>
					  	</tr>
					  	<tr id="eNameTr_BL" style="display:<%if ekind<>16 then%>none<%end if%>;"> 
					   		<td align="center"  bgcolor="#FAECC5">브랜드이벤트 링크</td>
					   		<td bgcolor="#FFFFFF"> 
					   		 <input type="hidden" name="elType" value="<%=chkiif(eKind=16,"I","E")%>"> 
					   		 <input type="text" id="elUrl" name="elUrl" size="60" maxlength="128" value="<%= elkurl %>" class="text" > 
					   		</td>
					   	</tr> 
					 	<tr>
					   		<td align="center"  bgcolor="#FAECC5">대표상품정보<br/>및<br/>배너</td>
					   		<td bgcolor="#FFFFFF">
					   			<font color="red"><b>※ 카테고리메인과 엔조이이벤트 리스트에 나오는 이미지.<br/>대표상품이미지를 안넣으면 대표상품코드를 반드시 넣어야함.<br/>대표상품이미지가 없으면 대표상품코드의 기본 이미지를 사용하게 됨.</b></font><br/>
								대표상품코드 : <input type="text" name="etcitemid" value="<%=eEtcitemid%>"><br/>
								대표상품이미지(420x420) : <input type="button" name="etcitem" value="상품대표배너" onClick="jsSetImg('<%=eFolder%>','<%=eEtcitemimg%>','etcitemban','etciitem')" class="button">
					   			<div id="etciitem" style="padding: 5 5 5 5">
					   				<%IF eEtcitemimg <> "" THEN %>
					   				<img  src="<%=eEtcitemimg%>" border="0">
					   				<a href="javascript:jsDelImg('etcitemban','etciitem');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
					</table>
				    </div>
				</td>
				<td  valign="top">
				    <div id="divMA1" style="display:<%if not (isMobile or isApp) then%>none<%end if%>;">
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tR>
							<td bgcolor="#e3f1fb" align="center"  colspan="2"><b>Mobile / App</b></td>
						</tr>
						<tr>
							<td align="center" bgcolor="#e3f1fb">작업전달사항</td>
							<td bgcolor="#FFFFFF"> 
								<textarea name="tFwdMo" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd_mo%></textarea>
							</td>
						</tr>
						<tr>
							<td align="center" bgcolor="#e3f1fb"><b>서브카피</B></td>
					   		<td bgcolor="#FFFFFF"> <input type="text" name="subsEN" size="70" maxlength="70" value="<%=enamesub%>"  OnKeyUp="jsAddByte(this);"> 
					   		    <input type="text" name="subSize" size="3" value="" class="text_ro" style="text-align:right" readonly>Byte
					   		     <p style="color:#602030;font-size:11px;"> [ 최대 70byte까지 등록가능합니다. ]</p>
					   		    <script type="text/javascript">
					   		        jsAddByte(frmEvt.subsEN);
					   		     </script>
					   		 </td>
					   	</tr> 
						<tr>
							<td align="center"  bgcolor="#e3f1fb">상품리스트 스타일</td>
							<td bgcolor="#FFFFFF">
								<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>> 격자형&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>> 리스트형&nbsp;&nbsp;
								<input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>> BIG형
							</td>
						</tr>	 
						<tr>
							<td align="center"  bgcolor="#e3f1fb">텍스트 타이틀</td>  
							<td bgcolor="#FFFFFF">
							    <a href="javascript:jsOpen('<%=vmobileUrl%>/event/eventmain.asp?eventid=<%=eCode%>','M');">  + 미리보기 </a>  
								<div id="dvTxT">
								<table border="0" cellpadding="3" cellspacing="1" class="a" width="100%">  
								<tr><%dim chkeCmt, chkeIps, chkeGf, chkeBS%>
									<td colspan="2">
										<% if rdCmt="" THEN rdCmt=1%>
										<div id="dvCmt"  style="display:<%IF not ecomment THEN %>none<%end if%>;"> 
										   <table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
											        <input type="radio" name="rdCmt" value="1" <%if rdCmt="1" THEN%>checked<%END IF%>>코멘트
											        <input type="radio" name="rdCmt" value="2" <%if rdCmt="2" THEN%>checked<%END IF%>>테스터 코멘트
											        <input type="checkbox" name="chkeCmt" value="0" <%if chkeCmt="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eCmt');">사용안함(이미지로 등록)
                                                </th>
        										<tr>
        											<td bgcolor="#e3f1fb">주제</td><td bgcolor="#FFFFFF"><textarea type="text" name="eCmtMT" cols="50" rows="3" <%if chkeCmt="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=eCmtMT%></textarea> <span style="color:#602030;font-size:11px;">[200자 이내]</span></td>
        										</tr>
        										<tR >
        											<td bgcolor="#e3f1fb">당첨자수/<br/>사은품</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eCmtST"  <%if chkeCmt="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=db2Html(eCmtST)%></textarea></td>
        										</tr>
        							        </table> 
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="dvIps" style="display:<%IF not eitemps THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
												    상품후기 
												    <input type="checkbox" name="chkeIps" value="0" <%if chkeIps="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eIps');">사용안함(이미지로 등록)
												  </th> 
												<tr>
													<td bgcolor="#e3f1fb">주제</td><td bgcolor="#FFFFFF"><textarea  name="eIpsMT" cols="50" rows="3"  <%if chkeIps="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=eIpsMT%></textarea> <span style="color:#602030;font-size:11px;">[200자 이내]</span></td>
												</tr>
												<tR>
													<td bgcolor="#e3f1fb">당첨자수/<br/>사은품</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eIpsST"  <%if chkeIps="0" THEN%>class="textarea_ro" disabled<%else%>class="textarea"<%end if%>><%=db2Html(eIpsST)%></textarea></td>
												</tr>
											 </table>
										</div>
									</td>
								</tr>
								 <tr>
									<td colspan="2">
										<div id="dvGf"  style="display:<%IF not egift THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%"  bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">
												    사은품 
												    <input type="checkbox" name="chkeGf" value="0" <%if chkeGf="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eGf');">사용안함(이미지로 등록)
												</th> 
												<tr>
													<td bgcolor="#FFFFFF"><textarea  name="eGfMT" cols="50"  rows="3" <%if chkeGf="0" then%>class="textarea_ro" disabled<%else%> class="textarea"<%end if%>><%=eGfMT%></textarea> <span style="color:#602030;font-size:11px;">[200자 이내]</span></td>
												</tr> 
											 </table>
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="dvBS" style="display:<%IF not eBookingsell THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left"  bgcolor="#e3f1fb"> 
												    예약판매 
												    <input type="checkbox" name="chkeBS" value="0" <%if chkeBS="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eBS');">사용안함(이미지로 등록)
												</th> 
												<tr>
													<td bgcolor="#FFFFFF"><textarea name="eBSMT" cols="50"  rows="3"  <%if chkeBs="0" then%>class="textarea_ro" disabled<%else%> class="textarea"<%end if%>><%=eBSMT%></textarea> <span style="color:#602030;font-size:11px;">[200자 이내]</span></td>
												</tr> 
											 </table>
										</div>
									</td>
								</tr> 
								</table>
								</div>
							</td>
						</tr>
						<!-- 상품 이벤트 -->
						<tr id="itemevt" style="display:<%=chkiif(ekind="13","","none")%>;">
							<td bgcolor="#e3f1fb" align="center" colspan="2">
								<div>
								<table border="0" cellpadding="3" cellspacing="1" class="a" width="100%">
								<tr>
									<td bgcolor="#e3f1fb" align="center"><strong>상품이벤트</strong></td>
								</tr>
								<tr>
									<td>
										<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
										<tr>
											<td align="center" bgcolor="#e3f1fb"><b>태그</B></td>
											<td bgcolor="#FFFFFF">
												<input type="radio" name="ietag" value="7" <%=chkiif(tmpietag="7","checked","")%> onclick="jstagchk(this.value);"/> 할인 <input type="radio" name="ietag" value="2" <%=chkiif(tmpietag="2","checked","")%> onclick="jstagchk(this.value);"/> 쿠폰 <input type="text" size="5" name="ietagval" value="<%=tmpietagval%>" <%=chkiif(tmpietag="7" Or tmpietag = "2" ,"","readOnly")%> class="text_ro"/> <input type="radio" name="ietag" value="1" <%=chkiif(tmpietag="1","checked","")%> onclick="jstagchk(this.value);"/> GiFT <input type="radio" name="ietag" value="4" <%=chkiif(tmpietag="4","checked","")%> onclick="jstagchk(this.value);"/> 무료배송 <input type="radio" name="ietag" value="5" <%=chkiif(tmpietag="5","checked","")%> onclick="jstagchk(this.value);"/> 1:1 <input type="radio" name="ietag" value="6" <%=chkiif(tmpietag="6","checked","")%> onclick="jstagchk(this.value);"/> 1+1 <input type="radio" name="ietag" value="3" <%=chkiif(tmpietag="3","checked","")%> onclick="jstagchk(this.value);"/> 예약배송
											</td>
										</tr>
										<tr>
											<td align="center" bgcolor="#e3f1fb"><b>프로모션 내용</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="mcopy" maxlength="50" value="<%=tmpmcopy%>"/><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: 오늘 단 하루, UDH-02 전기렌지 증정! )</div></td>
										</tr>
										<tr>
											<td align="center" bgcolor="#e3f1fb"><b>서브 내용</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="scopy" maxlength="50" value="<%=tmpscopy%>"/><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: 선착순 100명/ 소진 시 조기종료 )</div></td>
										</tr>
										</table>
									</td>
								</tr>
								</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr> 
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					    <tr>
            				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">와이드배너<br/>(모바일리스트)</td>
            				<td bgcolor="#FFFFFF">
            				<input type="button" name="btnMoBan2014" value="와이드 배너" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo2014%>','banMoList','spanbanMoList')" class="button">
            					 <div id="spanbanMoList" style="padding: 5 5 5 5">
            						<%IF ebimgMo2014 <> "" THEN %>
            						<img  src="<%=ebimgMo2014%>" border="0">
            						<a href="javascript:jsDelImg('banMoList','spanbanMoList');"><img src="/images/icon_delete2.gif" border="0"></a>
            						<%END IF%>
            					</div>
            					<p style="color:#602030;font-size:11px;">[ 권장 이미지 : JPEG, 60%, 750px × 406px ]</p>
            				</td>
            			</tr>  
            			<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">사은품 이미지 </td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" name="btnicon" value="사은품이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=gimg%>','gift','spangift')" class="button">
					   			<div id="spangift" style="padding: 5 5 5 5">
					   				<%IF gimg <> "" THEN %>
					   				<a href="javascript:jsImgView('<%=gimg%>')"><img  src="<%=gimg%>" width="400" border="0"></a>
					   				<a href="javascript:jsDelImg('gift','spangift');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
					</table>
				</td> 
				<td style="vertical-align:top">
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					    <tr>
            				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">모바일 연결배너<br/>(최대 5개)</td>
            				<td bgcolor="#FFFFFF">
	            				<input type="button" value="모바일 추가 배너(<%=chkiif(evt_m_addimg_cnt<>"",evt_m_addimg_cnt,"0")%>)" onClick="popmoaddimg();" class="button">
            				</td>
            			</tr>  
					</table>
				</td>
			</tr>
			<tr>
				<td valign="top">
				    <div id="divPC2" style="display:<%if not isWeb then%>none<%end if%>;">
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tR>
							<td bgcolor="#FAECC5" align="center" colspan="2"><b>PC-WEB</b></td>
						</tr>
						<tr>
							<td width="100" height="50" align="center"  bgcolor="#FAECC5">화면템플릿</td>
							<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview",etemp,false,"onchange=""jsChangeFrm(this.value,'P');"""%>
								<span id="w_slide" style="display:<% If etemp <> 1 And etemp <> 3 And etemp <> 7 then %>none<% End If %>"><input type="checkbox" name="slide_w_flag" id="slide_w_flag" value="Y" <%=chkiif(slide_w_flag="Y","checked","")%>><label for="slide_w_flag">슬라이드사용</label>&nbsp;<input type="button" value="등록/수정" onclick="slidechk('w');return false;"/>&nbsp;</span>
								<span id="divGM" style="display:<%if etemp <> 3 and etemp <> 7 then%>none<%end if%>;">
									<input type="button" value="그룹관리" onClick="jsAddGroup('<%=eCode%>','','I','P');" class="button" style="color:blue;width:80" >
									<span  style="float:right;"><input type="checkbox" value="1" name="sgroup_w"  <%=chkiif(sgroup_w=true," checked","")%>> 최상위 랜덤노출</span>
									  <%IF not blngroup THEN%>  
									  <div style="padding:5 0 5 0px;display:;" id="divForm" >
									   <input type="button" value="Tab1+기차5 그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','1','P');" class="button">, 
									   <input type="button" value="Tab2+기차5 그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','2','P');" class="button">,
									   <input type="button" value="Tab3+기차5  그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','3','P');" class="button">  
									  </div> 
									  <%END IF%> 
								</span>
							</td>
						</tr> 
						<tr>
							<td bgcolor="#FAECC5" width="100" align="Center">이미지<br>&<br>HTML</td>
							<td bgcolor="#ffffff">
								<!-- 1.메인 탑-->
					   			<div id="divFrm1" style="display:<%if etemp <> 1 then%>none<%end if%>;">
					   				<input type="hidden" name="main" value="<%=emimg%>">
						   			<input type="button" name="btnMain" value="메인TOP이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=emimg%>','main','spanmain')" class="button">
						   			<div id="spanmain" style="padding: 5 5 5 5">
						   				<%IF emimg <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg%>')"><img  src="<%=emimg%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main','spanmain');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
								   	<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('1');">PC-WEB 예시</span> 
									<div id="notice1" style="display:block">
									&lt;map name="Mainmap"&gt;<br>
									<font color="blue">상품페이지 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoProduct('<font color="blue">상품번호</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">이벤트페이지로 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain('<font color="blue">이벤트코드</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">이벤트 그룹 페이지로 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventGroupMain('<font color="blue">이벤트코드</font>','<font color="blue">그룹코드</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">이벤트 사은품 팝업 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:popShowGiftImg('<font color="blue">이벤트코드</font>');" onfocus="this.blur();"&gt;<br>
									<font color="blue">브랜드페이지 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:GoToBrandShop('<font color="blue">브랜드아이디</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;<br>
									<font color="blue">레드리본 메인 링크시</font><br>
									&lt;area shape="rect" coords="0,0,0,0" href="javascript:TnGotoEventMain_New('<font color="blue">이벤트코드</font>');" onfocus="this.blur();"&gt;<br>
									&lt;/map&gt;
									</map> <br>
									<font color="blue">기차형 타이틀 이미지로 링크시</font><br>
									&lt;area shape="circle" coords="186,250,144" href="#event_namelink1" onfocus="this.blur();"&gt;<br>
									href="#event_namelink2" href="#event_namelink3" 등등 href에 숫자를 바꿔줌. &lt;area끼리는 칸을 내리지말고 꼭 붙임.<br>
									</div> 
									<br>
									<b>이미지 경로 http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> 로 변경되었습니다.<br>
									<textarea name="tHtml" rows="20" style="width:100%;font-size:11px;"><%=ehtml%></textarea>
								</div>
								<!-- 3.그룹형-->
								<div id="divFrm3" style="display:<%if not ( etemp = 3 or etemp = 7) then%>none<%end if%>;"> 
									<%IF isArray(arrGroup) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>그룹코드</td>					
										<td>상위그룹</td>
										<td>그룹명</td>
										<td>정렬순서</td>					
										<td>이미지</td>
										<td>전시여부</td>
										<td>관리</td>
									</tr>
									<%FOR intg = 0 To UBound(arrGroup,2)%>				   						
									<tr <%if not arrGroup(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup(0,intg)%></td>						
										<td  align="center"><%IF isnull(arrGroup(7,intg))THEN%>최상위<%ELSE%>[<%=arrGroup(5,intg)%>]<%=db2html(arrGroup(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup(1,intg))%></td>	
										<td  align="center"><%=arrGroup(2,intg)%></td>									   									
										<td  align="center">  
											<a href="javascript:jsImgView('<%=arrGroup(3,intg)%>');"><img src="<%=arrGroup(3,intg)%>" width="50" border="0"></a> 
										</td>	
										<td  align="center"><%if arrGroup(8,intg) then%>Y<%else%>N<%end if%></td>						   									
										<td  align="center">
											<input type="button" name="btnU" value="수정" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup(0,intg)%>','P')" class="button">
											<!--<input type="button" name="btnD" value="삭제" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="상품등록" onclick="popRegItem('<%=eCode%>','<%=arrGroup(0,intg)%>','P')"  class="button">
											<% IF arrGroup(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup(0,intg) &"' target='_blank'>미리보기</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%NEXT%>
									</table>
								<%END IF%>	 
								</div>
								<!-- /3.그룹형-->
								<!-- 5.수작업-->
								<div id="divFrm5" style="display:<%if  not ( etemp = 5 or etemp = 6) then%>none<%end if%>;">
									<table border="0" cellpadding="1" cellspacing="3" class="a">
										<tr>
											<td>
											    <!-- <input type="button" value="이미지관리"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/','/event/');" class="input_b"> -->
											    <input type="button" value="이미지관리(신)"  onclick="jsManageEventImageNew('<%=eCode%>')" class="input_b">
											    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											    <input type="button" value="이미지관리(구)"  onclick="jsManageEventImage('<%=eCode%>')" class="input_b">
											    <p>
											    <b>이미지관리(신)</b> : 새로이 변경된 이미지 관리<br>
											    <b>이미지관리(구)</b> : 기존에 저장된 이미지 리스트(이미지추가없음. 새로운 이미지 추가는 이미지관리(신)에서만.)<br>
											    ※ 이벤트 이미지 시스템 관리 차원에서 eventIMG 라는 새로운 폴더에 이벤트시작年 폴더를 추가하여 그 안에 이벤트코드별 폴더를 생성하게 됩니다.<br>
											    추후 몇달이 지난 뒤에 이미지관리(구)는 사용을 하지않고 이미지관리(신)만 사용하게 됩니다.<br>
											    그때까지는 불편사항이 있으시더라도 시스템관리 차원에 의한 조치이므로 양해바랍니다.
											</td>
										</tr>
										<tr>
										    <td><b>이미지 경로 http://<font color="RED">webimage.</font>10x10.co.kr/eventIMG/이벤트시작年/XXX/</b> 로 변경되었습니다.</td>
										</tr>
										<tr>
											<td><textarea name="tHtml5" rows="25" style="width:100%;font-size:11px;"><%=ehtml5%></textarea></td>
										</tr> 
									</table>
								</div>
								<!-- /5.수작업-->
							</td>
						</tr>
						<tr>
                		    <td bgcolor="#FAECC5" width="100" align="Center">Exec File
                		        <br/><span style="color:#602030;font-size:11px;">[ 개발 실행파일]</span>
                		        </td>
                			<td bgcolor="#ffffff"  >
                		         <input type="radio" name="rdoEF" value="0" <%if not blnExec then%>checked<%end if%>>비실행 
						        <input type="radio" name="rdoEF" value="1" <%if blnExec then%>checked<%end if%>>실행 <input type="text" name="sEFP" size="60" class="text" value="<%=eExecFile%>"> 
                		    </td>
                		</tr>
					</table>	
				    </div>
				</td>
				<td valign="top">
				    <div id="divMA2" style="display:<%if not (isMobile or isApp) then%>none<%end if%>;">
					<table width="100%" border="0" align="left" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tR>
							<td bgcolor="#e3f1fb" align="center"  colspan="2"><b>Mobile / App</b></td>
						</tr>
						<tr>
							<td width="100" height="50" align="center"  bgcolor="#e3f1fb">화면템플릿</td>
							<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview_mo",etemp_mo,false,"onchange=""jsChangeFrm(this.value,'M');"""%>
								<span id="m_slide" style="display:<% If etemp <> 1 And etemp <> 3 then %>none<% End If %>"><input type="checkbox" name="slide_m_flag" id="slide_m_flag" value="Y" <%=chkiif(slide_m_flag="Y","checked","")%>><label for="slide_m_flag">슬라이드사용</label>&nbsp;<input type="button" value="등록/수정" onclick="slidechk('m');return false;"/>&nbsp;</span>
								<span id="divGM_mo" style="display:<%if etemp_mo <> 3 and etemp_mo <> 7 then%>none<%end if%>;">
									<input type="button" value="그룹관리" onClick="jsAddGroup('<%=eCode%>','','I','M');" class="button" style="color:blue;width:80" >
									<span style="float:right;"><input type="checkbox" value="1" name="sgroup_M" <%=chkiif(sgroup_m=true," checked","")%>> 최상위 랜덤노출</span>
									<%IF not blngroup_mo THEN%>
									<div style="padding:3 0 3 0px;display:;" id="divForm_mo">
									    <input type="button" value="Tab1+기차5 그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','1','M');" class="button">, 
									    <input type="button" value="Tab2+기차5 그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','2','M');" class="button">,
									    <input type="button" value="Tab3+기차5  그룹생성" onClick="jsAddProcGroup('<%=eCode%>','F','3','M');" class="button">   
									</div>    
									<%END IF%> 
								</span> 
							</td>
						</tr> 
						<tr>
							<td bgcolor="#e3f1fb" width="100" align="Center">이미지<br>&<br>HTML</td>
							<td bgcolor="#ffffff" valign="top">
									<!-- 1.메인 탑-->
								<div id="divMFrm1" style="display:<%if etemp_mo <> 1 then%>none<%end if%>;">
									<input type="hidden" name="main_mo" value="<%=emimg_mo%>">
						   			<input type="button" name="btnMain_mo" value="메인TOP이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=emimg_mo%>','main_mo','spanmain_mo')" class="button">
						   			<div id="spanmain_mo" style="padding: 5 5 5 5">
						   				<%IF emimg_mo <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg_mo%>')"><img  src="<%=emimg_mo%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main_mo','spanmain_mo');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
								  	<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('2');">Mobile-WEB 예시</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('3');">APP 예시</span>
									<div id="notice2" style="display:">
										<font color="blue">상품페이지 링크시</font><br>
										&lt;a href="/category/category_itemprd.asp?itemid=<span style="color:red">상품코드</span>"&gt; 상품페이지 링크 &lt;/a&gt;<br>
										<font color="blue">이벤트페이지로 링크시</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">이벤트코드</span>"&gt; 이벤트페이지 링크 &lt;/a&gt;<br>
										<font color="blue">이벤트 그룹 페이지로 링크시</font><br>
										&lt;a href="/event/eventmain.asp?eventid=<span style="color:red">이벤트코드</span>&eGc=<span style="color:red">그룹코드</span>"&gt; 이벤트 그룹 페이지 링크 &lt;/a&gt;<br>
										<font color="blue">브랜드페이지 링크시</font><br>
										&lt;a href="/street/street_brand.asp?makerid=<span style="color:red">브랜드코드</span>"&gt; 브랜드페이지 링크 &lt;/a&gt;<br>
									</div>
									<div id="notice3" style="display:none">
										※패이지내에서 이동할때※<br/>
										<font color="blue">상품페이지 링크시</font><br>
										&lt;a href="/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<span style="color:red">상품코드</span>"&gt; 상품페이지 링크 &lt;/a&gt;<br>
										<font color="blue">이벤트페이지로 링크시</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">이벤트코드</span>"&gt; 이벤트페이지 링크 &lt;/a&gt;<br>
										<font color="blue">이벤트 그룹 페이지로 링크시</font><br>
										&lt;a href="/apps/appCom/wish/web2014/event/eventmain.asp?eventid=<span style="color:red">이벤트코드</span>&eGc=<span style="color:red">그룹코드</span>"&gt; 이벤트 그룹 페이지 링크 &lt;/a&gt;<br>
										<br>
										※팝업으로 페이지 열때※<br/>
										※수작업 iframe추가할땐 일때 <span style="color:blue">parent.</span> 함수명으로 추가※<br/>
										ex) &lt;a href="#" onclick="<span style="color:blue">parent.</span>fnAPPpopupProduct('<span style="color:red">상품코드</span>'); return false;"&gt; 상품페이지 링크 &lt;/a&gt;<br>
										<font color="blue">상품페이지 링크시</font><br>
										&lt;a href="#" onclick="fnAPPpopupProduct('<span style="color:red">상품코드</span>'); return false;"&gt; 상품페이지 링크 &lt;/a&gt;<br>
										<font color="blue">이벤트페이지로 링크시</font><br>
										&lt;a href="#" onclick="fnAPPpopupEvent('<span style="color:red">이벤트코드</span>'); return false;"&gt; 이벤트페이지 링크 &lt;/a&gt;<br>
										<font color="blue">브랜드페이지 링크시</font><br>
										&lt;a href="#" onclick="fnAPPpopupBrand('<span style="color:red">브랜드명</span>'); return false;"&gt; 브랜드 링크 &lt;/a&gt;<br>
										<font color="blue">카테고리 링크시</font><br>
										&lt;a href="#" onclick="fnAPPpopupCategory('<span style="color:red">카테고리번호</span>'); return false;"&gt; 카테고리 링크 &lt;/a&gt;<br>
									</div>
									<br>
									<b>이미지 경로 http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> 로 변경되었습니다.<br>
									<textarea name="tHtml_mo" rows="20" style="width:100%;font-size:11px;"><%=ehtml_mo%></textarea>
								</div>
								<!-- 3.그룹형-->
								<div id="divMFrm3" style="display:<%if not ( etemp_mo = 3 or etemp_mo = 7) then%>none<%end if%>;">
									<%IF isArray(arrGroup_mo) THEN %>
									<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
									<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
										<td>그룹코드</td>					
										<td>상위그룹</td>
										<td>그룹명</td>
										<td>정렬순서</td>					
										<td>이미지</td>
										<td>전시여부</td>
										<td>관리</td>
									</tr>
									<% dim sumi,i
									FOR intg = 0 To UBound(arrGroup_mo,2)
									 sumi= 0
									%>				   						
									<tr <%if not arrGroup_mo(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
										<td  ><%IF arrGroup_mo(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup_mo(0,intg)%> 
										     <% if intg < UBound(arrGroup_mo,2)  then 
                                				 for i = 1 to (UBound(arrGroup_mo,2)-intg)%> 
                                				<%if arrGroup_mo(9,intg) = arrGroup_mo(9,intg+i) then
                                					sumi = sumi + 1 
                                				 %>
                                				 + <%=arrGroup_mo(0,intg+i)%>
                                				<%else 
                                					exit for
                                				 end if 
                                				next
                                			end if    %>
										 </td>						
										<td  align="center"><%IF isnull(arrGroup_mo(7,intg))THEN%>최상위<%ELSE%>[<%=arrGroup_mo(5,intg)%>]<%=db2html(arrGroup_mo(7,intg))%><%END IF%></td>	
										<td  align="center"><%=db2html(arrGroup_mo(1,intg))%></td>	
										<td  align="center"><%=arrGroup_mo(2,intg)%></td>									   									
										<td  align="center">  
											<a href="javascript:jsImgView('<%=arrGroup_mo(3,intg)%>');"><img src="<%=arrGroup_mo(3,intg)%>" width="50" border="0"></a> 
										</td>			
										<td  align="center"><%if arrGroup_mo(8,intg) then%>Y<%else%>N<%end if%></td>				   									
										<td  align="center">
											<input type="button" name="btnU" value="수정" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')" class="button">
											<!--<input type="button" name="btnD" value="삭제" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup_mo(0,intg)%>')"  class="button">-->
											<input type="button" name="btnD" value="상품등록" onclick="popRegItem('<%=eCode%>','<%=arrGroup_mo(0,intg)%>','M')"  class="button">
											<% IF arrGroup_mo(5,intg) = 0 THEN %>
											<% 		Response.Write "<a href=""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup_mo(0,intg) &"','M');"">미리보기</a>"
											 %>
											<% END IF %>
										</td>					   									
									</tr>
									<%
									     intg = intg+sumi
									NEXT%>
									</table>
									<%END IF%> 
								</div>
								<!-- /3.그룹형-->
								<!-- 5.수작업-->
								<div id="divMFrm5" style="display:<%if not ( etemp_mo = 5 or etemp_mo = 6) then%>none<%end if%>;">
									<table border="0" cellpadding="1" cellspacing="3" class="a">
										<tr>
											<td>
											    <!-- <input type="button" value="이미지관리"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/','/event/');" class="input_b"> -->
											    <input type="button" value="이미지관리(신)"  onclick="jsManageEventImageNew('<%=eCode%>')" class="input_b">
											    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
											    <input type="button" value="이미지관리(구)"  onclick="jsManageEventImage('<%=eCode%>')" class="input_b">
											    <p>
											    <b>이미지관리(신)</b> : 새로이 변경된 이미지 관리<br>
											    <b>이미지관리(구)</b> : 기존에 저장된 이미지 리스트(이미지추가없음. 새로운 이미지 추가는 이미지관리(신)에서만.)<br>
											    ※ 이벤트 이미지 시스템 관리 차원에서 eventIMG 라는 새로운 폴더에 이벤트시작年 폴더를 추가하여 그 안에 이벤트코드별 폴더를 생성하게 됩니다.<br>
											    추후 몇달이 지난 뒤에 이미지관리(구)는 사용을 하지않고 이미지관리(신)만 사용하게 됩니다.<br>
											    그때까지는 불편사항이 있으시더라도 시스템관리 차원에 의한 조치이므로 양해바랍니다.
											</td>
										</tr>
										<tr>
										    <td><b>이미지 경로 http://<font color="RED">webimage.</font>10x10.co.kr/eventIMG/이벤트시작年/XXX/</b> 로 변경되었습니다.</td>
										</tr>
										<tr>
											<td><textarea name="tHtml5_mo" rows="25" style="width:100%;font-size:11px;"><%=ehtml5_mo%></textarea></td>
										</tr> 
									</table>
								</div>
								<!-- /5.수작업--> 
							</td>
						</tr>
						<tr>
						    <td bgcolor="#e3f1fb" width="100" align="Center">Exec File<br/> <span style="color:#602030;font-size:11px;">[ 개발 실행파일]</span></td>
							<td bgcolor="#ffffff"  >
						        <input type="radio" name="rdoEF_mo" value="0" <%if not blnExec_mo then%>checked<%end if%>>비실행 
						        <input type="radio" name="rdoEF_mo" value="1" <%if blnExec_mo then%>checked<%end if%>>실행 
						        <input type="text" name="sEFP_mo" size="60" class="text" value="<%=eExecFile_mo%>"> 
						    </td>
						</tr>
					</table>	
				</div>
				</td>
			</tr>
		</table>	 
	</td>
</tr>	
<tr>
	<td width="100%" align="right" >
		<% If etype<>"80" Then %><input type="image" src="/images/icon_save.gif"><% End If %>
		<a href="index.asp?menupos=<%=menupos%>&<%=strParm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
