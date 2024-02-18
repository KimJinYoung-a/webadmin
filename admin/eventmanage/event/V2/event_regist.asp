<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/event_regist.asp
' Description :  이벤트 개요 등록
' History : 2007.02.07 정윤정 생성
'           2015.03 이벤트 리뉴얼 정윤정  
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
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
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, etag, eonlyten, eisblogurl,ebrand,eSalePer
Dim ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,ebimg,etemp,emimg,ehtml,eisort,eiaddtype,edgid,edgid2,edgstat1,edgstat2,emdid,efwd,selPartner, dispcate
Dim enameEng, subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary,eNew
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm , sCateMid
Dim edpid, epsid, edgnm, emdnm,epsnm,edpnm
Dim isWeb, isMobile, isApp,enamesub, etype, isConfirm
dim maxDepth
dim ehtml5
dim tHtml5_mo, tHtml_mo, main_mo,emimg_mo,ehtml_mo,ehtml5_mo, efwd_mo
dim sWorkTag
Dim blnFull, blnIteminfo ,blnitemprice, evt_sortNo, blnWide  ,eDateView,eItemListType
Dim  eCmtCd,eIpsCd,eGfCd,eBSCd, rdCmt, eCmtMT, eCmtST, eIpsMT, eIpsST, eGfMT, eGfST, eBSMT, eBSST
dim arrText,intT
dim blnReqPublish
 dim chkeCmt, chkeIps, chkeGf, chkeBS 
eCode = Request("eC")
ekind = Request("eK")


maxDepth = 2 '전시카테고리 2depth까지 보여준다
elevel = 2 '중요도 보통으로 임시 설정
isWeb = True
isMobile = True
isApp = True
isConfirm = False
eItemListType = "1"
blnIteminfo = True
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
	eKind		= sKind
	etype		= requestCheckVar(Request("eventtype"),4)	'이벤트유형
	edgid  		= requestCheckVar(Request("sDgId"),32)		'담당 디자이너
	edgid2 		= requestCheckVar(Request("sDgId2"),32)		'서브 디자이너
	emdid  		= requestCheckVar(Request("sMdId"),32)		'담당 MD
	epsid  		= requestCheckVar(Request("sDgId"),32)		'담당 퍼블리셔
	edpid  		= requestCheckVar(Request("selDeId"),32)		'담당 개발자
		
	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무

	eOneplusone	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery= requestCheckVar(Request("chFreedelivery"),2)	'무료배송
	eBookingsell= requestCheckVar(Request("chBookingsell"),2)	'예약판매
	eDiary= requestCheckVar(Request("chDiary"),2)	'다이어리
	dispcate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리

	if emdid = "" then 
		emdid = session("ssBctId")
		emdnm = session("ssBctCname")
	end if	
	
	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&sMdNm="&eMDnm&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&dispcate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
	'#######################################
	
	esale = False
	egift= False
	ecoupon= False
	eonlyten= False
	eOneplusone= False
	eFreedelivery= False
	eBookingsell= False
	eDiary= False
	eNew= False
	ecomment = False
	ebbs 	= False
	eitemps	= False
	eisblogurl = False
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ekind =	cEvtCont.FEKind
	eman =	cEvtCont.FEManager
	escope =	cEvtCont.FEScope
	ename =	db2html(cEvtCont.FEName)
	enameEng =	db2html(cEvtCont.FENameEng) '이벤트 영문 추가
	subcopyK =	db2html(cEvtCont.FsubcopyK) '이벤트 영문 추가
	subcopyE =	db2html(cEvtCont.FsubcopyE) '이벤트 영문 추가
	enamesub	=	db2html(cEvtCont.FENamesub) '이벤트 타이틀 서브카피
	
	elevel =	cEvtCont.FELevel
	'estate =	cEvtCont.FEState
	eregdate =	cEvtCont.FERegdate
	 
 	
	isWeb	= cEvtCont.FIsWeb
	isMobile= cEvtCont.FIsMobile
	isApp	= cEvtCont.FIsApp

	etype = cEvtCont.FEType
	isConfirm = cEvtCont.FIsConfirm
	   
	if ekind = 19 then
	    isWeb = False
	    isMobile = True
	    isApp = True
	    ekind = 1
	elseif ekind = 25 then
	    isWeb = False
	    isMobile = False
	    isApp = True
	    ekind = 1
	elseif ekind = 26 then
	    isWeb = False
	    isMobile = True
	    isApp = False
	    ekind = 1
	elseif not (isWeb  or  isMobile  or isApp) or (isNull(isWeb) and isNull(isMobile) and isNull(isApp))  then 
	    isWeb = True
	    isMobile = False
	    isApp = False    
	    ekind = 1
    end if        
	      
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay 
	ecategory 	=	cEvtCont.FECategory
	tmp_cdl 		= cEvtCont.FECategory 
	dispcate	=	cEvtCont.FEdispcate
	esale 		= 	cEvtCont.FESale
	egift 		=	cEvtCont.FEGift
	ecoupon 	=	cEvtCont.FECoupon
	ecomment 	=	cEvtCont.FECommnet
	ebbs 		=	cEvtCont.FEBbs
	eitemps	 	=	cEvtCont.FEItemps
	eapply 		=	cEvtCont.FEApply
	eisort 		=	cEvtCont.FEISort
	 
	efwd 		=	db2html(cEvtCont.FEFwd)
	efwd_mo 		= db2html(cEvtCont.FEFwdMo) 

	ebrand			= cEvtCont.FEBrand
	etag		= db2html(cEvtCont.FETag)
 	eonlyten	= cEvtCont.FSisOnlyTen
 	eDiary		= cEvtCont.FSisDiary
 	eNew			= cEvtCont.FSisNew
 	eisblogurl	= cEvtCont.FSisGetBlogURL

	eOneplusone		= cEvtCont.FEOneplusOne
	eFreedelivery	= cEvtCont.FEFreedelivery
	eBookingsell	= cEvtCont.FEBookingsell
 
	'edgid 			= cEvtCont.FEDgId 
  	'emdid 			= cEvtCont.FEMdId 
	'epsid			= cEvtCont.FEPsId
	'edpid			= cEvtCont.FEDpId
	
	'edgnm 			= cEvtCont.FEDgName
  	'emdnm 			= cEvtCont.FEMdName 
	'epsnm			= cEvtCont.FEPsName
	'edpnm			= cEvtCont.FEDpName
	
	sWorkTag		= cEvtCont.FWorkTag
	blnFull			= cEvtCont.FEFullYN
 	blnWide			= cEvtCont.FEWideYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN 
 	blnitemprice	 = cEvtCont.FEItempriceYN
	eDateView		= cEvtCont.FEdateview
	eItemListType = cEvtCont.FEListType
	blnReqPublish   = cEvtCont.FisReqPublish
	set cEvtCont = nothing 
 
	If eItemListType = "" OR isNull(eItemListType) Then eItemListType = "1"

	    
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
	END IF
	
	
    
	if (ekind = 1 or ekind = 23) and (eSale or ecoupon ) then
	    dim tmpename
	    tmpename = Split(ename,"|") 
	  			 
	  	if Ubound(tmpename)>0 then
		    ename = tmpename(0)
		    eSalePer = tmpename(1)
		 end if

    end if
END IF 

 	if eisort = "" or isNull(eisort)  then eisort = 3
   if eCmtST = ""   then
	   eCmtST = "정성껏 코멘트를 남겨주신     분을 추첨하여           를 선물로 드립니다." 
    end if

    
     dim idepartmentid, sdepartmentname,clsMem
    '부서명 가져오기
set clsMem = new CTenByTenMember
	clsMem.Fuserid = emdid
	clsMem.fnGetDepartmentInfo
	idepartmentid		= clsMem.Fdepartment_id
 	sdepartmentname = clsMem.FdepartmentNameFull 
 set clsMem = nothing
	 
%>
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
       
       if(frm.blnMobile.checked || frm.blnApp.checked){
        if(!frm.subsEN.value){
            alert("Mobile/App 의 서브카피를 입력해주세요");
            frm.subsEN.focus();
            return false;
        }
    }
    
//	if(!frm.eventscope.value) {
//		alert("이벤트 범위를 선택해주세요");
//		frm.chkEscope[0].focus();
//		return false;
//	}

	  if(!frm.sEN.value){
	  	alert("이벤트명을 입력해주세요");
	  	frm.sEN.focus();
	  	return false;
	  }

	  if(frm.sEN.value.length > 80){
		alert("이벤트명은 60자까지만 가능합니다.다시 입력해주세요.");
	 	frm.sEN.focus();
	  	return false;
	  }

	   if(frm.sENEng.value.length > 120){
		alert("영문이벤트명은 120자까지만 가능합니다.다시 입력해주세요.");
	 	frm.sENEng.focus();
	  	return false;
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


	  	if(frm.sSD.value < nowDate){
	  		alert("시작일이 현재일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
	  		frm.sSD.focus();
	  		return false;
	  	}

		if((frm.chComm.checked||frm.chBbs.checked||frm.chItemps.checked||frm.isblogurl.checked)&&frm.sPD.value=="") {
	  		alert("당첨자 발표일을 선택해주세요 ");
	  		frm.sPD.focus();
	  		return false;
		}

		if(!frm.sMdId.value){
			alert('담당자를 지정하세요');
			return false;
		}

//	  if(!frm.eCT.value){
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
		 
		if((chk=='1' || chk=='23')  && (document.frmEvt.chSale.checked || frm.chCoupon.checked)){
		     document.all.spSale.style.display = "";
		}else{
		     document.all.spSale.style.display = "none";
	        document.frmEvt.sSP.value ="";
		}
	}

//-- jsLastEvent : 지난 이벤트 불러오기 --//
	function jsLastEvent(){
	  var winLast,eKind;
	  eKind = document.frmEvt.eventkind.options[document.frmEvt.eventkind.selectedIndex].value;
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=800,height=600, scrollbars=yes')
	  winLast.focus();
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
    			$("#divFrm5").show(); 
    		}else{
    			$("#divFrm1").show();
				$("#w_slide").show();
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
	
	function jsChkChannel(sChannel){ 
	    if (sChannel =="P"){
	        if(document.frmEvt.blnWeb.checked){
	            document.all.divPC1.style.display="";
	        }else{
	            document.all.divPC1.style.display="none";
	        }
	    }
	    if (sChannel =="M" || sChannel =="A"){
	        if(document.frmEvt.blnMobile.checked || document.frmEvt.blnApp.checked){
	            document.all.divMA1.style.display=""; 
	        }else{
	            document.all.divMA1.style.display="none"; 
	        }
	    }
	}
	
 
    function jsChkSale(){
	    var frm = document.frmEvt; 
	    if(( frm.eventkind.options[frm.eventkind.selectedIndex].value == 1 ||  frm.eventkind.options[frm.eventkind.selectedIndex].value == 23 )   && (document.frmEvt.chSale.checked|| frm.chCoupon.checked)){ 
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
</script>
<form name="frmEvt" method="post"  action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="imod" value="I">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="strparm" value="<%=strparm%>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
 
<tr>
	<td ><input type="button" value="지난 이벤트 내용 불러오기" class="button" onClick="jsLastEvent();"></td>
</tr>
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
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
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 타입</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chSale" <%IF esale   THEN%>checked<%END IF%> value="1" onClick="jsChkSale();">할인
		   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('g');">사은품
		   		<input type="checkbox" name="chCoupon" <%IF ecoupon   THEN%>checked<%END IF%> value="1" onClick="jsChkSale();">쿠폰
		   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten   THEN%>checked<%END IF%> value="1">Only-TenByTen
		   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone   THEN%>checked<%END IF%> value="1">1+1
				<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery   THEN%>checked<%END IF%> value="1">무료배송
				<input type="checkbox" name="chBookingsell" <%IF eBookingsell THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('b');">예약판매
				<input type="checkbox" name="chDiary" <%IF eDiary  THEN%>checked<%END IF%> value="1">DiaryStory
				<input type="checkbox" name="chNew" <%IF eNew   THEN%>checked<%END IF%> value="1">런칭
		   		</td>
			</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기능</td>
		   		<td bgcolor="#FFFFFF">
		   		<input type="checkbox" name="chComm" <%IF ecomment THEN%>checked<%END IF%> value="1" onClick="jsChkTitle('c');">코멘트
		   		<input type="checkbox" name="chBbs" <%IF ebbs  THEN%>checked<%END IF%> value="1" >게시판
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
		   	<tr id="eNameTr_A">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>">
		   			<span id="spSale"  style="display:<%if not ((ekind="1" or ekind="23")  and (esale or ecoupon )) then %>none<%end if%>;<%if esale then%>color:red;<%else%>color:green;<%end if%>"><b> 할인율: </b><input type="text" name="sSP" value="<%=eSalePer%>" size="10" class="text" >(예:40%~)</span>
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>영문 이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:none;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명<br>및 할인율</B></td>
		   		<td bgcolor="#FFFFFF">
		   			이벤트명: <input type="text" name="sEDN" size="50" maxlength="50" value=""><br>
		   			영문이벤트명: <input type="text" name="sEDNEng" size="50" maxlength="50" value=""><br>
		   			할인율: 최저 <input type="text" name="sSDc" size="4" value="0" style="text-align:right;">% ~
		   			최고 <input type="text" name="sMDc" size="4" value="0" style="text-align:right;">%<br>
		   			<font color=gray>※브랜드 스트리트에 보여질 할인율입니다. 실제로 상품에는 적용되지 않으니 상품에는 따로 할인을 적용해주세요.<br>이벤트 링크는 브랜드 스트리트로 연결되니 반드시 상세내용에 브랜드를 선택해주세요.</font>
		   		</td>
		   	</tr>  
		   	<tr>
		   		<td rowspan="2" align="center" bgcolor="<%= adminColor("tabletop") %>"><B>기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   			시작일 : <input type="text" id="termSdt" name="sSD" size="10" />
							<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
		   			~ 종료일 : <input type="text" id="termEdt" name="sED" size="10" />
							<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkEnd_trigger" onclick="return false;" />
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "termSdt", trigger    : "ChkStart_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "termEdt", trigger    : "ChkEnd_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					&nbsp;&nbsp;<input type="checkbox" name="endlessview"  value="Y"> <a title="상시노출 설정시 기간이 지난 이벤트도 이벤트 종료 안내 레이어 안뜨도록 설정">상시노출</a>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			당첨 발표일 : <input type="text" id="priceDt" name="sPD" size="10">
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
		   			<%sbGetOptStatusCodeValue "eventstate",estate,false,""%>
		   			<%''sbGetOptStatusCodeAuth "eventstate",0,"N",""%>
		   		</td>
		   	</tr>
			<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>중요도</b></td>
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
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">퍼블리싱</td>
		   		<td bgcolor="#FFFFFF"><input type="checkbox" name="chkReqP" value="1" <%if blnReqPublish then%>checked<%end if%>>  퍼블리싱 요청 <img src="/images/admin_help.gif" style="cursor:hand;" onClick="jsPubHelp();"></td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">디자이너 작업구분</td>
		   		<td bgcolor="#FFFFFF"><input type="text" name="sWorkTag" size="20" maxlength="16" class="text" value="<%=sWorkTag%>">
		   		    <input type="checkbox" name="chkMB"  onClick="jsChkMBReq();" <%if left(sWorkTag,4) ="[★★]" then%>checked<%end if%>> 모바일배너 요청시 체크    
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
		   			<input type="text" name="eLC" size="4" maxlength="10">
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
				    <div id="divPC1" style="display:<%if not isWeb then%>none<%end if%>;">
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
					  	<tr id="eNameTr_BL" style="display:none;"> 
            				<td align="center"  bgcolor="#FAECC5">브랜드이벤트 링크</td>
            				<td bgcolor="#FFFFFF"> 
            				 <input type="hidden" name="elType" value="I" > 
            				 <input type="text" id="elUrl" name="elUrl" size="60" maxlength="128" value="" class="text" > 
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
					   		<td bgcolor="#FFFFFF"> <input type="text" name="subsEN" size="70" maxlength="70" value="<%=enamesub%>" OnKeyUp="jsAddByte(this);"> <input type="text" name="subSize" size="3" value="" class="text_ro" style="text-align:right" readonly>Byte 
					   		  <p style="color:#602030;font-size:11px;"> [ 최대 70byte까지 등록가능합니다. ]</p>
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
								<div id="dvTxT">
								<table border="0" cellpadding="3" cellspacing="1" class="a" width="100%">  
								<tr>
									<td colspan="2">
										<% if rdCmt="" THEN rdCmt=1%>
										<div id="dvCmt"  style="display:<%IF not ecomment THEN %>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%"  bgcolor="#BDBDBD">  
												<th bgcolor="#e3f1fb" colspan="2" align="left">코멘트 
													<input type="radio" name="rdCmt" value="1" <%if rdCmt="1" THEN%>checked<%END IF%>>코멘트
													<input type="radio" name="rdCmt" value="2" <%if rdCmt="2" THEN%>checked<%END IF%>>테스터 코멘트
													<input type="checkbox" name="chkeCmt" value="0" <%if chkeCmt="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eCmt');">사용안함(이미지로 등록)
												</th>  
											<tr>
													<td bgcolor="#e3f1fb">주제</td><td bgcolor="#FFFFFF"><input type="text" name="eCmtMT" size="60" value="<%=eCmtMT%>" maxlength="200"></td>
												</tr>
												<tR >
													<td bgcolor="#e3f1fb">당첨자수/<br/>사은품</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eCmtST" class="Textarea"><%=db2Html(eCmtST)%></textarea></td>
												</tr>
											 </table>
										</div>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<div id="dvIps" style="display:<%IF not eitemps THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%" bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">상품후기
												      <input type="checkbox" name="chkeIps" value="0" <%if chkeIps="0" THEN%>checked<%END IF%> onClick="jsCmtStyle('eIps');">사용안함(이미지로 등록)
												     </th> 
												<tr>
													<td bgcolor="#e3f1fb">주제</td><td bgcolor="#FFFFFF"><input type="text" name="eIpsMT" size="60" value="<%=eIpsMT%>" maxlength="200"></td>
												</tr>
												<tR>
													<td bgcolor="#e3f1fb">당첨자수/<br/>사은품</td><td bgcolor="#FFFFFF"><textarea cols="70" rows="3" name="eIpsST" class="textarea"><%=db2Html(eIpsST)%></textarea></td>
												</tr>
											 </table>
										</div>
									</td>
								</tr>
								 <tr>
									<td colspan="2">
										<div id="dvGf"  style="display:<%IF not egift THEN%>none<%end if%>;">
											<table border="0" cellpadding="5" cellspacing="1" class="a" width="100%"  bgcolor="#BDBDBD">  
												<th colspan="2" align="left" bgcolor="#e3f1fb">사은품 
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
												<th colspan="2" align="left"  bgcolor="#e3f1fb">예약판매 
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
						<tr id="itemevt" style="display:none;">
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
												<input type="radio" name="ietag" value="7"/> 할인 <input type="radio" name="ietag" value="2"/> 쿠폰 <input type="text" size="5" name="ietagval" value=""/> <input type="radio" name="ietag" value="1"/> GiFT <input type="radio" name="ietag" value="4"/> 무료배송 <input type="radio" name="ietag" value="5"/> 1:1 <input type="radio" name="ietag" value="6"/> 1+1 <input type="radio" name="ietag" value="3"/> 예약배송
											</td>
										</tr>
										<tr>
											<td align="center" bgcolor="#e3f1fb"><b>프로모션 내용</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="mcopy" maxlength="50" /><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: 오늘 단 하루, UDH-02 전기렌지 증정! )</div></td>
										</tr>
										<tr>
											<td align="center" bgcolor="#e3f1fb"><b>서브 내용</B></td>
											<td bgcolor="#FFFFFF"><input type="text" size="70" name="scopy" maxlength="50" /><div style="color:#602030;font-size:11px;padding-top:5px;">( ex: 선착순 100명/ 소진 시 조기종료 )</div></td>
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
		</table>	 
	</td>
</tr>	
<tr>
	<td width="100%" align="right" >
		<input type="image" src="/images/icon_save.gif">
		<a href="index.asp?menupos=<%=menupos%>&<%=strParm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
