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
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'--------------------------------------------------------
' 변수선언
'--------------------------------------------------------
Dim eCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, echkdisp, eusing, etag, eonlyten, eisblogurl
Dim ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,elktype,elkurl,ebimg,etemp,emimg,ehtml,ehtml5, eisort,eiaddtype, edid, emid ,efwd,ebrand,eicon,ebimg2010
Dim selPartner,dopendate,dclosedate, sWorkTag, ebimgMo, eDispCate, eDateView , ebimgToday , ebimgMo2014 , enamesub
Dim intI
Dim arrGift, intg,blngift
Dim eFolder, backUrl
dim gimg : gimg = ""
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim ecommenttitle, elinkcode
Dim strparm , sCateMid
Dim cEGroup, arrGroup,intgroup,strG, blngroup
Dim blnFull, blnIteminfo ,blnitemprice, evt_sortNo, blnWide
Dim enameEng , subcopyK , subcopyE
Dim eOneplusone , eFreedelivery , eBookingsell, eDiary
Dim eEtcitemid , eEtcitemimg, eItemListType
Dim eitemid
	eCode		= requestCheckVar(Request("eC"),10)	'이벤트코드
	blnFull		= False
	blnWide		= False
	blnIteminfo	= True
	blnitemprice = False
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
	edid  		= requestCheckVar(Request("selDId"),32)		'담당 디자이너
	emid  		= requestCheckVar(Request("selMId"),32)		'담당 MD

	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무

	eOneplusone	 	= requestCheckVar(Request("chOneplusone"),2)	'oneplusone
	eFreedelivery	= requestCheckVar(Request("chFreedelivery"),2)	'무료배송
	eBookingsell	= requestCheckVar(Request("chBookingsell"),2)	'예약판매
	eDiary	= requestCheckVar(Request("chDiary"),2)	'다이어리
	edispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리

	strparm  = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&selDId="&edid&"&selMId="&emid&_
				"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOneplusone="&eOneplusone&"&chFreedelivery="&eFreedelivery&"&chBookingsell="&eBookingsell&"&disp="&edispCate&"&chOnlyTen="&eonlyten&"&chDiary="&eDiary
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
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	echkdisp 		= cEvtCont.FChkDisp
	tmp_cdl 		= cEvtCont.FECategory
	tmp_cdm			= cEvtCont.FECateMid
	eDispCate		= cEvtCont.FEDispCate

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
	if etemp = 5 or etemp = 6  THEN	'수작업 이벤트 일 경우 처리
		ehtml5 		= db2html(cEvtCont.FEHtml)
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml)
	end if
	eisort 			= cEvtCont.FEISort
	edid 			= cEvtCont.FEDId
	emid 			= cEvtCont.FEMId
	efwd 			= db2html(cEvtCont.FEFwd)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode   	= cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
 	blnFull			= cEvtCont.FEFullYN
 	blnWide			= cEvtCont.FEWideYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)
 	eonlyten		= cEvtCont.FSisOnlyTen
 	eisblogurl		= cEvtCont.FSisGetBlogURL
 	sWorkTag		= cEvtCont.FWorkTag

	blnitemprice	 = cEvtCont.FEItempriceYN

	eOneplusone	=	cEvtCont.FEOneplusOne
	eFreedelivery		=	cEvtCont.FEFreedelivery
	eBookingsell		=	cEvtCont.FEBookingsell
	eDiary = cEvtCont.FSisDiary

	eEtcitemid			=	cEvtCont.FEtcitemid
	eEtcitemimg		=	cEvtCont.FEtcitemimg
	eDateView		= cEvtCont.FEdateview
	eItemListType = cEvtCont.FEListType

	set cEvtCont = nothing
	IF elinkcode = 0 THEN elinkcode = ""

	 set cEGroup = new ClsEventGroup
	 	cEGroup.FECode = eCode
	  	arrGroup = cEGroup.fnGetEventItemGroup
	 set cEGroup = nothing

	 blngroup = False
	 IF isArray(arrGroup) THEN blngroup = True

	 If eItemListType = "" OR isNull(eItemListType) Then eItemListType = "1" End If
%>
<script type="text/javaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javaScript" src="/js/jquery.iframe-auto-height.js"></script>
<script type="text/javascript">
<!--
//-- jsEvtSubmit(form 명) : 이벤트 수정처리 --//
	function jsEvtSubmit(frm){

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

	if(!frm.eventscope.value) {
		alert("이벤트 범위를 선택해주세요");
		frm.chkEscope[0].focus();
		return false;
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

	    if(!frm.eCT.value){
	  		if(GetByteLength(frm.eCT.value) > 200){
	  			alert("comment title은 200자 이내로 작성해주세요");
	  			frm.eCT.focus();
	  			return false;
	  		}
	  	}


  		if(GetByteLength(frm.eTag.value) > 250){
  			alert("Tag는 250자 이내로 작성해주세요");
  			frm.eTag.focus();
  			return false;
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
		var blnSale, blnGift, blnCoupon;
		blnSale= "<%=esale%>";
		blnGift= "<%=egift%>";
		blnCoupon= "<%=ecoupon%>";

		if (sName!="sPD" && (blnSale=="True" || blnGift=="True"|| blnCoupon=="True")){
			if(confirm("기간을 변경시 할인, 사은품, 쿠폰에도 적용이 됩니다. 기간을 변경하시겠습니까?")){
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
			}
		}else{
				winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
				winCal.focus();
		}
	}

//-- jsImgDel(이미지 종류) : 이미지 화면에서 안보이게 --//
	function jsImgDel(sType){
	 if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	 	if(sType == "B"){
	 		document.frmEvt.oldBimg.value="";
	 		document.all.imgB.style.display="none";
	 	}else{
	 		document.frmEvt.oldMimg.value="";
	 		document.all.imgM.style.display="none";
	 	}
	 }
	}


//-- jsChangeFrm : 템플릿에 따른 화면 설정 변경:완전 수작업 --//
	function jsChangeFrm(iVal){
		$("div[id^='divFrm']").hide();

		if(iVal == 3 || iVal == 7){
			iframG.location.href = "iframe_eventitem_group.asp?eC=<%=eCode%>&ekind=<%=ekind%>";
			$("#divFrm3").show();
			$('#iframG').load(function() {
				$(this).height($(this).contents().find('body')[0].scrollHeight+30);
			});
		}else if(iVal == 5 || iVal == 6 ){
			//iframG.location.href = "about;blank";
			$("#divFrm5").show();
		}else{
			//iframG.location.href = "about;blank";
			$("#divFrm1").show();
		}
	}


//-- jsImgView : 이미지 확대화면 새창으로 보여주기 --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}


	function jsSetImg(sFolder, sImg, sName, sSpan){
		document.domain ="10x10.co.kr";
		var winImg;
		winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	//-- jsChkDisp : 상세화면 보기 --//
	function jsChkDisp(){
	 if(document.frmEvt.chkDisp.checked){
	  	eDetail.style.display = "";
	  }else{
	  	eDetail.style.display = "none";
	  }
	}

	function jsChkSubj(chk){
		if(chk=='16') {
			//브랜드할인일경우에는 제목 대신 할인율 범위로 표시
			eNameTr_A.style.display = "none";
			eNameTr_C.style.display = "none";
			eNameTr_B.style.display = "";
		} else {
			eNameTr_A.style.display = "";
			eNameTr_C.style.display = "";
			eNameTr_B.style.display = "none";
		}
	}

	function jsManageEventImage(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/eventManageDir.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function jsManageEventImageNew(evtcode){
	    var popwin = window.open('<%= uploadImgUrl %>/linkweb/event_admin/eventManageDir_new.asp?evtcode=' + evtcode,'eventManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
	}

	function popCommentXLS(ecd) {
		 var wCmtXls = window.open('pop_event_Comment_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	//2015.05.19 유태욱(푸드파이터 이벤트용으로 임시 생성-이벤트 종료후 삭제예정)
	function popCommentXLS2(ecd) {
		 var wCmtXls = window.open('pop_event_Comment_xls_2.asp?eC='+ecd,'pXls','width=400,height=150');
		 wCmtXls.focus();
	}

	function popBBSXLS(ecd) {
		 var wBBSXls = window.open('pop_event_board_xls.asp?eC='+ecd,'pXls','width=400,height=150');
		 wBBSXls.focus();
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

	//이벤트타입 선택 해제시 설정 변경 경고처리
	function jsChkType(sType,frm){
		if(!frm.checked){
			if(confirm(sType +"설정을 해제할 경우 해당 "+sType+"관리 상태도 종료처리됩니다. 설정을 해제하시겠습니까?")){
				return;
			}else{
				frm.checked = true;
			}
		}
	}
	// 배너 링크설정 Eable
	function jsEvtLink(bln){
		if (bln) {
			$("#elUrl").attr("readonly",true);
			$("#elUrl").attr("class","");
			$("#elUrl").addClass("text_ro");
		}else{
			$('#elUrl').removeAttr('readonly');
			$("#elUrl").attr("class","");
			$("#elUrl").addClass("text");
		}
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

	function workerlist()
	{
		var openWorker = null;
		var worker = frmEvt.selMId.value;
		openWorker = window.open('PopWorkerList.asp?worker='+worker+'&department_id=7','openWorker','width=700,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function workerDel()
	{
		var frm = document.frmEvt;

		frm.selMId.value = "";
		frm.doc_workername.value = "";
	}

	// 상품복사 리스트팝업
	function jsItemcopylist(){
		var winLast,eKind;
		winLast = window.open('pop_event_itemlist.asp?menupos=<%=menupos%>&eC=<%=eCode%>','pLast','width=550,height=600, scrollbars=yes')
		winLast.focus();
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

	function chgex(v){
		for (i=1;i<4;i++)
		{
			if (v == i)
			{
				$("#notice"+i).css("display","block");
			}else{
				$("#notice"+i).css("display","none");
			}
		}
	}
//-->
</script>
<form name="frmitemclear" method="post">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="IC">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<form name="frmEvt" method="post" action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="imod" value="U">
<input type="hidden" name="ban" value="<%=ebimg%>">
<input type="hidden" name="ban2010" value="<%=ebimg2010%>">
<input type="hidden" name="banMo" value="<%=ebimgMo%>">
<input type="hidden" name="banMoToday" value="<%=ebimgToday%>">
<input type="hidden" name="banMoList" value="<%=ebimgMo2014%>">
<input type="hidden" name="icon" value="<%=eicon%>">
<input type="hidden" name="main" value="<%=emimg%>">
<input type="hidden" name="gift" value="<%=gimg%>">
<input type="hidden" name="strparm" value="<%=strparm%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="etcitemban" value="<%=eEtcitemimg%>">
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<tr>
	<td>  <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 개요 등록  </font></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트코드</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0" >
		   			<tr>
		   				<td>
							<%=eCode%>
							<input type="button" value="상품 복사" onclick="jsItemcopylist();" class="button"/>
							<%
								Select Case ekind
									Case "19","25","26"
										Response.write "<input type='button' value='상품초기화' onclick='jsItemclear();' class='button' />"
								End select
							%>
						</td>
		   				<td>
						<%
							'이벤트 종류에 따른 프론트링크 페이지 선택
							Select Case ekind
								Case "7"		'위클리코디
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "13"		'상품 이벤트
									Response.Write "<td><a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & eitemid & "' target='_blank'>미리보기</a></td>"
								Case "14"		'소풍가는길
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "16"		'브랜드 할인행사
									Response.Write "<td><a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & ebrand & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>미리보기</a></td>"
								Case "22"		'DAY&(데이앤드)
									Response.Write "<td><a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case "26"		'모바일
									Response.Write "<td><a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
								Case Else		'쇼핑찬스 및 기타
									Response.Write "<td><a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "' target='_blank'>미리보기</a></td>"
							End Select
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
		   					<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드" onClick="window.open('pop_event_votelist_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle" title ="xls 다운로드 회원기반">
							<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드 비회원"  title ="xls 다운로드 비회원" onClick="window.open('pop_event_votelist_guest_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">

						<% If eCode = "65010" and session("ssBctId") = "stella0117" then %>
							<img src="/images/icon_excel_reply.gif" alt="냉동실을부택해 Excel다운로드" onClick="popCommentXLS2(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
						<% End if %>
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
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>종류</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventkind",ekind,False,"onChange=javascript:jsChkSubj(this.value);"%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>주체</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventmanager",eman,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>범위</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="hidden" name="eventscope" value="<%=escope%>">
		   			<label><input type="checkbox" name="chkEscope" <%=chkIIF((escope="1" or escope="2"),"checked","")%> onclick="jsSetPartner()"> 10x10</label>
		   			<label><input type="checkbox" name="chkEscope" <%=chkIIF((escope="1" or escope="3"),"checked","")%> onclick="jsSetPartner()"> 제휴몰</label>
		   			<span id="spanP" style="display:<%=chkIIF((escope="1" or escope="3"),"","none")%>">
		   			<select name="selP">
		   				<option value="">--제휴몰 전체--</option>
		   				<% sbOptPartner selPartner%>
		   			</select>
		   			</span>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A" style="display:<% if ekind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sEN" size="80" maxlength="120" value="<%=ename%>">
		   		</td>
		   	</tr>
			<tr style="display:<% if ekind="16" then Response.Write "none" %>;">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>2014이벤트<br/>서브카피</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="subsEN" size="60" maxlength="60" value="<%=enamesub%>">
		   		</td>
		   	</tr>
			<tr id="eNameTr_C">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>영문 이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sENEng" size="60" maxlength="60" value="<%=enameEng%>">
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_B" style="display:<% if ekind<>"16" then Response.Write "none" %>;">
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
		   			<font color=gray>※브랜드 스트리트에 보여질 할인율입니다. 실제로 상품에는 적용되지 않으니 상품에는 따로 할인을 적용해주세요.<br>이벤트 링크는 브랜드 스트리트로 연결되니 반드시 상세내용에 브랜드를 선택해주세요.</font>
		   		</td>
		   	</tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>서브 카피</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<table width="100%" border="0" cellpadding="0" cellspacing="0">
		   			<tr>
		   				<td width="50%" style="padding-right:3px;"><textarea name="subcopyK" style="width:100%; height:80px;" onclick="if(this.value=='한글')this.value='';" onblur="if(this.value=='')this.value='한글';" value="<%=subcopyK%>"><%=chkiif(subcopyK="","한글",subcopyK)%></textarea></td>
		   				<td width="50%"><textarea name="subcopyE" style="width:100%; height:80px;" onclick="if(this.value=='영문')this.value='';" onblur="if(this.value=='')this.value='영문';" value="<%=subcopyE%>"><%=chkiif(subcopyE="","영문",subcopyE)%></textarea></td>
		   			</tr>
		   			</table>
		   		</td>
			</tr>
		   	<tr>
		   		<td rowspan="2" align="center" bgcolor="<%= adminColor("tabletop") %>"><B>기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<%IF estate = 9 THEN%>
		   			시작일 : <%=esday%><input type="hidden" name="sSD" size="10" value="<%=esday%>">
		   			~ 종료일 : <%=eeday%> <input type="hidden" name="sED" value="<%=eeday%>" size="10" >
		   		<%ELSE%>
		   			시작일 : <input type="text" name="sSD" size="10" value="<%=esday%>" onClick="jsPopCal('sSD');"  style="cursor:hand;">
		   			~ 종료일 : <input type="text" name="sED" value="<%=eeday%>" size="10" onClick="jsPopCal('sED');" style="cursor:hand;">
		   		<%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td  bgcolor="#FFFFFF">
		   			당첨 발표일 : <input type="text" name="sPD" value="<%=epday%>" size="10" onClick="jsPopCal('sPD');" style="cursor:hand;">
		   			(당첨자가 있는 경우에만 등록)
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>상태</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%
		   				sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''if ekind="22" then
		   				''	'//데이앤드는 디자인파트만 사용해서 기존대로
		   				''	sbGetOptStatusCodeValue "eventstate",estate,false,""
		   				''else
		   				''	sbGetOptStatusCodeAuth "eventstate",estate,"M",""
		   				''end if
		   			%>
		   			<input type="hidden" name="eOD" value="<%=dopendate%>">
		   			<input type="hidden" name="eCD" value="<%=dclosedate%>">
		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  오픈처리일 : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  종료처리일 : <%=dclosedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>중요도</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%sbGetOptEventCodeValue "eventlevel",elevel,False,""%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>정렬번호</b></td>
		   		<td bgcolor="#FFFFFF">
		   			<input type="text" name="sortNo" value="<%=evt_sortNo%>" size="6" maxlength="5" style="text-align:right;" />
		   			(※숫자가 클수록 우선표시 됩니다. / Day&:회차)
		   		</td>
		   	</tr>
		   		<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>내용</b></td>
		   		<td bgcolor="#FFFFFF">
		   			상세내용 추가등록 <input type="checkbox" name="chkDisp" onClick="jsChkDisp();" <%IF echkdisp= 1 THEN%>checked<%END IF%>>
		   		</td>
		   	</tr>
		</table>
	</td>

</tr>
<tr>
	<td>
	 <div id="eDetail" style="display:<%IF echkdisp<> 1 THEN%>none;<%END IF%>">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">관리 카테고리</td>
					   		<td bgcolor="#FFFFFF">
					   			<%'DrawSelectBoxCategoryOnlyLarge "selCategory", ecategory,"" %>
					   			<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100"  align="center" bgcolor="<%= adminColor("tabletop") %>">전시 카테고리</td>
					   		<td bgcolor="#FFFFFF">
					   			<%=fnDispCateSelectBox(1,"","dispcate",eDispCate,"") %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
					   		<td bgcolor="#FFFFFF">
					   			<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 타입</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="checkbox" name="chSale" <%IF esale THEN%>checked onClick="jsChkType('할인',this);"<%END IF%> value="1">할인
					   		<input type="checkbox" name="chGift" <%IF egift  THEN%>checked  onClick="jsChkType('사은품',this);"<%END IF%> value="1">사은품
					   		<input type="checkbox" name="chCoupon" <%IF ecoupon THEN%>checked  onClick="jsChkType('쿠폰',this);"<%END IF%> value="1">쿠폰
					   		<input type="checkbox" name="chOnlyTen" <%IF eonlyten THEN%>checked<%END IF%> value="1">Only-TenByTen
					   		<input type="checkbox" name="chOneplusone" <%IF eOneplusone THEN%>checked<%END IF%> value="1">1+1
					   		<input type="checkbox" name="chFreedelivery" <%IF eFreedelivery THEN%>checked<%END IF%> value="1">무료배송
					   		<input type="checkbox" name="chBookingsell" <%IF eBookingsell THEN%>checked<%END IF%> value="1">예약판매
					   		<input type="checkbox" name="chDiary" <%IF eDiary THEN%>checked<%END IF%> value="1">DiaryStory
					   		</td>
						</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">이벤트 기능</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="checkbox" name="chComm" id="chComm" <%IF ecomment THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">코멘트
					   		<input type="checkbox" name="chBbs" <%IF ebbs THEN%>checked<%END IF%> value="1" >게시판
					   		<input type="checkbox" name="chItemps" <%IF eitemps THEN%>checked<%END IF%> value="1" >상품후기
					   		<input type="checkbox" name="isblogurl" id="isblogurl" <%IF eisblogurl THEN%>checked<%END IF%> value="1" onClick="jsChkBlogEnable()">Blog URL
					   		<!--<input type="checkbox" name="chApply" <%IF eapply THEN%>checked<%END IF%> value="1" >응모-->
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트 링크 타입</td>
					   		<td bgcolor="#FFFFFF">
					   			<label><input type="radio" name="elType" value="E" onclick="jsEvtLink(true);"  <% IF elktype="E" Then %>checked<% End IF %> >이벤트</label>
					   			<label><input type="radio" name="elType" value="I" onclick="jsEvtLink(false);" <% IF elktype="I" Then %>checked<% End IF %>>직접입력</label>
					   			&nbsp;<input type="text" name="elUrl" id="elUrl" size="40" maxlength="128" value="<%= elkurl %>" <% IF elktype="E" THEN%>class="text_ro" readOnly<%ELSE%>class="text"<%END IF %>>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center"  bgcolor="<%= adminColor("tabletop") %>">상품정렬방법</td>
					   		<td bgcolor="#FFFFFF">
					   			<%sbGetOptEventCodeValue "itemsort",eisort,False,""%>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당디자이너</td>
					   		<td bgcolor="#FFFFFF">
					   			<%sbGetDesignerid "selDId",edid,""%>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당자</td>
					   		<td bgcolor="#FFFFFF"><% sbGetwork "selMId",emid,"" %></td>
					   	</tr>
					   	<tr>
					   		<td align="center" bgcolor="<%= adminColor("tabletop") %>">작업전달사항</td>
					   		<td bgcolor="#FFFFFF">
					   			작업구분 <input type="text" name="sWorkTag" size="20" maxlength="16" value="<%= sWorkTag %>" class="text"> <font color="darkgray">(for Designer)</font><br />
					   			<textarea name="tFwd" rows="15" style="width:100%;font-size:12px;font-family:'Malgun Gothic',dotum;"><%=efwd%></textarea>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">Comment Title</td>
					   		<td bgcolor="#FFFFFF">
					   			(200자 이내)		   			<Br>
					   			<textarea name="eCT" rows="2" style="width:100%;"><%=ecommenttitle%></textarea>
					   		</td>
					   	</tr>
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
					   			<% if sKind="19" then Response.Write " <font color=darkred>※ PC버전 이벤트 번호 입력시 코멘트 연동</font>" %>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">페이스북 앱 연결</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" value="페이스북 앱 연결 등록창" class="button" onClick="window.open('pop_event_facebookapp.asp?ecode=<%=eCode%>','facebookpop','width=500,height=400');">

					   		</td>
					   	</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td style="padding: 10 0 5 0"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 화면이미지 등록</td></tr>
			<tr>
				<td>
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">화면구성</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkFull" value="0" <%IF not blnFull THEN%>checked<%END IF%> onclick="if(this.checked) chkWide.checked=false;"> 왼쪽 메뉴&nbsp;&nbsp;
					   			<input type="checkbox" name="chkWide" value="1" <%IF blnWide THEN%>checked<%END IF%> onclick="if(this.checked) chkFull.checked=false;"> 와이드 페이지
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">상품리스트 스타일<br/>(Mobile,App 용)</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>> 격자형&nbsp;&nbsp;
					   			<input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>> 리스트형&nbsp;&nbsp;
					   			<input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>> BIG형
					   		</td>
					   	</tr>

					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">상품정보</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkIteminfo"  value="1"  <%IF blnIteminfo THEN%>checked<%END IF%>> 사용함
					   		</td>
					   	</tr>
						<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">상품 가격정보<br/><font color="#BB8866">[쿠폰및 할인가<br/>노출여부]</font></td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="chkItemprice"  value="1"  <%IF blnitemprice THEN%>checked<%END IF%>> 노출안함
					   		</td>
					   	</tr>
						<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">이벤트 기간<br/>노출여부</td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="checkbox" name="dateview"  value="1"  <%IF eDateView THEN%>checked<%END IF%>> 노출안함
					   		</td>
					   	</tr>
					   	<tr>
					   		<td align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">대표상품정보<br/>및<br/>배너</td>
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
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2011 기본배너</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBan" value="2011 배너이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=ebimg%>','ban','spanban')" class="button">
					   			<div id="spanban" style="padding: 5 5 5 5">
					   				<%IF ebimg <> "" THEN %>
					   				<img  src="<%=ebimg%>" border="0">
					   				<a href="javascript:jsDelImg('ban','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2010 기본배너</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBan2010" value="2010 배너이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=ebimg2010%>','ban2010','spanban2010')" class="button">
					   			<div id="spanban2010" style="padding: 5 5 5 5">
					   				<%IF ebimg2010 <> "" THEN %>
					   				<img  src="<%=ebimg2010%>" border="0">
					   				<a href="javascript:jsDelImg('ban2010','spanban2010');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
						<!-- <tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">모바일 리스트배너(앱 업데이트시 삭제 예정)</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBanMo" value="모바일 리스트배너 등록" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo%>','banMo','spanbanMo')" class="button">
					   			<div id="spanbanMo" style="padding: 5 5 5 5">
					   				<%IF ebimgMo <> "" THEN %>
					   				<img  src="<%=ebimgMo%>" border="0">
					   				<a href="javascript:jsDelImg('banMo','spanbanMo');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   			<p style="color:#602030;font-size:11px;">※ 권장 이미지 : JPEG, 50%, 560px × 380px</p>
					   		</td>
					   	</tr> -->
						<!-- <tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">모바일 Today배너(앱 업데이트시 삭제 예정)</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnBanToday" value="모바일 Today-enjoyevent배너 등록" onClick="jsSetImg('<%=eFolder%>','<%=ebimgToday%>','banMoToday','spanbanMoToday')" class="button">
					   			<div id="spanbanMoToday" style="padding: 5 5 5 5">
					   				<%IF ebimgToday <> "" THEN %>
					   				<img  src="<%=ebimgToday%>" border="0">
					   				<a href="javascript:jsDelImg('banMoToday','spanbanMoToday');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
								<p style="color:#602030;font-size:11px;">※ 권장 이미지 : JPEG, 50%, 600px × 270px</p>
					   		</td>
					   	</tr> -->
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">2014 모바일<br/>리스트배너</td>
					   		<td bgcolor="#FFFFFF">
					   		<input type="button" name="btnMoBan2014" value="2014 모바일 리스트 배너" onClick="jsSetImg('<%=eFolder%>','<%=ebimgMo2014%>','banMoList','spanbanMoList')" class="button">
					   			<div id="spanbanMoList" style="padding: 5 5 5 5">
					   				<%IF ebimgMo2014 <> "" THEN %>
					   				<img  src="<%=ebimgMo2014%>" border="0">
					   				<a href="javascript:jsDelImg('banMoList','spanbanMoList');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
								<p style="color:#602030;font-size:11px;">※ 권장 이미지 : JPEG, 50%, 640px × 340px</p>
					   		</td>
					   	</tr>
						<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">기본아이콘 </td>
					   		<td bgcolor="#FFFFFF">
					   			<input type="button" name="btnicon" value="리스트아이콘 등록" onClick="jsSetImg('<%=eFolder%>','<%=eicon%>','icon','spanicon')" class="button">
					   			<div id="spanicon" style="padding: 5 5 5 5">
					   				<%IF eicon <> "" THEN %>
					   				<img  src="<%=eicon%>">
					   				<a href="javascript:jsDelImg('icon','spanicon');"><img src="/images/icon_delete2.gif" border="0"></a>
					   				<%END IF%>
					   			</div>
					   		</td>
					   	</tr>
					   	<tr>
					   		<td width="100" align="center" colspan="2" bgcolor="<%= adminColor("tabletop") %>">사은품 이미지 </td>
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
					   	<tr>
					   		<td width="30" align="center" rowspan="2"  bgcolor="<%= adminColor("tabletop") %>"> 이<br>벤<br>트<br><br>상<br>세<br>페<br>이<br>지<br> </td>
					   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">화면템플릿</td>
					   		<td bgcolor="#FFFFFF"><%sbGetOptEventCodeValue "eventview",etemp,false,"onchange=""jsChangeFrm(this.value);"""%></td>
					   	</tr>
					   	<tr>
					   		<td width="65" align="center"  bgcolor="<%= adminColor("tabletop") %>">이미지<br>&<br>HTML</td>
					   		<td bgcolor="#FFFFFF">
					   			<!-- 1.메인 탑-->
					   			<div id="divFrm1" style="display:;">
						   			<input type="button" name="btnMain" value="메인TOP이미지 등록" onClick="jsSetImg('<%=eFolder%>','<%=emimg%>','main','spanmain')" class="button">
						   			<div id="spanmain" style="padding: 5 5 5 5">
						   				<%IF emimg <> "" THEN %>
						   				<a href="javascript:jsImgView('<%=emimg%>')"><img  src="<%=emimg%>" width="400" border="0"></a>
						   				<a href="javascript:jsDelImg('main','spanmain');"><img src="/images/icon_delete2.gif" border="0"></a>
						   				<%END IF%>
						   			</div>
					   				<hr>
									<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('1');">PC-WEB 예시</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('2');">Mobile-WEB 예시</span>||<span style="color:red;font-weight:800;cursor:pointer;" onclick="chgex('3');">APP 예시</span>
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
									<div id="notice2" style="display:none">
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
										&lt;a href="#" onclick="fnAPPpopupCategory('<span style="color:red">카테고리번호</span>'); return false;"&gt; 카테고리 링크 &lt;/a&gt;<br><br>
										※앱 웹 구분 스크립트※<br/>
										&lt;script&gt;<br/>
										var chkapp = navigator.userAgent.match('tenapp');<br/>
										if ( chkapp ){<br/>
										&nbsp;&nbsp;&nbsp;//앱영역 스크립트<br/>
										}else{<br/>
										&nbsp;&nbsp;&nbsp;//모바일영역 스크립트<br/>
										}<br/>
										&lt;/script&gt;<br/>
									</div>
									<br>
									<b>이미지 경로 http://<font color="RED">webimage.</font>10x10.co.kr/event/XXX/</b> 로 변경되었습니다.<br>
						   			<textarea name="tHtml" rows="20" style="width:100%;font-size:11px;"><%=ehtml%></textarea>
					   			</div>
					   			<!-- 3.그룹형-->
					   			<div id="divFrm3" style="display:none;">
					   				<iframe id="iframG" src="about:blank" frameborder="0" width="100%" class="autoheight"></iframe>
					   			</div>
					   			<!-- /3.그룹형-->
					   				<!-- 5.수작업-->
					   			<div id="divFrm5" style="display:none;">
					   				<table border="0" cellpadding="1" cellspacing="3" class="a">
					   					<tr>
					   						<td>
					   						    <!-- <input type="button" value="이미지관리"  onclick="TnFtpUpload('D:/home/cube1010/imgstatic/event/<%= eFolder%>/','/event/<%= eFolder%>/');" class="input_b"> -->
					   						    <input type="button" value="이미지관리(신)"  onclick="jsManageEventImageNew('<%= eFolder%>')" class="input_b">
					   						    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					   						    <input type="button" value="이미지관리(구)"  onclick="jsManageEventImage('<%= eFolder%>')" class="input_b">
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
					</table>
				</td>
			</tr>
		</table>
	</div>
	</td>

</tr>
<tr>
	<td width="100%" height="40" align="right">
		<input type="image" src="/images/icon_save.gif">
		<a href="index.asp?<%=strparm%>"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</table>
</form>
 <script language="javascript">
<!--
jsChangeFrm('<%=etemp%>');
//-->
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
