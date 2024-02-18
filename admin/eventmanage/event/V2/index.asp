<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/index.asp
' Description :  이벤트 등록 - 화면설정
' History : 2007.02.07 정윤정 생성
'           2012.02.13 허진원 - 미니달력 교체
'			2014.03.10 정윤정 - 관심항목 최이령(fotoark), 이주경(arlejk) 예외사항 설정
'           2015.03 정윤정 - 이벤트 리뉴얼
'           2017.04.14 허진원 - 서브 디자이너 추가
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
	response.write "<script type='text/javascript'>"
	response.write "	alert('사용불가 페이지');history.back();"
	response.write "</script>"
	response.End
	Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

	'변수선언
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory, sCateMid ,sState,sKind,esale,egift,ecoupon,ebrand,eonlyten,etype,isConfirm
	Dim strparm
	Dim edgid, edgid2,edgstat1,edgstat2, emdid, epsid, edpid, edgnm, edgnm2, emdnm, epsnm, edpnm, eDiary
	dim eopo,efd,ebs,enew
	dim blnWeb, blnMobile, blnApp
	dim dispCate, maxDepth
	dim blnReqPublish ,sSort
	dim isResearch

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	maxDepth = 2
	
	isResearch = requestCheckVar(Request("isResearch"),1)
	if isResearch ="" then isResearch ="0"
	'## 검색 #############################
	sDate 		= requestCheckVar(Request("selDate"),1)  	'기간
	sSdate 		= requestCheckVar(Request("iSD"),10)
	sEdate 		= requestCheckVar(Request("iED"),10)

	sEvt 		= requestCheckVar(Request("selEvt"),10)  	'이벤트 코드/명 검색
	strTxt 		= requestCheckVar(Request("sEtxt"),60)

	sCategory	= requestCheckVar(Request("selC"),10) 		'카테고리
	sCateMid	= requestCheckVar(Request("selCM"),10) 		'카테고리(중분류)
	dispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리
	sState		= requestCheckVar(Request("eventstate"),4)	'이벤트 상태
	 
	sKind 		= requestCheckVar(Request("eventkind"),32)	'이벤트종류
	edgid  		= requestCheckVar(Request("sDgId"),32)		'담당 디자이너
''	edgid2 		= requestCheckVar(Request("sDg2Id"),32)		'서브 디자이너
	emdid  		= requestCheckVar(Request("sMdId"),32)		'담당 MD
	epsid  		= requestCheckVar(Request("sPsId"),32)		'담당 퍼블리셔
	edpid  		= requestCheckVar(Request("sDpId"),32)		'담당 개발자
	
	edgnm  		= requestCheckVar(Request("sdgnm"),32)		'담당 디자이너
''	edgnm2 		= requestCheckVar(Request("sdg2nm"),32)		'서브 디자이너
	emdnm  		= requestCheckVar(Request("smdnm"),32)		'담당 MD
	epsnm  		= requestCheckVar(Request("spsnm"),32)		'담당 퍼블리셔
	edpnm  		= requestCheckVar(Request("sdpnm"),32)		'담당 개발자

	if Request("designerstatus")<>"" AND Request("designerstatus") <> "," then
		edgstat1	= requestCheckVar(Request("designerstatus")(1),2)		'담당 디자이너 상태
		edgstat2	= requestCheckVar(Request("designerstatus")(2),2)		'서브 디자이너 상태
	end if

	ebrand		= requestCheckVar(Request("ebrand"),32)		'브랜드
	esale		= requestCheckVar(Request("chSale"),2) 		'세일유무
	egift		= requestCheckVar(Request("chGift"),2)		'사은품유무
	ecoupon	 	= requestCheckVar(Request("chCoupon"),2)	'쿠폰유무
	eonlyten	= requestCheckVar(Request("chOnlyTen"),2)	'Only-TenByTen유무
	eDiary		= requestCheckVar(Request("chDiary"),2)	'다이어리 유무
	eopo		= requestCheckVar(Request("chopo"),1)	'원플러스원
	efd		= requestCheckVar(Request("chfd"),1)	'무료배송
	ebs		= requestCheckVar(Request("chbs"),1)	'예약판매
	enew		= requestCheckVar(Request("chnew"),1)	'new
	
	blnWeb		= requestCheckVar(Request("isWeb"),1)
	blnMobile	= requestCheckVar(Request("isMobile"),1)
	blnApp		= requestCheckVar(Request("isApp"),1)
	
	dispCate 	= requestCheckvar(request("disp"),16)
	blnReqPublish= requestCheckvar(request("chkPus"),1)
	sSort       = requestCheckvar(request("sSort"),2)

	etype		= requestCheckvar(request("eventtype"),4)
	isConfirm	= requestCheckvar(request("blnCnfm"),1)
	
	if isResearch="0" and sKind="" then
		skind="1,12,13,23,27,28,29,31"
	end if


	'이벤트 첫페이지 관심항목이 보이도록 
	IF (sKind="" and isResearch="0") or sKind="1,12" THEN
		if (session("ssAdminPsn")="11" or session("ssAdminPsn")="21") and (not ( session("ssBctId")="fotoark" or session("ssBctId")="arlejk" or session("ssBctId")="barbie8711")) then
			'MD부서라면 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너) - 최이령(fotoark), 이주경(arlejk), 차선화(barbie8711) 제외
			sKind = "1,12,13,16,17,23,24"
		else
			'기타 (쇼핑찬스,전체,상품,브랜드,다이어리,테스터,신규디자이너,모바일,브랜드Week)
			sKind = "1,12,13,16,17,23,24,19,25,26,31"
		end if
	end if
	strparm  = "isWeb="&blnWeb&"&isMobile="&blnMobile&"&isApp="&blnApp&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&selCM="&sCateMid&"&eventstate="&sState&"&eventkind="&sKind&"&sDgId="&edgid&"&sMdId="&emdid&"&spsid="&epsid&"&sdpid="&edpid&_
				"&sdgnm="&edgnm&"&smdnm="&emdnm&"&spsnm="&epsnm&"&sdpnm="&edpnm&"&ebrand="&ebrand&"&chSale="&esale&"&chGift="&egift&"&chCoupon="&ecoupon&"&chOnlyTen="&eonlyten&"&disp="&dispCate&"&chDiary="&eDiary&"&sDg2Id="&edgid2&"&sdg2nm="&edgnm2&"&designerstatus="&edgstat1&"&designerstatus="&edgstat2
	'#######################################
 	if sSort = "" then sSort = "CD"
 	if blnReqPublish= "" then blnReqPublish = False     
 	    
	'데이터 가져오기
	set cEvtList = new ClsEvent
		cEvtList.FCPage = iCurrpage		'현재페이지
		cEvtList.FPSize = iPageSize		'한페이지에 보이는 레코드갯수

		cEvtList.FSfDate 	= sDate		'기간 검색 기준
		cEvtList.FSsDate 	= sSdate	'검색 시작일
		cEvtList.FSeDate 	= sEdate	'검색 종료일
		cEvtList.FSfEvt 	= sEvt		'검색 이벤트명 or 이벤트코드
		cEvtList.FSeTxt 	= strTxt	'검색어
		cEvtList.FScategory = sCategory	'검색 카테고리
		cEvtList.FScateMid	= sCateMid	'검색 카테고리(중분류)
		cEvtList.FEDispCate	= dispCate	'검색 전시카테고리
		cEvtList.FSstate 	= sState	'검색 상태
	 
		cEvtList.FSedid   	= edgid
'		cEvtList.FSedid2   	= edgid2
		cEvtList.FSemid   	= emdid
		cEvtList.FSepsid   	= epsid
		cEvtList.FSedpid   	= edpid
		
		cEvtList.FSednm   	= edgnm
		cEvtList.FSemnm   	= emdnm
		cEvtList.FSepsnm   	= epsnm
		cEvtList.FSedpnm   	= edpnm
		
		cEvtList.FEDgStat1	= edgstat1
		cEvtList.FEDgStat2	= edgstat2
		
		
		cEvtList.FSkind 	= sKind
		cEvtList.FEBrand 	= ebrand
		cEvtList.FSisSale 	= esale
		cEvtList.FSisGift 	= egift
		cEvtList.FSisCoupon	= ecoupon
		cEvtList.FSisOnlyTen= eonlyten
		cEvtList.FSisDiary = eDiary
		cEvtList.FSisoneplusone   = eopo
		cEvtList.FSisfreedelivery = efd
		cEvtList.FSisbookingsell  = ebs
		cEvtList.FSisNew          = enew
		
		cEvtList.FIsWeb = blnWeb
		cEvtList.FIsMobile = blnMobile
		cEvtList.FIsApp = blnApp
		
		cEvtList.FRectEvtType = etype
		cEvtList.FRectIsConfirm = isConfirm
		
		cEvtList.FIsReqPublish = blnReqPublish
		cEvtList.FSort          = sSort
 		arrList = cEvtList.fnGetEventList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotCnt	'전체 데이터  수
 	set cEvtList = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	Dim arreventlevel, arreventstate, arreventkind, arreventtype, arrdsnStat
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventlevel = fnSetCommonCodeArr("eventlevel",False)
	arreventstate= fnSetCommonCodeArr("eventstate",False)
	arreventkind= fnSetCommonCodeArr("eventkind",False)
	arreventtype= fnSetCommonCodeArr("eventtype",False)
	arrdsnStat = fnSetCommonCodeArr("designerstatus",False)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!-- 
	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}

	function jsSearch(sType){
		var frm = document.frmEvt
		if (sType == "A"){
				frm.iSD.value = "";
				frm.iED.value = "";
				frm.eventstate.value = "";
				frm.sEtxt.value = "";
				frm.selC.value = "";
		}

		if(frm.selEvt.value== "evt_code"&&frm.sEtxt.value!=""){
			frm.sEtxt.value = frm.sEtxt.value.replace(/\s/g, "");
			if(!IsDigit(frm.sEtxt.value)){
				alert("이벤트코드는 숫자만 가능합니다.");
				frm.sEtxt.focus();
				return;
			}
		}
 
		frm.submit();
	}

	function SubmitForm() {
		jsSearch('E');
	}

	function jsSchedule(){
		var winS;
		winS = window.open('/admin/eventmanage/event/pop_event_schedule.asp','popwin','width=1200, height=800, scrollbars=yes');
		winS.focus();
	}




	function jsCodeManage(){
		var winCode;
		winCode = window.open('/admin/eventmanage/code/popManageCode.asp','popCode','width=450,height=600,scrollbars=yes,resizable=yes');
		winCode.focus();
	}

	function prize(evt_code){

		 var prize = window.open('/admin/eventmanage/event/pop_event_prize.asp?evt_code='+evt_code,'prize','width=800,height=600,scrollbars=yes,resizable=yes');
		 prize.focus();

	}
	
	function jsGetID(sType, iCid, sUserID){
		var openWorker = window.open('PopWorkerList.asp?sType='+sType+'&department_id='+iCid+'&sUserid='+sUserID,'openWorker','width=350,height=570,scrollbars=yes');
		openWorker.focus();
	}
	
	function jsDelID(sType){ 
		eval("document.frmEvt.s"+sType+"Id").value = "";
		eval("document.frmEvt.s"+sType+"Nm").value = ""; 
	}
	
	 //리스트 정렬
	 function jsSort(sValue,i){  
	  
	 	document.frmEvt.sSort.value= sValue; 
	 	 
		   if (-1 < eval("document.all.img"+i).src.indexOf("_alpha")){
	        document.frmEvt.sSort.value= sValue+"D";  
	    }else if (-1 < eval("document.all.img"+i).src.indexOf("_bot")){
	     		document.frmEvt.sSort.value= sValue+"A";  
	    }else{
	       document.frmEvt.sSort.value= sValue+"D";  
	    } 
	    
	   
		 document.frmEvt.submit();
	}
	
	//날짜 지정
	function jsSetDate(iValue){
	    var currentDate = new Date(); 

        var month = currentDate.getMonth() + 1;
        var day = currentDate.getDate();
        var year = currentDate.getFullYear();
 
        var preDate = new Date(currentDate.setMonth(month-iValue)); 
        var pmonth = preDate.getMonth() ; 
        var pday = preDate.getDate();
        var pyear = preDate.getFullYear();
        
        if (month <10){
            month = "0"+month;
        }
        
         if (pmonth <10){
            pmonth = "0"+pmonth;
        }
        
         if (pday <10){
            pday = "0"+pday;
        }
        
         if (day <10){
            day = "0"+day;
        }
        
 	    document.frmEvt.iSD.value = pyear+"-"+pmonth+"-"+pday; 
	    document.frmEvt.iED.value = year+"-"+month+"-"+day;
	}
	
	//
	function makeThumbBanTxt(evtcode,oldview){
	    var popwinThumbTxt = window.open('http://110.93.128.113/pSvr/makeCateEvtBanner.asp?evtcode='+evtcode+'&oldview='+oldview,'popwinThumbTxt','width=680,height=570,scrollbars=yes');
		popwinThumbTxt.focus();
	}
	
	//미리보기
	function jsOpen(sPURL,sTG){ 
	    if (sTG =="M" ){ 
	        var winView = window.open(sPURL,"popView","width=400, height=600,scrollbars=yes,resizable=yes,location=yes");
	    }
	}
//-->
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style>
	select {font-size:12px; vertical-align:top;}
	input[type=button], input[type=text] {vertical-align:top;}
</style>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  action="index.asp" onSubmit="return jsSearch('E');">
	<input type="hidden" name="menupos" value="<%=menupos%>"> 
	<input type="hidden" name="isResearch" value="1"> 
	<input type="hidden" name="sSort" value="<%=sSort%>">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>브랜드:</td>
				<td><% drawSelectBoxDesignerwithName "ebrand", ebrand %></td>
				<td colspan="4">
					관리 <!-- #include virtual="/common/module/categoryselectbox_event.asp"-->
					/ 전시 카테고리 : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
				</td>
			</tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">이벤트종류:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><%sbGetOptCommonCodeArr "eventkind", sKind, True,True,"onChange='javascript:document.frmEvt.submit();'"%></td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">코드/명:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selEvt">
			    	<option value="evt_code" <%if Cstr(sEvt) = "evt_code" THEN %>selected<%END IF%>>이벤트코드</option>
			    	<option value="evt_name" <%if Cstr(sEvt) = "evt_name" THEN %>selected<%END IF%>>이벤트명</option>
			    	<option value="evt_tag" <%if Cstr(sEvt) = "evt_tag" THEN %>selected<%END IF%>>TAG</option>
			    	<option value="evt_sub" <%if Cstr(sEvt) = "evt_sub" THEN %>selected<%END IF%>>서브카피</option>
			    	</select>
			        <input type="text" name="sEtxt" value="<%=strTxt%>" maxlength="60" onkeydown="if(event.keyCode==13) document.frmEvt.submit();" /></td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">이벤트타입:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" >
			    	<input type="checkbox" name="chSale" <%IF Cstr(esale)="1" THEN%> checked <%END IF%>  value="1">할인
					<input type="checkbox" name="chGift" <%IF Cstr(egift)="1" THEN%> checked<%END IF%>  value="1">사은품
					<input type="checkbox" name="chCoupon" <%IF Cstr(ecoupon)="1" THEN%> checked<%END IF%> value="1">쿠폰 
					<input type="checkbox" name="chOnlyTen" <%IF Cstr(eonlyten)="1" THEN%> checked<%END IF%> value="1">Only-TenByTen 
					<input type="checkbox" name="chopo" <%IF Cstr(eopo)="1" THEN%> checked<%END IF%> value="1">1+1  
			    </td> 
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">이벤트유형:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
						<%sbGetOptCommonCodeArr "eventtype", eType, True,True,"onChange='javascript:document.frmEvt.submit();'"%> &nbsp; 
						<input type="checkbox" name="blnCnfm" <%=chkIIF(Cstr(isConfirm)="1","checked","")%> value="1">승인완료
				</td>
				<td colspan="3" style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"> 
					<input type="checkbox" name="chfd" <%IF Cstr(efd)="1" THEN%> checked<%END IF%> value="1">무료배송
					<input type="checkbox" name="chbs" <%IF Cstr(ebs)="1" THEN%> checked<%END IF%> value="1">예약판매 
					<input type="checkbox" name="chDiary" <%IF Cstr(eDiary)="1" THEN%> checked<%END IF%> value="1">DiaryStory
					<input type="checkbox" name="chnew" <%IF Cstr(enew)="1" THEN%> checked<%END IF%> value="1">런칭
				</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">진행상태:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
				    <% Dim arrCode 
                     	arrCode= fnSetCommonCodeArr("eventstate", False)
 	
 	                 %>
				    <select name="eventstate" class="select" onChange="javascript:SubmitForm();">
				        <option value="">선택</option>
				        <% 	IF isArray(arrCode) THEN
                         	For intLoop =0 To UBound(arrCode,2)
                         	    if arrCode(0,intLoop) = 1 THEN
                         	 
                        %>
                            <option value="1^3" <%If CStr(sState) = "1^3" THEN%>selected<%END IF%>>이미지등록요청+디자이너작업중</option>
                        <%      elseif arrCode(0,intLoop) = 9 THEN%>    
                            <option value="6^9" <%If CStr(sState) = "6^9" THEN%>selected<%END IF%>>오픈+종료</option>
                        <%      end if%>
                        	<option value="<%=arrCode(0,intLoop)%>" <%If CStr(sState) = CStr(arrCode(0,intLoop)) THEN%>selected<%END IF%> <%if arrCode(2,intLoop) ="N" then%>style="color:gray;"<%end if%>> <%=arrCode(1,intLoop)%></option>
                        <%
                        	Next
                        	End IF
                        	%>
				     
				    </select>  
				  </td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">기간:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;"><select name="selDate">
			    	<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>시작일 기준</option>
			    	<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>종료일 기준</option>
			    	<option value="O" <%if Cstr(sDate) = "O" THEN %>selected<%END IF%>>오픈일 기준</option>
			        </select>
			         <input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			        <input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "iSD", trigger    : "iSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "iED", trigger    : "iED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					
			        <!--<input type="text" id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" style="vertical-align:top" />
			        <!-- <input type="image" src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" style="vertical-align:top" /> ~
			        <input type="text" id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" style="vertical-align:top" /> <input type="image" src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" style="vertical-align:top" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "iSD", trigger    : "iSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "iED", trigger    : "iED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>-->
					&nbsp;
					<input type="button" value="최근 1달" class="button" onClick="jsSetDate(1)">
					<input type="button" value="최근 3달" class="button" onClick="jsSetDate(3)">
			    </td>
				<td colspan="2" style="border-top:1px solid <%= adminColor("tablebg") %>;">&nbsp;</td>
			</tr>
			<tr>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">디자인상태:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<span style="white-space:nowrap; font-family:'Lucida Console';">PC <%sbGetOptEventCodeValue "designerstatus",edgstat1,True,""%></span> &nbsp;
					<span style="white-space:nowrap; font-family:'Lucida Console';">MW <%sbGetOptEventCodeValue "designerstatus",edgstat2,True,""%></span>
				</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;">담당자:</td>
			    <td style="border-top:1px solid <%= adminColor("tablebg") %>;" colspan="3">
			    	<span style="white-space:nowrap;">기획자  <input type="hidden" name="sMdId" value="<%=emdid%>"><input type="name" name="sMdNm" value="<%=eMDnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="선택" onClick="jsGetID('Md','162','<%=emdid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Md');" title="담당자 지우기" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">디자이너 <input type="hidden" name="sDgId" value="<%=edgid%>"><input type="name" name="sDgNm" value="<%=edgnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="선택" onClick="jsGetID('Dg','152','<%=edgid%>');">&nbsp;<input type="button" value="&times"  class="button" onClick="jsDelID('Dg');" title="담당자 지우기" /></span> &nbsp;
			    	<span style="white-space:nowrap;">퍼블리셔 <input type="hidden" name="sPsId" value="<%=epsid%>"><input type="name" name="sPsNm" value="<%=epsnm%>"class="text"  size="10">&nbsp;<input type="button" class="button" value="선택"  onClick="jsGetID('Ps','157','<%=epsid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Ps');" title="담당자 지우기" /></span> &nbsp;&nbsp;
			    	<span style="white-space:nowrap;">개발자  <input type="hidden" name="sDpId" value="<%=edpid%>"><input type="name" name="sDpNm" value="<%=edpnm%>" class="text"  size="10">&nbsp;<input type="button" class="button" value="선택" onClick="jsGetID('Dp','130','<%=edpid%>');"> <input type="button" value="&times"  class="button" onClick="jsDelID('Dp');" title="담당자 지우기" /></span>
			    </td>
			</tr>
			<tr>
			<tr> 
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">채널:</td>
				<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
					<input type="checkbox" name="isWeb" value="1" <%if blnWeb="1" then%>checked<%end if%>> PC-Web
					<input type="checkbox" name="isMobile"  value="1" <%if blnMobile="1" then%>checked<%end if%>> Mobile
					<input type="checkbox" name="isApp"  value="1" <%if blnApp="1" then%>checked<%end if%>> App
				</td>
			    <tD colspan="4" style="border-top:1px solid <%= adminColor("tablebg") %>;"><input type="checkbox" name="chkPus" value="1" <%if blnReqPublish THEN%>checked<%end if%>> 퍼블리싱 요청작업</td>
			</tr>
			</table>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch('E');">
		</td>
	</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
        <td align="left">
        	<input type="button" value="새로등록" onclick="jsGoUrl('event_regist.asp?menupos=<%=menupos%>&<%=strParm%>');" class="button">
	    </td>
	    <td align="right">
	       	<input type="button" value="스케쥴" onclick="jsSchedule();"  class="button">
	       <!--	<input type="button" value="통계" onclick=" ">  -->
	       <% if C_ADMIN_AUTH then %><input type="button" value="코드관리" onclick="jsCodeManage();"  class="button"><%END IF%>
        </td>
	</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="20">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap rowspan="2">채널</td>
    	<td nowrap rowspan="2">이벤트종류</td>
    	<td nowrap rowspan="2">이벤트유형</td>
    	<td nowrap rowspan="2" onClick="javascript:jsSort('C','1');" style="cursor:hand;"><b>이벤트코드</b><img src="/images/list_lineup<%IF sSort="CD" THEN%>_bot<%ELSEIF sSort="CA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
    	<td nowrap rowspan="2" onClick="javascript:jsSort('S','2');" style="cursor:hand;"><b>중요도</b> <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot<%ELSEIF sSort="SA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
      	<td nowrap rowspan="2">진행상태</td>
      	<td nowrap rowspan="2">배너</td>
      	<td nowrap rowspan="2">와이드배너</td>
      	<td nowrap rowspan="2">이벤트명</td>
      	<td nowrap rowspan="2">카테고리</td>
      	<td nowrap rowspan="2">브랜드</td>
      	<td width="60" rowspan="2" onClick="javascript:jsSort('D','3');" style="cursor:hand;"><b>시작일</b> <img src="/images/list_lineup<%IF sSort="DD" THEN%>_bot<%ELSEIF sSort="DA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
      	<td width="60" rowspan="2">종료일</td>
      	<td rowspan="2"  onClick="javascript:jsSort('I','4');" style="cursor:hand;"><b>이미지요청일</b> <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
      	<td nowrap colspan="5">담당자</td>
      	<td nowrap rowspan="2">관리</td>
     </tr>
     <tr align="center" bgcolor="<%= adminColor("tabletop") %>">	 
        <td nowrap>기획자</td>
      	<td colspan="2" nowrap>디자이너</td>
      	<td nowrap>퍼블리셔</td>
      	<td nowrap>개발자<br />/ 검수자</td>
    </tr>

    <%IF isArray(arrList) THEN
		Dim itemSortvalue
		Dim strURL
		Dim isMobile, isApp, isWeb
		 dim tmpename, ename,eSalePer
		
    	For intLoop = 0 To UBound(arrList,2) 
		
		'2014-08-27 김진영 / 변수에 순서값 저장
		Select Case arrList(27,intLoop)
			Case "1"	itemSortvalue = "sitemid"
			Case "2"	itemSortvalue = "slsell"
			Case "3"	itemSortvalue = "sevtitem"
			Case "4"	itemSortvalue = "sbest"
			Case "5"	itemSortvalue = "shsell"
		End Select
		
		isWeb = False
		isMobile = False
		isApp = False
		
		IF isNull(arrList(30,intLoop)) and isNull(arrList(31,intLoop)) and isNull(arrList(32,intLoop)) then
			if arrList(1,intLoop) = "19" THEN
				isWeb = False
				isMobile = True
				isApp = True
			ELSEIF arrList(1,intLoop) = "25"  THEN
				isWeb = False
				isMobile = False
				isApp = True
			ELSEIF arrList(1,intLoop) = "26"  THEN	
				isWeb = False
				isMobile = True
				isApp = False
			ELSE
				isWeb = True
				isMobile = False
				isApp = False	
			END IF
		END IF	
		IF 	 not isNull(arrList(30,intLoop))  THEN	
			isWeb = arrList(30,intLoop)
		END IF	
		IF 	 not isNull(arrList(31,intLoop)) THEN
			 isMobile = arrList(31,intLoop)
		END IF	 
		IF 	 not isNull(arrList(32,intLoop)) THEN
			isApp = arrList(32,intLoop)	
		END IF	
		
		 
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>
    	    <% 	
			dim sMoblie,sWeb,sApp
			 sApp = ""
			 sMoblie=""
			Select Case arrList(1,intLoop)
				Case "7"		'위클리코디
					 sWeb = "<a href='" & vwwwUrl & "/guidebook/weekly_coordinator.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
				Case "13"		'상품 이벤트
					sWeb =  "<a href='" & vwwwUrl & "/shopping/category_prd.asp?itemid=" & arrList(21,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/category/category_itemPrd.asp?itemid=" & arrList(21,intLoop) & "','M');"">" 
				Case "14"		'소풍가는길
					sWeb =  "<a href='" & vwwwUrl & "/guidebook/picnic/picnic.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
'				Case "15"		'브랜드데이
'					sWeb =  "<a href='" & vwwwUrl & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
'					sMoblie =  "<a href='" & mobileUrl & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
				Case "16"		'브랜드 할인행사
					sWeb =  "<a href='" & vwwwUrl & "/street/street_brand_sub06.asp?makerid=" & arrList(14,intLoop) & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>"  
				Case "22"		'DAY&(데이앤드)
					sWeb = "<a href='" & vwwwUrl & "/guidebook/dayand.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"   
				Case "26"		'모바일
					sWeb =  "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
				Case "29"		'헤이썸띵
					sWeb =  "<a href='" & vwwwUrl & "/HSProject/?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie =  "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
				Case Else		'쇼핑찬스 및 기타
					sWeb = "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>"  
					sMoblie = "<a href= ""javascript:jsOpen('" & vmobileUrl & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"  
					sApp ="<a href= ""javascript:jsOpen('" & vmobileUrl & "/apps/appCom/wish/web2014/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "','M');"">"
			End Select 
		%>
    	    <%IF isWeb THEN %> <%=sWeb%>Web</a><%END IF%>
    	    <%=chkIIF(isMobile,"<br />" & sMoblie & "<font color=""blue"">Mobile</font></a>","")%>
    	    <%=chkIIF(isApp,"<br />" & sApp & "<font color=""red"">App</font></a>","")%>
    	</td>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventkind,arrList(1,intLoop))%></a></a></td>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventtype,arrList(39,intLoop))%></a></a></td>
		<% 	
			
			'이벤트 종류에 따른 프론트링크 페이지 선택
			IF isMobile or isApp  THEN '모바일/앱일때..
				strURL = vmobileUrl
			ELSE	'기타..
				strURL = vwwwUrl
			END IF	
			Select Case arrList(1,intLoop)
				Case "7"		'위클리코디
					Response.Write "<td><a href='" & strURL & "/guidebook/weekly_coordinator.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "13"		'상품 이벤트
					Response.Write "<td><a href='" & strURL & "/shopping/category_prd.asp?itemid=" & arrList(21,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "14"		'소풍가는길
					Response.Write "<td><a href='" & strURL & "/guidebook/picnic/picnic.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
'				Case "15"		'브랜드데이
'					Response.Write "<td><a href='" & strURL & "/street/street_brandday.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "16"		'브랜드 할인행사
					Response.Write "<td><a href='" & strURL & "/street/street_brand_sub06.asp?makerid=" & arrList(14,intLoop) & "&shop_event_yn=Y&shop_event_confirm=Y&shopview=3' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "22"		'DAY&(데이앤드)
					Response.Write "<td><a href='" & strURL & "/guidebook/dayand.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case "26"		'모바일
					Response.Write "<td><a href='" & strURL & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
				Case Else		'쇼핑찬스 및 기타
					Response.Write "<td><a href='" & strURL & "/event/eventmain.asp?eventid=" & arrList(0,intLoop) & "' target='_blank'>" & arrList(0,intLoop) & "</a></td>"
			End Select 
		%>
    	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventlevel,arrList(7,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=fnGetCommCodeArrDesc(arreventstate,arrList(8,intLoop))%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(34,intLoop) <> "" THEN%> <img src="<%=arrList(34,intLoop)%>" width="100" border="0"><%END IF%></a></td>
        <td>
            <a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%IF arrList(35,intLoop) <> "" THEN%> <img src="<%=arrList(35,intLoop)%>" width="100" border="0"><%END IF%></a>
            <!-- // 2017.04.18 사용하지 않음으로 제거 (모바일 텍스트 이미지 배너)
            <%IF arrList(37,intLoop) <> "" THEN%> <br><img src="<%=arrList(37,intLoop)%>" width="100" border="0" onclick="makeThumbBanTxt('<%=arrList(0,intLoop)%>','on')"><%ELSE%><%IF arrList(35,intLoop) <> "" THEN%><br><input type="button" value="생성" onclick="makeThumbBanTxt('<%=arrList(0,intLoop)%>','')"><%END IF%><%END IF%>
            //-->
        </td>  
      	<td align="left">
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=chkIIF(Not(arrList(25,intLoop)="" or isNull(arrList(25,intLoop))),"["&arrList(25,intLoop)&"] ","")%>
      		<%   ename =  arrList(4,intLoop)  
      		     eSalePer = ""
      		    if  (arrList(15,intLoop) or arrList(17,intLoop)) then 
            	    tmpename = Split(ename,"|")  
            	  	if Ubound(tmpename)>0 then
            		    ename = tmpename(0)
            		    eSalePer = tmpename(1)
            		 end if
            
                end if
             %>   
      		<%=db2html(ename)%>
      		<% if arrList(15,intLoop)  then%>
      		<font color="red"><%=eSalePer%></font>
      		<% elseif arrList(17,intLoop) then%>
      		<font color="green"><%=eSalePer%></font>
      		<% end if%>
      		<%if arrList(15,intLoop)  then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif" border="0"><%end if%>
      		<%if arrList(16,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif" border="0"><%end if%>
      		<%if arrList(17,intLoop) then%>&nbsp;<img src="http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif" border="0"><%end if%>
      		</a>
      	</td>
      	<td>
      		<a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>">
      		<%=arrList(12,intLoop)%>
      		<%
      		if arrList(22,intLoop) <> "" then
      			response.write "(" & arrList(22,intLoop) &")"
      		end if
      		'전시카테고리
      		if arrList(26,intLoop)<>"" then
      			response.write chkIIF(arrList(12,intLoop)<>"","<br/>","") & "<font color='#4030A0'>" & arrList(26,intLoop) & "</font>"
      		end if
      		%>
      		</a>
      	</td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(14,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(5,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(6,intLoop)%></a></td> 
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(36,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(23,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(11,intLoop)%><br /><span style="color:#B88;"><%=fnGetCommCodeArrDesc(arrdsnStat,arrList(42,intLoop))%></span></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(41,intLoop)%><br /><span style="color:#B88;"><%=fnGetCommCodeArrDesc(arrdsnStat,arrList(43,intLoop))%></span></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(28,intLoop)%></a></td>
      	<td><a href="event_modify.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>"><%=arrList(29,intLoop)%><%=chkiif(arrList(38,intLoop)<>"","<br />" & arrList(38,intLoop),"")%></a></td>
      	<td align="left" nowrap><input type="button" value="상품" class="button" onClick="javascript:jsGoUrl('eventitem_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>&selsort=<%=itemSortvalue%>')">
      		<%IF arrList(13,intLoop) > "1900-01-01" THEN%><input type="button" value="당첨" class="button" onClick="jsGoUrl('/admin/eventmanage/event/eventprize_regist.asp?eC=<%=arrList(0,intLoop)%>&menupos=<%=menupos%>&<%=strparm%>')"><%END IF%>
      		<%if arrList(15,intLoop)  then%> <input type="button" value="할인(<%=arrList(18,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/sale/salelist.asp?eC=<%=arrList(0,intLoop)%>&menupos=290');"><%end if%>
      		<%if arrList(16,intLoop) then%> <input type="button" value="사은품(<%=arrList(19,intLoop)%>)" class="button" onClick="jsGoUrl('/admin/shopmaster/gift/giftlist.asp?eC=<%=arrList(0,intLoop)%>&menupos=1045');"><%end if%>
      		<!--<%if arrList(17,intLoop) then%> <input type="button" value="쿠폰" class="button" onClick="jsGoUrl('coupon');"><%end if%>	-->
      		<% If arrList(20,intLoop) = "N" Then %>
      		<table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:3 0 0 0;"><input type="button" class="button" style="width:105;" value="당첨자없음 설정" onclick="prize(<%= arrList(0,intLoop) %>);"></td></tr></table>
      		<% End IF %>
      	</td>
    </tr>
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="20">등록된 내용이 없습니다.</td>
   	</tr>
   <%END IF%>
</table>
 </form>
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
