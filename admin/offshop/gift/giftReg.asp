<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 관리
' History : 2010.03.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/gift/gift_Cls.asp"-->

<%
Dim clsGift , evt_code, cEvent, gift_status, sStateDesc , gift_code ,gift_scope
Dim gift_name, gift_startdate, gift_enddate, opendate, gift_type , gift_itemname
dim makerid , gift_range1 , gift_range2 , giftkind_code , giftkind_type ,giftkind_cnt
dim giftkind_limit, gift_using ,regdate , adminid , giftkind_name , closedate
dim gift_img, gift_imgurl, gift_img_50X50_url

dim cEvtCont
dim chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , evt_opendate , evt_closedate , brand , partMDid ,evt_forward ,issale
dim evt_comment , evt_regdate , shopid , isgift ,israck ,isprize , isracknum ,racknum  ,img_basic
dim shopname
dim evt_sStateDesc

dim itemgubun, shopitemid, itemoption, shopitemname
dim gift_itemgubun, gift_shopitemid, gift_itemoption
dim receiptstring

dim gift_scope_add, giftkind_limit_sold
'
''// 사은품 영수증 표시내역
'Function fnComGetEventConditionStr(ByVal Fgiftkind_type, ByVal Fgift_scope,ByVal Fgift_type,ByVal Fgift_range1, ByVal Fgift_range2,ByVal FGiftName,ByVal Fgiftkind_cnt, ByVal Fgiftkind_orgcnt, ByVal Fgiftkind_limit, ByVal Fgiftkind_givecnt,ByVal FMakerid)
'Dim reStr
'dim remainEa
'
'        reStr = ""
'        if (FMakerid<> "") then
'        	reStr = reStr + FMakerid + " "
'        end if
'        if (Fgift_scope="1") then
'            reStr = reStr + "전체 구매 고객 "
'        elseif (Fgift_scope="2") then
'            reStr = reStr + "이벤트등록상품 "
'        elseif (Fgift_scope="3") then
'            reStr = reStr + "선택브랜드상품 "
'        elseif (Fgift_scope="4") then
'            reStr = reStr + "이벤트그룹상품"
'        elseif (Fgift_scope="5") then
'            reStr = reStr + "선택상품"
'        end if
'
'        if (Fgift_type="1") then
'            reStr = reStr + "모든 구매자"
'        elseif (Fgift_type="2") then
'            if (Fgift_range2=0) then
'                reStr = reStr + CStr(Fgift_range1) + " 원 이상 구매시 "
'            else
'                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 원 구매시 "
'            end if
'        elseif (Fgift_type="3") then
'            if (Fgift_range2=0) then
'                reStr = reStr + CStr(Fgift_range1) + " 개 이상 구매시 "
'            else
'                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 개 구매시 "
'            end if
'        end if
'        reStr = reStr &"'"&  FGiftName &"' "
'        reStr = reStr &  Cstr(Fgiftkind_orgcnt) & " 개 "
'
'        if (Fgiftkind_type=2) then
'            reStr = reStr + "[1+1]"
'             reStr = reStr & "(총 "& Cstr(Fgiftkind_cnt) & " 개)"
'        elseif (Fgiftkind_type=3) then
'            reStr = reStr + "[1:1]"
'             reStr = reStr & "(총 "& Cstr(Fgiftkind_cnt) & " 개)"
'        end if
'         reStr = reStr + " 증정"
'
'
'        if Fgiftkind_limit<>0 then
'            reStr = reStr & " 총한정 [" & Fgiftkind_limit & "]"
'            remainEa = Fgiftkind_limit-Fgiftkind_givecnt
'            if (remainEa<0) then remainEa=0
'             reStr = reStr & " 현재남은수량 " & remainEa
'        end if
'        fnComGetEventConditionStr = reStr
' End Function

menupos = requestCheckVar(request("menupos"),10)
evt_code = requestCheckVar(Request("evt_code"),10)
gift_code = requestCheckVar(Request("gift_code"),10)
'gift_type = 2

if evt_code = "" then
	Alert_return("잘못된 접속입니다. 먼저 이벤트를 등록하세요.")
	dbget.close()	:	response.End
end if

'==============================================================================
set cEvtCont = new cevent_list
	cEvtCont.frectevt_code = evt_code	'이벤트 코드

	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont_off
	evt_kind = cEvtCont.FOneItem.fevt_kind
	evt_name = cEvtCont.FOneItem.fevt_name
	evt_startdate = cEvtCont.FOneItem.Fevt_startdate
	evt_enddate = cEvtCont.FOneItem.Fevt_enddate
	evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
	evt_state =	cEvtCont.FOneItem.Fevt_state
	IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '기간 초과시 종료표기
	evt_regdate	= cEvtCont.FOneItem.fevt_regdate
	evt_using = cEvtCont.FOneItem.Fevt_using
	shopid = cEvtCont.FOneItem.fshopid
	shopname = cEvtCont.FOneItem.fshopname
	evt_opendate = cEvtCont.FOneItem.fopendate
	evt_closedate = cEvtCont.FOneItem.fclosedate

	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay_off
	chkdisp = cEvtCont.FOneItem.FChkDisp
	issale = cEvtCont.FOneItem.fissale
	isgift = cEvtCont.FOneItem.fisgift
	israck = cEvtCont.FOneItem.fisrack
	isprize = cEvtCont.FOneItem.fisprize
	isracknum = cEvtCont.FOneItem.fisracknum
	partMDid = cEvtCont.FOneItem.FpartMDid
	evt_forward	= db2html(cEvtCont.FOneItem.Fevt_forward)
	brand = cEvtCont.FOneItem.Fbrand
	evt_comment = cEvtCont.FOneItem.fevt_comment
 	chkdisp	= cEvtCont.FOneItem.fchkdisp
	img_basic = cEvtCont.FOneItem.fimg_basic

set cEvtCont = nothing



'==============================================================================
set cEvent = new cevent_list
	cEvent.Frectevt_code = evt_code
	cEvent.fnGetEventConts

	gift_name 		= cEvent.foneitem.fevt_name
	gift_startdate	= cEvent.foneitem.fevt_startdate
	gift_enddate	= cEvent.foneitem.fevt_enddate
	gift_status  	= cEvent.foneitem.fevt_state
	opendate		= cEvent.foneitem.FOpenDate

	evt_sStateDesc = cEvent.foneitem.fevt_statedesc
	sStateDesc = cEvent.foneitem.fevt_statedesc		'신규 등록일때 이벤트의 상태를 가져온다.
set cEvent = nothing



'==============================================================================
dim isregstate

isregstate = true


'//신규등록
if gift_code = "" then

	'이벤트 상태와 사은품 상태 매칭처리(오픈이전 상태는 모두 대기상태)
	if gift_status < 6 then gift_status = 0
	giftkind_cnt = 1

'//수정
else
	isregstate = false

	set clsGift = new cgift_list
		clsGift.frectgift_code = gift_code
		clsGift.fnGetGiftConts_off

		gift_name			= clsGift.foneitem.fgift_name
		gift_scope 			= clsGift.foneitem.fgift_scope
		evt_code			= clsGift.foneitem.fevt_code
		gift_type			= clsGift.foneitem.fgift_type
		gift_range1			= clsGift.foneitem.fgift_range1
		gift_range2 		= clsGift.foneitem.fgift_range2
		giftkind_code		= clsGift.foneitem.fgiftkind_code

		makerid				= clsGift.foneitem.fmakerid

		itemgubun			= clsGift.foneitem.fitemgubun
		shopitemid			= clsGift.foneitem.fshopitemid
		itemoption			= clsGift.foneitem.fitemoption
		shopitemname		= clsGift.foneitem.fshopitemname

		gift_itemgubun		= clsGift.foneitem.fgift_itemgubun
		gift_shopitemid		= clsGift.foneitem.fgift_shopitemid
		gift_itemoption		= clsGift.foneitem.fgift_itemoption

		giftkind_type		= clsGift.foneitem.fgiftkind_type
		giftkind_cnt		= clsGift.foneitem.fgiftkind_cnt
		giftkind_limit		= clsGift.foneitem.fgiftkind_limit
		gift_startdate		= clsGift.foneitem.fgift_startdate
		gift_enddate		= clsGift.foneitem.fgift_enddate
		gift_status			= clsGift.foneitem.fgift_status
		gift_using     		= clsGift.foneitem.fgift_using
		regdate				= clsGift.foneitem.fregdate
		adminid 			= clsGift.foneitem.fadminid
		giftkind_name 		= clsGift.foneitem.fgiftkind_name
		opendate			= clsGift.foneitem.fopendate
		closedate			= clsGift.foneitem.fclosedate
		gift_itemname		= clsGift.foneitem.fgift_itemname
		gift_img			= clsGift.foneitem.fgift_img

		gift_imgurl			= clsGift.foneitem.GetMobileGiftImage
		gift_img_50X50_url	= clsGift.foneitem.GetMobileGiftImage50X50

		receiptstring		= clsGift.foneitem.fnGetReceiptString

		gift_scope_add		= clsGift.foneitem.fgift_scope_add
		giftkind_limit_sold	= clsGift.foneitem.fgiftkind_limit_sold

	set clsGift = nothing

	  '공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	sStateDesc 	= fnSetCommonCodeArr_off("gift_status",False)
end if
%>

<script language="javascript">

	//사은품 등록
	function jsSubmitGift(){
		var frm = document.frmReg;

		// ====================================================================
		// 날짜 검증
		if(!frm.gift_name.value){
			alert("제목을 입력해 주세요");
			frm.gift_name.focus();
			return;
		}

		if(!frm.gift_startdate.value ){
		  	alert("시작일을 입력해주세요");
		 	frm.gift_startdate.focus();
		  	return;
	  	}

	  	if(frm.gift_enddate.value){
		  	if(frm.gift_startdate.value > frm.gift_enddate.value){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	return;
		  	}
		}

		var nowDate = "<%=date()%>";

		if(frm.gift_status.value==7){
			if(frm.opendate.value !=""){
				nowDate = '<%IF opendate <> ""THEN%><%=FormatDate(opendate,"0000-00-00")%><%END IF%>';
			}

			/*
			if(frm.gift_startdate.value < nowDate){
				alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
				return;
			}
			*/
		}

		// ====================================================================
		if (frm.gift_scope.value == "") {
			alert("증정대상을 선택하세요.");
			return;
		} else if (frm.gift_scope.value == "1") {
			//
		} else if (frm.gift_scope.value == "5") {
			// 등록 상품
			if (frm.shopitemid.value == "") {
				alert("상품을 지정해주세요.");
				return;
			}
		} else if (frm.gift_scope.value == "6") {
			// 등록 브랜드
			if (frm.makerid.value == "") {
				alert("브랜드를 지정해주세요.");
				return;
			}
		} else if (frm.gift_scope.value == "7") {
			// 증정대상직접설정
			if (frm.gift_scope_add.value == "") {
				alert("증정대상직접설정을 입력해주세요.");
				return;
			}
		}

		// ====================================================================
		if (frm.gift_type.value == "") {
			alert("증정조건을 설정해주세요.");
			return;

		} else if (frm.gift_type.value != "1") {
			if ((frm.gift_range1.value*0 != 0) || (frm.gift_range2.value*0 != 0) || (frm.gift_range1.value == "") || (frm.gift_range2.value == "")) {
				alert("증정범위를 정확히 입력하세요.");
				return;
			}

			if (frm.gift_range1.value*1 == 0) {
				if (confirm("증정범위를 0 으로 설정하였습니다. 진행하시겠습니까?") != true) {
					return;
				}
			}

		} else {
			frm.gift_range1.value = 0;
			frm.gift_range2.value = 0;
		}

		// ====================================================================
		if(frm.gift_shopitemid.value == ""){
			alert("사은품을 검색해서 입력해주세요.");
			return;
		}

		if ((frm.giftkind_cnt.value*0 != 0) || (frm.giftkind_cnt.value == "")) {
			alert("사은품 수량을 정확히 입력하세요.");
			return;
		}

		if (frm.chkLimit.checked == true) {
			if ((frm.giftkind_limit.value*0 != 0) || (frm.giftkind_limit.value == "")) {
				alert("사은품 한정 수량을 정확히 입력하세요.");
				return;
			}
		} else {
			frm.giftkind_limit.value = 0;
			frm.giftkind_limit_sold.value = 0;
		}

		// ====================================================================
		if(frm.gift_itemname.value == ""){
			frm.gift_itemname.value = frm.giftkind_name.value;
		}

		// ====================================================================
		if (confirm("저장하시겠습니까?") == true) {
			jsChkGiftScope(frm.gift_scope.value);
			jsResetHiddenData();

			frm.submit();
		}

	}

	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	// 증정대상
	// 표시만 바꾼다. 안보이는 부분 데이타 삭제는 저장할때 한다.
	function jsChkGiftScope(iVal){
		jsHideAll();

		if(iVal == 1){
			// 전체증정
		} else if (iVal == 5) {
			// 등록상품
			document.all.showitemid.style.display = "";
		} else if (iVal==6) {
			// 등록브랜드
			document.all.showmakerid.style.display = "";
		} else if (iVal==7) {
			// 매장판단
			document.all.showaddcondition.style.display = "";
		}else{
			// ERROR
		}

		if (iVal != 5) {
			// 증정대상이 등록상품이 아니면 설정해제한다.
			var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
			var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

			chk2.checked=false;
			chk3.checked=false;

			jsResetGiftNo();
		}
	}

	function jsHideAll() {
		document.all.showmakerid.style.display = "none";
		document.all.showitemid.style.display = "none";
		document.all.showaddcondition.style.display = "none";
	}

	function jsResetHiddenData() {
		// 브랜드정보 리셋
		if (document.all.showmakerid.style.display == "none") {
			document.all.makerid.value = "";
		}

		// 상품정보 리셋
		if (document.all.showitemid.style.display == "none") {
			document.all.itemgubun.value = "";
			document.all.shopitemid.value = "";
			document.all.itemoption.value = "";
			document.all.shopitemname.value = "";
		}

		// 증정대상직접설정 리셋
		if (document.all.showaddcondition.style.display == "none") {
			document.all.gift_scope_add.value = "";
		}
	}

	function jsResetGiftNo() {
		// 사은품 수량설정
		var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
		var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

		if ((chk2.checked == true) || (chk3.checked == true) || (document.all.gift_type.value == 1)) {
			document.all.gift_range1.readOnly=true;
			document.all.gift_range2.readOnly=true;
			document.all.giftkind_cnt.readOnly=true;
			document.all.gift_range1.style.backgroundColor='#E6E6E6';
			document.all.gift_range2.style.backgroundColor='#E6E6E6';
			document.all.giftkind_cnt.style.backgroundColor='#E6E6E6';

			if (document.all.gift_type.value == 1) {
				document.all.giftkind_cnt.readOnly=false;
				document.all.giftkind_cnt.style.backgroundColor='';
			}
		} else {
			document.all.gift_range1.readOnly=false;
			document.all.gift_range2.readOnly=false;
			document.all.giftkind_cnt.readOnly=false;
			document.all.gift_range1.style.backgroundColor='';
			document.all.gift_range2.style.backgroundColor='';
			document.all.giftkind_cnt.style.backgroundColor='';

			return;
		}

		if (chk2.checked == true) {
			// 1+1
			document.all.gift_range1.value=1;
			document.all.gift_range2.value=0;
			document.all.giftkind_cnt.value=1;
		} else if (chk3.checked == true) {
			// 1:1
			document.all.gift_range1.value=1;
			document.all.gift_range2.value=0;
			document.all.giftkind_cnt.value=1;
		} else {
			// 없음
		}
	}

	function jsChkGiftType(iVal){
		if(iVal==1){
			document.all.gift_range1.readOnly=true;
			document.all.gift_range2.readOnly=true;
			document.all.gift_range1.style.backgroundColor='#E6E6E6';
			document.all.gift_range2.style.backgroundColor='#E6E6E6';

			document.all.gift_range1.value=0;
			document.all.gift_range2.value=0;
		}else{
			document.all.gift_range1.readOnly=false;
			document.all.gift_range2.readOnly=false;
			document.all.gift_range1.style.backgroundColor='';
			document.all.gift_range2.style.backgroundColor='';
		}

		if (iVal != 3) {
			// 증정조건이 수량이 아니면 설정해제한다.
			var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
			var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

			chk2.checked=false;
			chk3.checked=false;

			jsResetGiftNo();
		}
	}

	// 1+1 ,1:1 체크
	function jsCheckKT(ev,ch){

		var chk 	= document.getElementById(ev);
		var chftf 	= chk.checked;
		var chk2 	= document.getElementById('tmpgiftkind_givecnt2');
		var chk3 	= document.getElementById('tmpgiftkind_givecnt3');

		chk2.checked=false;
		chk3.checked=false;

		if (document.all.gift_scope.value != 5) {
			alert("증정대상을 등록상품으로 설정해야만 체크할 수 있습니다.");
			return;
		}

		if (document.all.gift_type.value != 3) {
			alert("증정조건을 수량으로 설정해야만 체크할 수 있습니다.");
			return;
		}

		chk.checked=chftf;

		if(chftf){
			document.frmReg.giftkind_givecnt.value= chk.value;
		}else{
			document.frmReg.giftkind_givecnt.value=0;
		}
		jsResetGiftNo();
	}

	function jsChkLimit(){
		if(document.frmReg.chkLimit.checked){
			document.all.giftkind_limit.readOnly=false;
			document.all.giftkind_limit.style.backgroundColor='';
		}else{
			document.all.giftkind_limit.readOnly=true;
			document.all.giftkind_limit.style.backgroundColor='#E6E6E6';
			document.frmReg.giftkind_limit.value = "";
		}
	}

	// 사은품등록내역 가져오기
	function jsImport(ec){
		var pp = window.open('/admin/offshop/gift/popGiftList.asp?eC='+ec,'popim','scrollbars=yes,resizable=yes,width=900,height=600');

	}

	// 사은품 검색
	function jsSearchGiftItem(){
		var winkind;
		winkind = window.open('popgiftKindReg.asp?giftkind_name='+ urlencode(document.frmReg.giftkind_name.value),'popkind','width=800, height=500,scrollbars=yes,resizable=yes');
		// winkind = window.open('popgiftKindReg.asp?giftkind_name='+ document.frmReg.giftkind_name.value,'popkind','width=800, height=500,scrollbars=yes,resizable=yes');
		winkind.focus();
	}

	// 등록상품 검색
	function jsSearchTargetItem(itemgubun){
		var winkind;
		winkind = window.open('popTargetItemReg.asp?shopitemname='+document.frmReg.shopitemname.value + '&itemgubun=' + itemgubun,'jsSearchTargetItem','width=800, height=600,scrollbars=yes,resizable=yes');
		winkind.focus();
	}

	function urlencode(plaintxt) {
		   return escape(plaintxt).replace("+","%2C");
	}

	window.onload = function() {
		jsChkGiftScope(document.all.gift_scope.value);
	}

function popUploadGiftItemimage(frm) {
	var mode, imagekind, pk;

	if (frm.gift_code.value == "") {
		alert("먼저 사은품 정보를 저장하세요");
		return;
	}

	if (frm.gift_img.value == "") {
		mode = "addimage";
	} else {
		mode = "editimage";
	}

	imagekind = "mobilegiftitemimage";
	pk = frm.gift_code.value;


	var popwin = window.open("/common/pop_upload_image.asp?mode=" + mode + "&imagekind=" + imagekind + "&pk=" + pk + "&50X50=Y","popUploadGiftItemimage","width=390 height=120 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1"  >
<form name="frmReg" method="post" action="giftProc.asp">
<input type="hidden" name="mode" value="giftedit">
<input type="hidden" name="gift_code" value="<%=gift_code%>">
<input type="hidden" name="giftkind_givecnt" value="0">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드</td>
			<td bgcolor="#FFFFFF">
				<%=evt_code%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">적용샵</td>
			<td bgcolor="#FFFFFF">
				<%= shopid %>(<%= shopname %>)
			</td>
		</tr>
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트명</td>
			<td bgcolor="#FFFFFF">
				<%=evt_name%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <%= evt_startdate %> ~ 종료일 : <%= evt_enddate %>
			</td>
		</tr>
		<tr height="25">
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">상태</td>
			<td bgcolor="#FFFFFF">
				<%=replace(evt_sStateDesc,"오픈예정","오픈")%>
			</td>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">브랜드</td>
			<td bgcolor="#FFFFFF">
				<%= brand %>
			</td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
			<td bgcolor="#FFFFFF" colspan="3"><input type="button" class="button" value="이전 사은품정보 가져오기" onClick="jsImport('<%= evt_code %>');"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF" width="400">
				<font color="gray"><%=gift_name%></font><input type="hidden" name="gift_name" value="<%=gift_name%>">
			</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<input type="hidden" name="gift_startdate" value="<%=gift_startdate%>">
			<input type="hidden" name="gift_enddate" value="<%=gift_enddate%>">
			<td bgcolor="#FFFFFF">
				<font color="gray">
				시작일 :
				<%=gift_startdate%>
				~ 종료일 :
				<%=gift_enddate%>
				</font>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">증정대상</td>
			<td bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "gift_scope", gift_scope, isregstate, True, "onchange='jsChkGiftScope(this.value);'" %>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
			<input type="hidden" name="gift_status" value="<%=gift_status%>">
			<td bgcolor="#FFFFFF">
				<% if gift_code = "" then %>
						<font color="gray"><%=replace(sStateDesc,"오픈예정","오픈")%></font>
				<% else %>
						<%=replace(fnGetCommCodeArrDesc_off(sStateDesc,gift_status),"오픈예정","오픈")%>
				<% end if %>
				<input type="hidden" name="opendate" value="<%=opendate%>">
				<input type="hidden" name="closedate" value="<%=closedate%>">
				<%IF opendate <> "" THEN%><span style="padding-left:10px;">오픈처리일: <%=opendate%></span><%END IF%>
				<%IF closedate <> "" THEN%><br><span style="padding-left:42px;">종료처리일: <%=closedate%></span><%END IF%>
			</td>
		</tr>

		<tr id="showmakerid" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">등록브랜드</td>
			<td  width="400" bgcolor="#FFFFFF" colspan="3">
				<% drawSelectBoxDesignerwithName "makerid", makerid %>
			</td>
		</tr>

		<tr id="showitemid" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">등록상품</td>
			<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
			<input type="hidden" name="shopitemid" value="<%= shopitemid %>">
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			<td bgcolor="#FFFFFF" colspan="3">
				<input type="text" name="shopitemname"  value="<%= shopitemname %>" size="40" maxlength="60" style="background-color:#E6E6E6;" readonly>
				<input type="button" class="button" value="ON 상품검색" onClick="jsSearchTargetItem('10');">
				<input type="button" class="button" value="OFF 상품검색" onClick="jsSearchTargetItem('90');">
			</td>
		</tr>

		<tr id="showaddcondition" style="display:none">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정대상직접설정</td>
			<td bgcolor="#FFFFFF" colspan="3">
				<input type="text" name="gift_scope_add" size="30" value="<%= gift_scope_add %>">
				* 예시 : 2011년도 대학교 신입생 한정, 주부한정
			</td>
		</tr>

		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정조건</td>
			<td width="400" bgcolor="#FFFFFF">
				<%sbGetOptCommonCodeArr_off "gift_type", gift_type, isregstate, True,"onchange='jsChkGiftType(this.value);'" %>
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center" width="100">증정범위</td>
			<td bgcolor="#FFFFFF">
				<% if gift_code = "" then %>
					<input type="text" name="gift_range1" size="10" style="text-align:right" value="0"> 이상
					~ <input type="text" name="gift_range2" size="10" style="text-align:right" value="0"> 미만
				<% else %>
					<input type="text" name="gift_range1" size="10" style="text-align:right;<%IF gift_type= "1" THEN%>background-color:#E6E6E6; readonly<%ELSE%>"<%END IF%> value="<%=gift_range1%>"> 이상
					~ <input type="text" name="gift_range2" size="10" style="text-align:right;<%IF gift_type= "1" THEN%>background-color:#E6E6E6; readonly<%ELSE%>"<%END IF%> value="<%=gift_range2%>"> 미만
				<% end if %>
				(ex. 20개 이상: 20~0)
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품명</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="hidden" name="gift_itemgubun" value="<%= gift_itemgubun %>">
				<input type="hidden" name="gift_shopitemid" value="<%= gift_shopitemid %>">
				<input type="hidden" name="gift_itemoption" value="<%= gift_itemoption %>">
				<input type="hidden" name="giftkind_code" value="<%=giftkind_code%>">
				<input type="text" name="giftkind_name"  value="<%=giftkind_name%>" size="40" maxlength="60" style="background-color:#E6E6E6;" readonly>
				<input type="button" class="button" value="사은품검색" onClick="jsSearchGiftItem();">
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품수량</td>
			<td bgcolor="#FFFFFF">
					<input type="text" name="giftkind_cnt" size="4" maxlength="10" value="<%=giftkind_cnt%>" style="text-align:right;"> 개씩
						<label title="동일상품증정 1+1" ><input type="checkbox" name="tmpgiftkind_givecnt2" onclick="jsCheckKT('tmpgiftkind_givecnt2');"  value="2" <%IF CStr(giftkind_type) = "2" THEN%>checked<%END IF%>>1+1(동일상품) </label>
						<label title="다른상품증정 1:1" ><input type="checkbox" name="tmpgiftkind_givecnt3" onclick="jsCheckKT('tmpgiftkind_givecnt3');" value="3" <%IF CStr(giftkind_type) = "3" THEN%>checked<%END IF%>>1:1(다른상품) </label>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사은품한정수량</td>
			<td bgcolor="#FFFFFF">
				<input type="checkbox" name="chkLimit" onClick="jsChkLimit();" <%IF giftkind_limit > 0 THEN%>checked<%END IF%> <% if (shopid = "") or (shopid = "all") then %>disabled<% end if %>>한정
				<input type="text" name="giftkind_limit" size="4" value="<%=giftkind_limit%>" style="text-align:right;" <%IF giftkind_limit = 0 THEN%>style="background-color:#E6E6E6;" readonly<%END IF%>> 개(한정수량 있을 경우에만 입력)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">한정소모수량</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="giftkind_limit_sold" size="4" value="<%=giftkind_limit_sold%>" style="text-align:right;">
			</td>
		</tr>
		<% if gift_code <> "" then %>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용유무</td>
			<td bgcolor="#FFFFFF" colspan=3>
				<input type="radio" name="gift_using" value="Y" <%IF gift_using = "Y" THEN%>checked<%END IF%>>사용
				<input type="radio" name="gift_using" value="N" <%IF gift_using = "N" THEN%>checked<%END IF%>>사용안함
			</td>
		</tr>
		<% end if %>
		<tr height="25">
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">영수증출력</td>
			<td bgcolor="#FFFFFF" colspan=3>
				<%= receiptstring %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">사은품명<br>(모바일앱표시)</td>
			<td  width="400" bgcolor="#FFFFFF">
				<input type="text" name="gift_itemname"  value="<%=gift_itemname%>" size="40" maxlength="60">
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"  width="100">이미지<br>(모바일앱표시)</td>
			<td bgcolor="#FFFFFF">
				<% if (gift_code <> "") then %>
					<% if (gift_img <> "") then %>
						<img src="<%= gift_img_50X50_url %>"><br>
						<img src="<%= gift_imgurl %>"><br>
						<input type="button" class="button" value="수정하기" onclick="popUploadGiftItemimage(frmReg)">
					<% else %>
						<input type="button" class="button" value="등록하기" onclick="popUploadGiftItemimage(frmReg)">
					<% end if %>
				<% end if %>
				<input type="hidden" name="gift_img" value="<%= gift_img %>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="button" onclick="jsSubmitGift();" value="저장" class="button">
		<input type="button" onclick="history.back();" value="취소" class="button">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->