<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  할인 관리 등록
' History : 2008.04.07 정윤정 생성
'			2022.07.06 한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim sMode
Dim sCode, clsSale,isRate, isMargin, isStatus, egCode, isUsing, dOpenDay,isMValue,dCloseDay
Dim eCode, cEvent
Dim sTitle, dSDay, dEDay, sBrand,eState
Dim cEGroup,blngroup,arrGroup,intgroup

Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate, sStatus
Dim strParm

Dim  clsSaleItem
Dim  smargin
Dim acURL
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage, iSubCurrpage,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
dim makerid, sailyn,invalidmargin, sRectItemidArr
dim intTime, sSType, sSTime, sETime
dim sSalestatus, sItemSale


eCode     = requestCheckVar(Request("eC"),10)
sCode     = requestCheckVar(Request("sC"),10)
makerid =  requestCheckVar(Request("makerid"),32)
sailyn	=  requestCheckVar(Request("sailyn"),1)

invalidmargin=  requestCheckVar(Request("invalidmargin"),1)
sRectItemidArr=  requestCheckVar(Request("sRectItemidArr"),400)
sSalestatus 	= requestCheckVar(Request("salestatus"),4)
sItemSale	= requestCheckVar(Request("selItemStatus"),4)

acURL =Server.HTMLEncode("/admin/shopmaster/sale/saleitemProc.asp?sC="&sCode)

iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
IF iCurrpage = "" THEN	iCurrpage = 1
iSubCurrpage = requestCheckVar(Request("iSC"),10)	'현재 상품목록 페이지 번호
IF iSubCurrpage = "" THEN	iSubCurrpage = 1
iPageSize = 50		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격
isRate = 0
isUsing = true
sMode  = "I"
isStatus =0
sSType = 1

if sRectItemidArr<>"" then
	dim iA ,arrTemp,arrItemid
	sRectItemidArr = replace(sRectItemidArr,",",chr(10))
	sRectItemidArr = replace(sRectItemidArr,chr(13),"")
	arrTemp = Split(sRectItemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	if arrItemid<>"" and not(isnull(arrItemid)) then
	sRectItemidArr = trim(left(arrItemid,len(arrItemid)-1))
	end if
end if

'마진형태에 따른 매입가 생성-------------------------------------------------------
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice)
	Dim orgMRate
	if orgPrice <>0 then '원 마진율
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if

	SELECT CASE MarginType
		Case 1	'동일마진
			fnSetSaleSupplyPrice = salePrice- fix(salePrice*(orgMRate/100))
		Case 2	'업체부담
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'반반부담
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10부담
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'직접설정
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT
End Function
'-----------------------------------------------------------------------------------
IF sCode <> "" THEN
	set clsSale = new CSale
	sMode = "U"
	clsSale.FSCode  = sCode
	clsSale.fnGetSaleConts

	sTitle 		= clsSale.FSName
	isRate 		= clsSale.FSRate
	isMargin 	= clsSale.FSMargin
	eCode 		= clsSale.FECode
	egCode		= clsSale.FEGroupCode
	dSDay 		= clsSale.FSDate
	dEDay 		= clsSale.FEDate
	isStatus 	= clsSale.FSStatus
	isUsing     = clsSale.FSUsing
	dOpenDay	= clsSale.FOpenDate
	isMValue	= clsSale.FSMarginValue
	dCloseDay 	= clsSale.FCloseDate
	sSType      = clsSale.FSType
	set clsSale = nothing

	sSTime = mid(dSDay,12,2)
	sETime = mid(dEDay,12,2)
	dSDay = left(dSDay,10)
	dEDay = left(dEDay,10)


	'할인 상품정보
	set clsSaleItem = new CSaleItem
	clsSaleItem.FCPage = iSubCurrpage			'// skyer9, 2017-08-16
	clsSaleItem.FPSize = iPageSize
	clsSaleItem.FSCode = sCode
	clsSaleItem.FRectMakerid = makerid
	clsSaleItem.FRectsailyn = sailyn
	clsSaleItem.FRectinvalidmargin =invalidmargin
	clsSaleItem.FRectItemidArr = sRectItemidArr
	clsSaleItem.FRectSaleStatus = sSalestatus
  clsSaleItem.FRectItemSaleStatus = sItemSale

	arrList = clsSaleItem.fnGetSaleItemList
	iTotCnt = clsSaleItem.FTotCnt	'전체 데이터  수

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수

	'동기간내 상품쿠폰 정보 접수
	Dim arrItemCoupon, iclp
	arrItemCoupon = clsSaleItem.fnGetCouponListBySaleInfo
	set clsSaleItem = nothing

END IF
  
	IF eCode = "0" THEN eCode = ""
	IF eCode <> "" THEN		'이벤트 연관 일경우
		IF sCode = "" THEN
		set cEvent = new ClsEventSummary
			cEvent.FECode = eCode
			cEvent.fnGetEventConts
			sTitle 	= cEvent.FEName
			dSDay	= cEvent.FESDay
			dEDay	= cEvent.FEEDay
			sBrand	= cEvent.FBrand
			isStatus= cEvent.FEState
			dOpenDay= cEvent.FEOpenDate
		set cEvent = nothing
	   END IF
		set cEGroup = new ClsEventGroup
		 	cEGroup.FECode = eCode
		  	arrGroup = cEGroup.fnGetEventItemGroup
		set cEGroup = nothing

		 blngroup = False
		 IF isArray(arrGroup) THEN blngroup = True
	END IF

	IF dSDay ="" THEN dSDay = date()
	IF isStatus < 6 THEN isStatus = 0
    IF sETime ="" then sETime = 23
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim arrsalemargin, arrsalestatus
	arrsalemargin = fnSetCommonCodeArr("salemargin",False)
	arrsalestatus= fnSetCommonCodeArr("salestatus",False)

'-검색----------------------------------------
	' iSerachType    = requestCheckVar(Request("selType"),4)		'검색구분
	' sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'검색어
	' sDate     		= requestCheckVar(Request("selDate"),1)		'검색일 기준
	' sSdate     	= requestCheckVar(Request("iSD"),10)		'시작일
	' sEdate     	= requestCheckVar(Request("iED"),10)		'종료일
	' sStatus		= requestCheckVar(Request("salestatus"),4)	' 상태
	' iCurrpage		= requestCheckVar(Request("iC"),10)	'현재 페이지 번호

	' strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&sStatus
'---------------------------------------------
Function numBerBurim(idx, sosu)
	Dim tmpSu
	tmpSu = FormatNumber(idx - 0.5/10^sosu, sosu)
	If cstr(int(tmpSu)) = cstr(formatnumber(tmpSu,0)) Then
		numBerBurim = formatnumber(tmpSu,0)
	Else
		numBerBurim = tmpSu
	End If
End Function
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
	//-- jsPopCal : 달력 팝업 --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function jsSubmitSale(){
		var frm = document.frmReg;
		var dt = new Date();
    var month = dt.getMonth()+1;
    var day = dt.getDate();
    var year = dt.getFullYear();
   	var stime = dt.getHours();

    var StTime = $("#sSTi").val();
		var EtTime = $("#sETi").val();
		if (StTime.length == 1 ){ StTime = "0"+StTime}
		if (EtTime.length == 1 ){ EtTime = "0"+EtTime}
		var StDate = frm.sSD.value +" "+StTime +":"+ $("#sSTSec").val();
		var EtDate = frm.sED.value +" "+EtTime +":"+ $("#sETSec").val();

	    var saletype = $("input:radio[name='rdoT']:checked").val();

        if ( month <10 ){ month = "0"+month }
        if (day<10 ){ day = "0"+day}
        if (stime<10 ){ stime = "0"+stime}

		var nowDate ;
		if (saletype==2){	nowDate	= year+"-"+ month+"-"+day+" "+stime+":00:00" }
		else {nowDate	= year+"-"+ month+"-"+day+" 00:00:00"}

		if(typeof(frm.chkstatus)=="object"){
			if(frm.chkstatus.checked) {
				frm.salestatus.value = frm.chkstatus.value;
			}
		}

		if(!frm.sSN.value){
			alert("제목을 입력해 주세요");
			frm.sSN.focus();
			return false;
		}

		if(!frm.sSD.value ){
		  	alert("시작일을 입력해주세요");
		  	frm.sSD.focus();
		  	return false;
	  	}

	 if(frm.salestatus.value==7){
	 	if(frm.sOD.value !=""){
	 	    if (saletype == 2)  {
	 	    	nowDate = '<%IF dOpenDay <> "" THEN%><%=dOpenDay%><%END IF%>';
	 	    }else{
	 	        nowDate = '<%IF dOpenDay <> "" THEN%><%=FormatDate(dOpenDay,"0000-00-00")%><%END IF%>';
	 	    }
		}

		if(StDate < nowDate){
			alert("시작일이 오픈일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
		  	frm.sSD.focus();
		  	return false;
		 }
	  }

	  if(!frm.sED.value ){
		  	alert("종료일을 입력해주세요");
		  	frm.sED.focus();
		  	return false;
	  	}


	   	if(EtDate < nowDate){
				alert("종료일이 현재 날짜보다  빠르면 안됩니다. 종료일을 다시 선택해주세요");
			  	frm.sED.focus();
			  	return false;
			 }

		  	if(StDate > EtDate){
			  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
			  	frm.sED.focus();
			  	return false;
		  	}


	  	if(!frm.iSR.value){
			alert("할인율을 입력해 주세요");
			frm.iSR.focus();
			return false;
		}


	}

	function jsChSetValue(iVal){
		if(iVal ==5){
			document.all.divM.style.display = "";
		}else{
			document.all.divM.style.display = "none";
		}
	}

	// 페이지 이동
function jsGoPage(iP){
	location.href="saleReg.asp?menupos=<%=menupos%>&sC=<%=sCode%>&eC=<%=eCode%>&makerid=<%=makerid%>&sailyn=<%=sailyn%>&invalidmargin=<%=invalidmargin%>&sRectItemidArr=<%=sRectItemidArr%>&salestatus=<%=sSalestatus%>&selItemStatus=<%=sItemSale%>&iSC="+iP;
}

// 새상품 추가 팝업
function addnewItem(eC,egC){
		var popwin;
		if ( eC > 0 ){
			popwin = window.open("/admin/eventmanage/common/pop_eventitem_addinfo.asp?acURL=<%=acURL%>&eC="+eC+"&egC="+egC, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}else{
			popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>&PR=S&sC=<%=sCode%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}
		popwin.focus();
}

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function CkDisPrice(){
	CkDisOrOrg(true);
}

function CkOrgPrice(){
	CkDisOrOrg(false);
}

//원가 할인가 적용
function CkDisOrOrg(isDisc){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(isDisc==true){
					frm.iDSPrice.value = Math.floor10(frm.saleprice.value, 1);
					frm.iDBPrice.value = frm.salesupplyprice.value;
					frm.iDSMargin.value= frm.salemargin.value;
					frm.saleItemStatus.value = 7;
				}else{
					if(frm.orgSailYn.value =="Y"){
						frm.iDSPrice.value = frm.orgSailPrice.value;
						frm.iDBPrice.value = frm.orgSailSupplyPrice.value;
						frm.iDSMargin.value= frm.orgSailMarginValue.value;
					}else{
						frm.iDSPrice.value = frm.orgPrice.value;
						frm.iDBPrice.value = frm.orgSupplyPrice.value;
						frm.iDSMargin.value= frm.orgMarginValue.value;
					}
					frm.saleItemStatus.value = 9;
				}
			}
			reCALbyPrice(frm.itemid.value);
		}
	}
}

//선택상품 저장
function saveArr(){
	var frm;
	var pass = false;
	var ovPer = 0;
	var ovLimitPer= 0;
	var ovLimitID= "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	frmarr.itemid.value = "";
	frmarr.sailyn.value = "";
	frmarr.iDSPrice.value ="";
	frmarr.iDBPrice.value ="";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				//check Not AvaliValue
				if (!IsDigit(frm.iDSPrice.value)){
					alert('숫자만 가능합니다.');
					frm.iDSPrice.focus();
					return;
				}

				if (frm.iDSPrice.value<1){
					alert('금액을 정확히 입력하세요.');
					frm.iDSPrice.focus();
					return;
				}

				if (!IsDigit(frm.iDBPrice.value)){
					alert('숫자만 가능합니다.');
					frm.iDBPrice.focus();
					return;
				}

				if (frm.iDBPrice.value<1){
					alert('금액을 정확히 입력하세요.');
					frm.iDBPrice.focus();
					return;
				}

				if(Math.round((frm.orgPrice.value-frm.iDSPrice.value)/frm.orgPrice.value*100)>=50) {
					ovPer++;
				}

				if(frm.mwdiv.value!='M') {
					var limitMarPrc = frm.orgSupplyPrice.value-((frm.orgPrice.value-frm.iDSPrice.value)*0.5);
					var limitMarPer = (frm.iDSPrice.value-limitMarPrc)/frm.iDSPrice.value*100;
					if(parseInt(limitMarPrc)>parseInt(frm.iDBPrice.value)) {
						ovLimitPer++;
						ovLimitID+= frm.itemid.value+"("+limitMarPer.toFixed(1)+"%),"
					}
				}

				frmarr.itemid.value = frmarr.itemid.value + frm.itemid.value + ","
				//if (frm.sailyn[0].checked){
					//frmarr.sailyn.value = frmarr.sailyn.value + "Y" + ","
				//}else{
					//frmarr.sailyn.value = frmarr.sailyn.value + "N" + ","
				//}
				frmarr.iDSPrice.value = frmarr.iDSPrice.value + frm.iDSPrice.value + ","
				frmarr.iDBPrice.value = frmarr.iDBPrice.value + frm.iDBPrice.value + ","
				frmarr.saleItemStatus.value = frmarr.saleItemStatus.value + frm.saleItemStatus.value+","

			}
		}
	}

	if(ovPer>0) {
		if(!confirm('!!!\n\n\n선택 상품중에 할인율이 매우 높은 상품(50%+)이 있습니다!\n\n입력하신 내용이 맞습니까?\n\n')) {
			return;
		}
	}

	if(ovLimitPer>0) {
		if(!confirm('!!!\n\n선택 상품 중 업체 부담율이 50%가 초과하는 상품이 있습니다!\n'+ovLimitID+'\n\n입력하신 내용이 맞습니까?\n\n')) {
			return;
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frmarr.submit();
	}

}

function delArr(){
	var frm;
	var pass = false;
	var frmdel = document.frmdel;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	frmdel.itemid.value = "";

	frmdel.itemid.value = "";
	frmdel.sailyn.value = "";
	frmdel.iDSPrice.value ="";
	frmdel.iDBPrice.value ="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if(frm.orgSailYn.value =="Y"){
					frm.iDSPrice.value = frm.orgSailPrice.value;
					frm.iDBPrice.value = frm.orgSailSupplyPrice.value;
					frm.iDSMargin.value= frm.orgSailMarginValue.value;
				}else{
					frm.iDSPrice.value = frm.orgPrice.value;
					frm.iDBPrice.value = frm.orgSupplyPrice.value;
					frm.iDSMargin.value= frm.orgMarginValue.value;
				}
				frm.saleItemStatus.value = 9;
				frmdel.itemid.value = frmdel.itemid.value + frm.itemid.value + ","
				frmdel.iDSPrice.value = frmdel.iDSPrice.value + frm.iDSPrice.value + ","
				frmdel.iDBPrice.value = frmdel.iDBPrice.value + frm.iDBPrice.value + ","
				frmdel.saleItemStatus.value = frmdel.saleItemStatus.value + frm.saleItemStatus.value+","
			}
			reCALbyPrice(frm.itemid.value);
		}
	}

	var ret = confirm('삭제하시겠습니까?');

	if (ret){
		frmdel.submit();
	}

}

// 마진율 재계산
function reCALbyPrice(fid) {
	var frm = document["frmBuyPrc_" + fid];
	var tmpsalePercent;
	var SplitSalePer;
	if(frm.iDSPrice.value>0) {
		frm.iDSMargin.value = Math.round(((frm.iDSPrice.value-frm.iDBPrice.value)/frm.iDSPrice.value)*100);
	} else {
		frm.iDSMargin.value = 0;
	}

	//할인율 표시
	var iorgPrice = frm.orgPrice.value;
	var isailprice = frm.iDSPrice.value;

	//var isalePercent = Math.round((iorgPrice-isailprice)/iorgPrice*100);
	var tmpsalePercent = (iorgPrice-isailprice)/iorgPrice*100;
	var isalePercent = (parseInt(tmpsalePercent*100)/100).toFixed(2);
	SplitSalePer = String(isalePercent).split(".")[1];

	if(isalePercent>=50) {
		document.getElementById("lyrSpct"+fid).style.color="#EE0000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="bold";
	} else {
		document.getElementById("lyrSpct"+fid).style.color="#000000";
		document.getElementById("lyrSpct"+fid).style.fontWeight="normal";
	}
	if(SplitSalePer != "00"){
		document.getElementById("lyrSpct"+fid).innerHTML = isalePercent + "%";
	}else{
		document.getElementById("lyrSpct"+fid).innerHTML = String(isalePercent).split(".")[0] + "%";
	}
}

// 매입가 재계산
function reCALbyMargin(fid) {
	var frm = document["frmBuyPrc_" + fid];
	if(frm.iDSMargin.value>0) {
		frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(frm.iDSMargin.value/100)));
	} else {
		frm.iDBPrice.value = frm.iDSPrice.value;
	}
}

function lfn_keychk(obj){
	var val = obj.value;

	if (val){
		var re = /[^0-9|.]/gi;
		obj.value = val.replace(re, '');

		var split = val.split(".");
		if(split.length > 2) {  //콤마 1개 이상 못들어오도록.
			obj.value = val.substr(0,val.length-1);
		}

		if(split[0] > 99){   // 정수 2자리 이상 입력못하도록
			if(split[0].length > 2) {
				obj.value = val.substr(0,val.length-1);
			}
		}

		if(split[1] != null){   //소수점 아래 2자리 넘지못하도록.
			if(split[1].length > 2) {
				obj.value = val.substr(0,val.length-1);
			}
		}
	}
}

function decimalAdjust(type, value, exp) {
	if (typeof exp === 'undefined' || +exp === 0) {
		return Math[type](value);
	}
	value = +value;
	exp = +exp;
	if (isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0)) {
		return NaN;
	}
	value = value.toString().split('e');
	value = Math[type](+(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp)));
	value = value.toString().split('e');
	return +(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp));
}
if (!Math.floor10) {
	Math.floor10 = function(value, exp) {
	return decimalAdjust('floor', value, exp);
	};
}

//세일모드체크
function salemodeChk(){
	var saleMode;
	var salemodePer;
	saleMode = document.getElementById("salemode").value;
	salemodePer = document.getElementById("salemodePer").value;

	if(saleMode == ""){
		alert("일괄 적용할 것을 선택해주세요");
		document.frmdummi.salemode.focus();
		return;
	}
	if(salemodePer == ""){
		alert("숫자를 입력하세요");
		document.frmdummi.salemodePer.focus();
		return;
	}

	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	if (saleMode == "1"){
		alert('적용한 할인율에 대한 판매가는 1원단위 내림한 값입니다.\n\nex)15897 -> 15890원이 됩니다');
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (saleMode == "1"){
					frm.iDSPrice.value = frm.orgPrice.value - ((frm.orgPrice.value * salemodePer) / 100)
					frm.iDSPrice.value = Math.floor10(frm.iDSPrice.value, 1)
				}else{
					if(salemodePer>0) {
						frm.iDBPrice.value = Math.round(frm.iDSPrice.value*(1-(salemodePer/100)));
					} else {
						frm.iDBPrice.value = frm.iDSPrice.value;
					}
					frm.iDSMargin.value = salemodePer;
				}
			}
		}
	}
}
// 새상품 추가 엑셀 팝업
function pop_upload(eC,egC){
	var popwin;
	popwin = window.open("/admin/shopmaster/sale/popRegFile.asp?sC=<%=sCode%>&eC="+eC+"&egC="+egC, "popup_item", "width=500,height=230,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 상품 다운로드
function pop_down(eC,egC){
	document.frmSearch.target="hidifr";
	document.frmSearch.action="saleItemReg_csv.asp";
	document.frmSearch.submit();
}

//시간 설정 변경가능여부
function jsdvTime(iType){
	if (document.frmReg.iTotCnt.value > 0 ){
		alert("등록된 상품이 있을때는 할인타입 변경이 불가능합니다.");
		if (iType==1)	{
			document.frmReg.rdoT[0].checked = false;
			document.frmReg.rdoT[1].checked = true;
		}else{
			document.frmReg.rdoT[0].checked = true;
			document.frmReg.rdoT[1].checked = false;
		}
		return;
	}
    if (iType == 1){
    	document.getElementById('spST').style.display = "none";
    	document.getElementById('spET').style.display = "none";
        document.getElementById('sSTi').disabled = true;
        document.getElementById('sETi').disabled = true;
         document.getElementById('sSTi').value = 0;
         document.getElementById('sETi').value = 0;
    }else{
    	document.getElementById('spST').style.display= "";
    	document.getElementById('spET').style.display= "";
        document.getElementById('sSTi').disabled  = false;
        document.getElementById('sETi').disabled  = false;
    }
}

//상품검색
function jsSearch(){
	document.frmSearch.target="_self";
	document.frmSearch.action="saleReg.asp";
	document.frmSearch.submit();
}

</script>
<iframe id="hidifr" src="" width="0" height="0" frameborder="0"></iframe>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  >
<form name="frmReg" method="post" action="saleProc.asp?<%=strParm%>" onSubmit="return jsSubmitSale();">
<input type="hidden" name="sM" value="<%=sMode%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iTotCnt" value ="<%=iTotCnt%>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<%IF eCode <> "" THEN%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">이벤트코드(그룹)</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=eCode%>
			<div id="dgiftgroup" style="display:<%IF NOT blngroup THEN%>none<%END IF%>;">
			<%IF isArray(arrGroup) THEN%>
				그룹선택: <select name="selG">
			   	<%
			   		For intgroup = 0 To UBound(arrGroup,2)
			   	%>
			   		<option value="<%=arrGroup(0,intgroup)%>" <%IF Cstr(egCode) = Cstr(arrGroup(0,intgroup)) THEN %> selected<%END IF%>> <%=arrGroup(0,intgroup)%>(<%=db2html(arrGroup(1,intgroup))%>)</option>
				<%	Next
				%>
			   	</select>
			 <%ELSE%>
			 <input type="hidden" name="selG" value="0">
			 <%END IF%>
			</div>
			</td>
		</tr>
		<%END IF%>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center">할인코드</td>
			<td  bgcolor="#FFFFFF"><%=sCode%></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 할인타입</td>
			<td  bgcolor="#FFFFFF">
				<input type="radio" value="1" name="rdoT" onClick="jsdvTime(1);" <%if sSType =1 then%>checked<%end if%>>일반할인
				<input type="radio" value="2" name="rdoT" onClick="jsdvTime(2);" <%if sSType =2 then%>checked<%end if%>>타임특가
			</td>
		</tr>
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>"  align="center"> 제목</td>
			<td bgcolor="#FFFFFF">
				<%IF eCode <> "" THEN %>
					<%=sTitle%><input type="hidden" name="sSN" value="<%= ReplaceBracket(sTitle) %>">
				<%ELSE%>
					<input type="text" name="sSN" size="30" maxlength="64" value="<%= ReplaceBracket(sTitle) %>" class="input">
				<%END IF%>
			</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <%IF eCode <> "" THEN %><%=dSDay%><input type="hidden" name="sSD" value="<%=dSDay%>">
				            <%ELSE%><input type="text" name="sSD" size="10"   onClick="jsPopCal('sSD');"  style="cursor:hand;" value="<%=dSDay%>"  class="input">
				             <span id="spST" style="display:<%if sSType="1" then%>none<%end if%>">
				                    <select name="sSTi" id="sSTi" class="select" <%if sSType =1 then%>disabled<%end if%>>
				                        <%For intTime=0 To 23%>
				                        <option value="<%=intTime%>" <%if Cint(sSTime)=intTime then%>selected<%end if%>><%=intTime%></option>
				                        <%Next%>
				                    </select>
				                   <input type="text" name="sSTSec" id="sSTSec" readonly value="00:00" size="5" class="text_ro">
				              </span>
				            <%END IF%>
				~ 종료일 : <%IF eCode <> "" THEN %><%=dEDay%><input type="hidden" name="sED" value="<%=dEDay%>">
				            <%ELSE%><input type="text" name="sED"  size="10" onClick="jsPopCal('sED');" style="cursor:hand;" value="<%=dEDay%>"  class="input">
				             <span id="spET" style="display:<%if sSType="1" then%>none<%end if%>">
				                    <select name="sETi" id="sETi" class="select" <%if sSType =1 then%>disabled<%end if%>>
				                        <%For intTime=0 To 23%>
				                        <option value="<%=intTime%>" <%if Cint(sETime)=intTime then%>selected<%end if%>><%=intTime%></option>
				                        <%Next%>
				                    </select>
				                   <input type="text" name="sETSec" id="sETSec" readonly value="59:59" size="5" class="text_ro">
				              </span>
				            <%END IF%>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 할인율</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iSR" size="4"  value="<%=isRate%>" style="text-align:right;"  class="input">%</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">마진구분</td>
			<td bgcolor="#FFFFFF"><%sbGetOptCommonCodeArr "salemargin", isMargin, False,True,"onchange='jsChSetValue(this.value);'"%>
			<span id="divM" style="display:<%IF isMargin<> 5 THEN %>none<%END IF%>;">할인마진<input type="text" size="4" name="isMV" maxlength="10" value="<%=isMValue%>" style="text-align:right;">%</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center"> 상태</td>
			<td bgcolor="#FFFFFF" <%if isStatus <>0 and not C_ADMIN_AUTH then%>colspan="3"<%end if%>>
					<input type="hidden" name="sOD" value="<%=dOpenDay%>">
					<input type="hidden" name="salestatus" value="<%=isStatus%>">
					<%=fnGetCommCodeArrDesc(arrsalestatus,isStatus)%>
				<%if eCode = "" then%>
				<%IF isStatus =0 then '등록대기 %>
					<input type="checkbox" name="chkstatus" value="7">오픈요청
				<%elseif isStatus = 6 or isStatus = 7 then '오픈 %>
					<input type="checkbox" name="chkstatus" value="9">종료요청
				<%elseif isStatus = 8 then %>
					<div style="padding-top:5px;">종료일: <%=dCloseDay%></div>
				<%end if%>
				<%end if%>
			</td>
			<%if isStatus =0 or C_ADMIN_AUTH then%>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용유무</td>
			<td bgcolor="#FFFFFF">
				<input type="radio" name="sSU" value="1" <%IF isUsing THEN%>checked<%END IF%>>사용 <input type="radio" name="sSU" value="0" <%IF not isUsing  THEN%>checked<%END IF%>>사용안함
			</td>
			<%end if %>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<input type="image" src="/images/icon_save.gif">
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
<%IF sCode <> "" THEN%>
<tr>
	<td>
		<form name="frmSearch" method="get" action="">
			<input type=hidden name=menupos value="<%=menupos%>">
			<input type=hidden name=sC value="<%=sCode%>">
			<input type=hidden name=eC value="<%=eCode%>">
			<input type=hidden name=iC value="<%=iCurrpage%>">
			<input type="hidden" name="iSR" value="<%=isRate%>">
			<input type="hidden" name="salemargin" value="<%=ismargin%>">
			<input type="hidden" name="isMValue" value="<%=isMValue%>">
			
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" bgcolor="#EEEEEE" align="center">검색조건</td>
				<td bgcolor="#ffffff">
					<table   border="0"  cellpadding="3" cellspacing="1" class="a">
					<tr>
						<td width="300"> 브랜드:
					   	<% drawSelectBoxDesignerWithName "makerid",makerid %>
						</td>
							<td>&nbsp;&nbsp;상품 할인:
								 <% drawSelectBoxSailYN "sailyn", sailyn %>
							</td>
							<td>상품코드:</td>
						<td rowspan="2" bgcolor="#FFFFFF"><textarea name="sRectItemidArr" rows="3" cols="10"><%=replace(sRectItemidArr,",",chr(10))%></textarea> </td>
					</tr>
					<tr>
						<td  bgcolor="#FFFFFF">
					      <input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >마진부족(or 역마진) 상품 검색
				    </td>
				    <td   >&nbsp;&nbsp;(할인코드)마스터할인:
						<select name="salestatus" class="select" >
						<option value="">전체</option>
						<option value="0"  <%if sSalestatus ="0" then%>selected<%end if%>>등록대기</option>
						<option value="7"  <%if sSalestatus ="7" then%>selected<%end if%>>할인예정</option>
						<option value="6"  <%if sSalestatus ="6" then%>selected<%end if%>>할인중</option>
						<option value="9"  <%if sSalestatus ="9" then%>selected<%end if%>>할인중(종료예정)</option>
						<option value="8"  <%if sSalestatus ="8" then%>selected<%end if%>>종료</option>
						</select>
			 			&nbsp;&nbsp;
						(할인코드)상품할인:
						<select name="selItemStatus" class="select"> <!--// 6-오픈, 7-오픈요청, 8-종료,9-종료요청-->
							<option value="">전체</option>
							<option value="7" <%if sItemSale ="7" then%>selected<%end if%>>할인예정</option>
							<option value="6" <%if sItemSale ="6" then%>selected<%end if%>>할인중</option>
							<option value="9" <%if sItemSale ="9" then%>selected<%end if%>>할인중(종료예정)</option>
							<option value="8" <%if sItemSale ="8" then%>selected<%end if%>>할인종료</option>
						</select>
					</td>
				</table>
				</td>
				<td  width="120" bgcolor="#EEEEEE" align="center">
					 <input type="button" class="button" value="등록된 상품 검색" onclick="jsSearch();">
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border=0>
		<form name=frmdummi>
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40" valign="bottom">
			<td align="left"><input type=button value="선택상품수정" onClick="saveArr()" class="button">
			<input type=button value="선택상품삭제" onClick="delArr()" class="button">
			</td>
			<td align="right">


				<!--<img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" style="vertical-align:bottom;cursor:pointer;" onclick="pop_upload('<%=eCode%>','<%=egCode%>');">-->
			&nbsp;&nbsp;
			대표할인율: <font color="blue"><%=isRate%>%</font>, 마진구분: <%=fnGetCommCodeArrDesc(arrsalemargin,isMargin)%><%IF isMargin = 5 THEN%>,&nbsp;대표할인마진율: <font color="blue"><%=isMValue%>%</font> <%END IF%>
			<input type="button" value="할인적용" onClick="CkDisPrice();" class="button">
			<input type="button" value="원 가격적용" onClick="CkOrgPrice();" class="button">
			&nbsp;&nbsp;<strong>|</strong>&nbsp;&nbsp;
				<select name="salemode" id="salemode" class="select">
					<option value="">-Choice-</option>
					<option value="1" selected>할인율</option>
					<option value="2">할인마진율</option>
				</select>
				<input type="text" id="salemodePer" name="salemodePer" onkeyup="lfn_keychk(this)" size="4" class="text">%
				<input type="button" value="선택적용" class="button" onclick="salemodeChk();">
			&nbsp;&nbsp;<strong>|</strong>&nbsp;&nbsp;
			<input type="button" value="상품 다운로드(엑셀)" class="button"  onclick="pop_down('<%=eCode%>','<%=egCode%>');">
			<strong>|</strong>
			<input type="button" value="상품 일괄등록(엑셀)" class="button"  onclick="pop_upload('<%=eCode%>','<%=egCode%>');">
			<input type="button" value="새상품 추가" onclick="addnewItem('<%=eCode%>','<%=egCode%>');" class="button">
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan="17" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iSubCurrpage%> / <%=iTotalPage%></b></td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td><input type="checkbox" name="ck_all" onclick="SelectCk(this)"></td>
				<td align="center">상품ID</td>
				<td align="center" >이미지</td>
				<td align="center">브랜드</td>
				<td align="center">상품명</td>
				<td align="center">계약<br>구분</td>
				<td align="center">할인상태</td>
				<td align="center">현재<br>판매가</td>
				<td align="center">현재<br>매입가</td>
				<td align="center">현재<br>마진율</td>

				<td align="center">원<br>소비자가</td>
				<td align="center">원<br>매입가</td>
				<td align="center">원<br>마진율</td>

				<td align="center">소비자가대비<br>할인율</td>
				<td align="center">할인<br>판매가</td>
				<td align="center">할인<br>매입가</td>
				<td align="center">할인<br>마진율</td>
		</tr>
		<%	Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin, iSalePercent
			Dim cpSP, cpSB, cpSM, strCpDesc, strCpList
			dim mOrgSailPrice, mOrgSailSuplyCash, sOrgSailYn, iOrgSailMargin
			Dim saleBuycashErr, saleBuycashErrExists : saleBuycashErrExists=false
			iSaleMargin=0
			iOrgMargin = 0
			iOrgSailMargin= 0
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			saleBuycashErr = arrList(3,intLoop)>arrList(14,intLoop)  ''2018/07/20
			if (saleBuycashErr) then saleBuycashErrExists=True
			
			mSPrice  =arrList(13,intLoop) - (arrList(13,intLoop)*(isRate/100))
			mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,arrList(13,intLoop),arrList(14,intLoop),mSPrice)

			if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
			 if arrList(13,intLoop)<>0 then
			 	iOrgMargin= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100
			 	iSalePercent = ((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop))*100
			 end if

			'기존 할인상품일 경우 기존 할인가격 가져오기
			'기존 할인상품일 경우에도 할인율 및 할인판매가.매입가 계산은 원소비자가격으로 한다
			sOrgSailYn = arrList(24,intLoop)
			mOrgSailPrice = arrList(22,intLoop)
			mOrgSailSuplycash = arrList(23,intLoop)
			if mOrgSailPrice <>0 then
			 	iOrgSailMargin= 100-fix(mOrgSailSuplycash/mOrgSailPrice*10000)/100
			 end if

			cpSP=0: cpSB=0: cpSM=0: strCpDesc="": strCpList=""
			if isArray(arrItemCoupon) then

				for icLp=0 to ubound(arrItemCoupon,2)
					if cStr(arrItemCoupon(4,icLp))=cStr(arrList(1,intLoop)) then
						'상품쿠폰판매가
						Select Case arrItemCoupon(1,icLp)
							Case "1"
								cpSP = mSPrice- CLng(arrItemCoupon(2,icLp)*mSPrice/100)
							Case "2"
								cpSP = mSPrice- arrItemCoupon(2,icLp)
							Case Else
								cpSP = mSPrice
						End Select
						'상품쿠폰매입가
						cpSB = arrItemCoupon(5,icLp)
						'상품쿠폰마진
						if cpSB>0 then cpSM = formatNumber(100-fix(cpSB/cpSP*10000)/100,0)

						strCpList = strCpList & "<tr align='center' onclick=""window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&iSerachType=1&sSearchTxt=" & arrItemCoupon(0,icLp) & "')"">" &_
								"<td>[" & arrItemCoupon(0,icLp) & "]</td>" &_
								"<td>" & arrItemCoupon(3,icLp) & "</td>" &_
								"<td>" & FormatNumber(cpSP,0) & "원</td>" &_
								"<td>" & FormatNumber(cpSB,0) & "원</td>" &_
								"<td " & chkIIF(cpSM<=5,"style='color:#ee0000;'","") & ">" & FormatNumber(cpSM,0) & "%</td>" &_
								"<td>" & left(arrItemCoupon(6,icLp),10) & "</td>" &_
								"<td>" & left(arrItemCoupon(7,icLp),10) & "</td>" &_
								"</tr>"
					end if
				next

				if strCpList<>"" then
					strCpDesc = "<div><font color=darkgreen style='cursor:pointer;' onmouseover=""$(this).find('div').show()"" onmouseout=""$(this).find('div').hide()"">상품쿠폰 ▶" &_
							"<div style='display:none;position:absolute;border:1px solid #C0C0C0;padding:5px;background-color:#FFFFFF;margin:-10px -20px;'>" &_
							"<table width='600' border='0' cellpadding='3' cellspacing='1' class='a'>" &_
							"<tr><td colspan='7' align='left'><strong>할인기간중 진행되는 쿠폰</strong></td></tr>" &_
							"<tr align='center' bgcolor='#F0F0F0'>" &_
							"<td colspan='2'>쿠폰명</td>" &_
							"<td>쿠폰할인가</td>" &_
							"<td>쿠폰매입가</td>" &_
							"<td>쿠폰할인마진</td>" &_
							"<td>시작일</td>" &_
							"<td>종료일</td>" &_
							"</tr>" &_
							strCpList &_
							"</table>" &_
							"</div></font></div>"
				end if

			end if
			
			%>
			<form name="frmBuyPrc_<%=arrList(1,intLoop)%>" >
			<input type=hidden name="itemid" value="<%=arrList(1,intLoop)%>">
			<input type=hidden name="saleprice" value="<%=mSPrice%>">
			<input type=hidden name="salesupplyprice" value="<%=mSBPrice%>">
			<input type=hidden name="salemargin" value="<%=iSaleMargin%>">

			<input type=hidden name="orgPrice" value="<%=arrList(13,intLoop)%>">
			<input type=hidden name="orgSupplyPrice" value="<%=arrList(14,intLoop)%>">
			<input type=hidden name="orgMarginValue" value="<%=iOrgMargin%>">

			<input type=hidden name="orgSailPrice" value="<%=mOrgSailPrice%>">
			<input type=hidden name="orgSailSupplyPrice" value="<%=mOrgSailSuplycash%>">
			<input type=hidden name="orgSailMarginValue" value="<%=iOrgSailMargin%>">
			<input type="hidden" name="orgSailYn" value="<%=sOrgSailYn%>">
			<input type="hidden" name="mwdiv" value="<%=arrList(17,intLoop)%>">

			<input type=hidden name="saleItemStatus" value="<%=arrList(4,intLoop)%>">
		 <tr align="center" bgcolor=<%IF cint(arrList(4,intLoop)) = 8 THEN%>"#B3B3B3"<%ELSE%>"#FFFFFF"<%END IF%>>
			    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			    <td><a href="javascript:jsGoPreItem('<%=wwwURL%>','<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
			    <td><%IF arrList(9,intLoop) <> "" THEN%><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(1,intLoop))%>/<%=arrList(9,intLoop)%>"><%END IF%></td>
			    <td><%=db2html(arrList(7,intLoop))%></td>
			    <td align="left">&nbsp;<%=db2html(arrList(8,intLoop))%></td>
			    <td><%= fnColor(arrList(17,intLoop),"mw") %></td>
			    <td>
			    	<%= fnColor(arrList(10,intLoop),"yn") %>&nbsp;<%IF arrList(4,intLoop) = 6 THEN%><font color="blue"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(4,intLoop))%>
			    	<%=chkIIF(strCpDesc>"",strCpDesc,"")%>
			    </td>

			    <td><%if arrList(10,intLoop)="Y" then%><font color="red"><%end if%><%=formatnumber(arrList(11,intLoop),0)%></td>
			    <td><%if arrList(10,intLoop)="Y" then%><font color="red"><%end if%><%=formatnumber(arrList(12,intLoop),0)%></td>
			    <td><%if arrList(10,intLoop)="Y" then%><font color="red"><%end if%><% if arrList(11,intLoop)<>0 then %>
					<%= 100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 %>%
					<% end if %>
				</td>

			    <td>
			    	<%=formatnumber(arrList(13,intLoop),0)%>
			    	<%if sOrgSailYn ="Y" then%>
			    	<br/><font color="#F08050">(<%=formatnumber((arrList(13,intLoop)-mOrgSailPrice)/arrList(13,intLoop)*100,0) %>%할)<%=formatnumber(mOrgSailPrice,0)%></font>
			    	<%end if%>
			    </td>
			    <td <%=CHKIIF(saleBuycashErr,"bgcolor='#CCCC66'","")%>>
			        <%=CHKIIF(saleBuycashErr,"<strong>","")%>
			    	<%=formatnumber(arrList(14,intLoop),0)%>
			    <%=CHKIIF(saleBuycashErr,"</strong>","")%>
			    	<%if sOrgSailYn ="Y" then%>
			    	<br/><font color="#F08050"><%=formatnumber(mOrgSailSuplycash,0)%></font>
			    	<%end if%>
			    </td>
			    <td>
			    	<%=iOrgMargin%>%
			    	<%if sOrgSailYn ="Y" then%>
			    	<br/><font color="#F08050"><%=formatnumber(iOrgSailMargin,0)%>%</font>
			    	<%end if%>
			    </td>

				<!-- <td id="lyrSpct<%=arrList(1,intLoop)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=formatnumber(iSalePercent,0)%>%</td> -->
				<td id="lyrSpct<%=arrList(1,intLoop)%>" style="<%=chkIIF(iSalePercent>=50,"color:#EE0000;font-weight:bold;","")%>"><%=numBerBurim(iSalePercent, 2)%>%</td>
			<%IF cint(arrList(4,intLoop)) = 8 or  cint(arrList(4,intLoop)) = 9 THEN%>
				<td><input type="text" name="iDSPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(2,intLoop)%></td>
			    <td <%=CHKIIF(saleBuycashErr,"bgcolor='#CCCC66'","")%>><input type="text" name="iDBPrice" size="6" maxlength="9" value="0" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"><br><%=arrList(3,intLoop)%></td>
			    <td><input type="text" name="iDSMargin" value="0" style="text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%<br><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%></td>
			<%ELSE%>
			    <td><input type="text" name="iDSPrice" size="6" maxlength="9" value="<%=arrList(2,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td <%=CHKIIF(saleBuycashErr,"bgcolor='#CCCC66'","")%>><input type="text" name="iDBPrice" size="6" maxlength="9" value="<%=arrList(3,intLoop)%>" style="text-align:right;" onkeyup="reCALbyPrice('<%=arrList(1,intLoop)%>')"></td>
			    <td><%  if arrList(2,intLoop)<>0 then smargin= 100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100 	%>
					<input type="text" name="iDSMargin" value="<%=smargin%>" style=text-align:right;" size="4" onkeyup="reCALbyMargin('<%=arrList(1,intLoop)%>')">%
			    </td>
			<%END IF%>
		</tr>
		</form>
		<%	next %>
		<%
		if (saleBuycashErrExists) then
		    response.write "<script>$(function(){alert('원 매입가보다 할인매입가가 큰 상품이 존재합니다.')});</script>"
		end if
		%>
		<% else %>
		<tr>
			<td colspan="17" bgcolor="#ffffff" align="center">등록된 내역이 없습니다.</td>
		</tr>
		<%
		END IF%>
		<!-- 페이징처리 -->
			<%
			iStartPage = (Int((iSubCurrpage-1)/iPerCnt)*iPerCnt) + 1

			If (iSubCurrpage mod iPerCnt) = 0 Then
				iEndPage = iSubCurrpage
			Else
				iEndPage = iStartPage + (iPerCnt-1)
			End If
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iSubCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			        <td  width="50" align="right"><a href="saleList.asp?menupos=<%=menupos%>"><img src="/images/icon_list.gif" border="0"></a></td>
			    </tr>
		</table>
	</td>
</tr>
<form name="frmarr" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sailyn" value="">
<input type="hidden" name="iDSPrice" value="">
<input type="hidden" name="iDBPrice" value="">
<input type="hidden" name="saleItemStatus" value="">
<input type="hidden" name="saleStatus" value="<%=isStatus%>">
</form>
<form name="frmdel" method="post" action="saleItemPRoc.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sC" value="<%=sCode%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="sailyn" value="">
<input type="hidden" name="iDSPrice" value="">
<input type="hidden" name="iDBPrice" value="">
<input type="hidden" name="saleItemStatus" value="">
<input type="hidden" name="saleStatus" value="<%=isStatus%>">
</form>
<%end if%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
