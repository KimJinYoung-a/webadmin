<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wconcept/wconceptcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, wconceptKeepSell, isSpecialPrice
Dim bestOrdMall, wconceptGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchBrand, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, wconceptYes10x10No, wconceptNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, wconceptGoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim oWconcept, xl, kjypageSize, purchasetype
Dim startMargin, endMargin
page    				= request("page")
kjypageSize				= request("kjypageSize")
research				= request("research")
itemid  				= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
startMargin				= request("startMargin")
endMargin				= request("endMargin")
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")
infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
wconceptGoodNo			= request("wconceptGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchBrand				= request("MatchBrand")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
wconceptYes10x10No		= request("wconceptYes10x10No")
wconceptNo10x10Yes		= request("wconceptNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
purchasetype 			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchBrand= ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = ""
End If

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

'wconcept 상품코드 엔터키로 검색되게
If wconceptGoodNo <> "" then
	Dim iA2, arrTemp2, arrwconceptGoodNo
	wconceptGoodNo = replace(wconceptGoodNo,",",chr(10))
	wconceptGoodNo = replace(wconceptGoodNo,chr(13),"")
	arrTemp2 = Split(wconceptGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrwconceptGoodNo = arrwconceptGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	wconceptGoodNo = left(arrwconceptGoodNo,len(arrwconceptGoodNo)-1)
End If

Set oWconcept = new Cwconcept
	oWconcept.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oWconcept.FPageSize					= kjypageSize
Else
	oWconcept.FPageSize					= 100
End If
	oWconcept.FRectCDL					= request("cdl")
	oWconcept.FRectCDM					= request("cdm")
	oWconcept.FRectCDS					= request("cds")
	oWconcept.FRectItemID				= itemid
	oWconcept.FRectItemName				= itemname
	oWconcept.FRectSellYn				= sellyn
	oWconcept.FRectLimitYn				= limityn
	oWconcept.FRectSailYn				= sailyn
	oWconcept.FRectStartMargin			= startMargin
	oWconcept.FRectEndMargin			= endMargin
	oWconcept.FRectMakerid				= makerid
	oWconcept.FRectwconceptGoodNo		= wconceptGoodNo
	oWconcept.FRectMatchCate			= MatchCate
	oWconcept.FRectMatchBrand			= MatchBrand
	oWconcept.FRectIsMadeHand			= isMadeHand
	oWconcept.FRectIsOption				= isOption
	oWconcept.FRectIsReged				= isReged
	oWconcept.FRectNotinmakerid			= notinmakerid
	oWconcept.FRectNotinitemid			= notinitemid
	oWconcept.FRectExcTrans				= exctrans
	oWconcept.FRectPriceOption			= priceOption
	oWconcept.FRectIsSpecialPrice     	= isSpecialPrice
	oWconcept.FRectDeliverytype			= deliverytype
	oWconcept.FRectMwdiv				= mwdiv
	oWconcept.FRectScheduleNotInItemid	= scheduleNotInItemid
	oWconcept.FRectIsextusing			= isextusing
	oWconcept.FRectCisextusing			= cisextusing
	oWconcept.FRectRctsellcnt			= rctsellcnt
	oWconcept.FRectPurchasetype 		= purchasetype

	oWconcept.FRectExtNotReg			= ExtNotReg
	oWconcept.FRectExpensive10x10		= expensive10x10
	oWconcept.FRectdiffPrc				= diffPrc
	oWconcept.FRectwconceptYes10x10No	= wconceptYes10x10No
	oWconcept.FRectwconceptNo10x10Yes	= wconceptNo10x10Yes
	oWconcept.FRectwconceptKeepSell		= wconceptKeepSell
	oWconcept.FRectExtSellYn			= extsellyn
	oWconcept.FRectInfoDiv				= infoDiv
	oWconcept.FRectFailCntOverExcept	= ""
	oWconcept.FRectFailCntExists		= failCntExists
	oWconcept.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oWconcept.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oWconcept.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oWconcept.getwconceptreqExpireItemList
Else
	oWconcept.getwconceptRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=wconceptList"& replace(DATE(), "-", "") &"_xl.xls"
Else
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
//크롬 업데이트로 alert 수정..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//크롬 업데이트로 alert 수정..2021-07-26 끝

// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=wconcept1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=wconcept1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=wconcept1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=wconcept1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function onlyJY(comp){
     if ((comp.name=="morningJY")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=true;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.value="D"
			comp.form.ExtNotReg.disabled = true;
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";
    	}
    }

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="wconceptKeepSell")&&(frm.wconceptKeepSell.checked)){ frm.wconceptKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="wconceptYes10x10No")&&(frm.wconceptYes10x10No.checked)){ frm.wconceptYes10x10No.checked=false }
	if ((comp.name!="wconceptNo10x10Yes")&&(frm.wconceptNo10x10Yes.checked)){ frm.wconceptNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "D"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "D"
   			comp.form.ExtNotReg.disabled = true;
			comp.form.extsellyn.value = "N";
			comp.form.sellyn.value = "Y";
   		}else{
			if (document.getElementById("NR").checked == false){
				comp.form.extsellyn.value = "Y";
			}else{
				comp.form.extsellyn.value = "";
				comp.form.sellyn.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="wconceptYes10x10No")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.isReged.checked = true;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="wconceptNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
			comp.form.notinmakerid.value = "";
			comp.form.notinitemid.value = "";
			comp.form.exctrans.value = "N";
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.wconceptYes10x10No.checked){
            comp.form.wconceptYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="";
	        comp.form.extsellyn.value = "Y";
    	}
    }
	if ((comp.name=="diffPrc")){
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
        }
	}

	if (comp.name=="reqEdit"){
		if (comp.checked){
			document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="wconceptYes10x10No")&&(frm.wconceptYes10x10No.checked)){ frm.wconceptYes10x10No.checked=false }
	if ((comp.name!="wconceptNo10x10Yes")&&(frm.wconceptNo10x10Yes.checked)){ frm.wconceptNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//등록여부 조건 Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/wconcept/popwconceptCateList.asp","popCatewconceptmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

//공통코드 검색
function fnwconceptCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;

	if(ccd == ''){
		alert('공통코드를 입력하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		xLink.location.href = "/admin/etc/wconcept/actwconceptReq.asp?cmdparam=wconceptCommonCode&CommCD="+ccd+"";
	}
}

// 선택된 상품 등록
function wconceptSelectRegProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('wconcept에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

function wconceptSelectRegStepProcess(v) {
	var chkSel=0;
	var strStep;
	switch(v) {
		case 1 :
			strStep="REGSTEP1";
			break;
		case 2 :
			strStep="REGSTEP2";
			break;
		case 3 :
			strStep="REGSTEP3";
			break;
		case 4 :
			strStep="REGSTEP4";
			break;
		case 5 :
			strStep="REGSTEP5";
			break;
		case 6 :
			strStep="REGSTEP6";
			break;

	}
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('wconcept에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = strStep;
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function wconceptSellYnProcess(chkYn) {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="판매";break;
		case "N": strSell="판매중지";break;
		case "X": strSell="승인전 상품삭제";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
       	document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function wconceptEditProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function wconceptSelectPriceEditProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 가격 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//컨텐츠 수정
function wconceptSelectContentEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품 컨텐츠를 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//이미지 수정
function wconceptSelectImageEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품 이미지를 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//추가 이미지 수정
function wconceptSelectAddImageEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품 추가이미지를 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ADDIMAGE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//옵션 수정
function wconceptSelectOptionEditProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품 옵션을 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//정보고시 수정
function wconceptSelectInfoCodeProcess(){
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품 옵션을 수정 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "INFOCODE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}


//상품 + 옵션 조회
function wconceptChkStatProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}

//상품 조회
function wconceptChkItemProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKITEM";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}

//옵션 조회
function wconceptChkOptProcess() {
	var chkSel=0;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("선택한 상품이 없습니다.");
			return;
		}
	}
	catch(e) {
		alert("상품이 없습니다.");
		return;
	}

    if (confirm('선택하신 ' + chkSel + '개 옵션을 조회 하시겠습니까?\n\n※wconcept와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKOPT";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}


//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=wconcept1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function wconceptEditProcessAuto() {
	var cnt = <%= oWconcept.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oWconcept.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		wconceptEditProcessAuto();
		// 5분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 5*60*1000);
	}
}

$(document).ready(function() {
    $('table').hide();
});
<% end if %>
function popXL()
{
    frmXL.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="https://newpin.wconcept.co.kr/Auth/Login" target="_blank">wconceptAdmin바로가기</a>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		wconcept 상품코드 : <textarea rows="2" cols="20" name="wconceptGoodNo" id="itemid"><%=replace(wconceptGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >wconcept 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >wconcept 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >wconcept 등록완료(전시)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">품절처리요망</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">등록상품 판매가능</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">등록여부조건Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
		<% If (session("ssBctID")="kjy8517") Then %>
			<input class="text" size="5" type="text" name="kjypageSize" value="<%= kjypageSize %>">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>wconcept 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptYes10x10No" <%= ChkIIF(wconceptYes10x10No="on","checked","") %> ><font color=red>wconcept판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptNo10x10Yes" <%= ChkIIF(wconceptNo10x10Yes="on","checked","") %> ><font color=red>wconcept품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptKeepSell" <%= ChkIIF(wconceptKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 배송방법이 택배(일반) 아닌것, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 옵션가 0원 판매중 상품이 없음(옵션 한정수량 5개 이하 포함), 옵션가가 판매가 100% 이상인 상품

<p />
<% end if %>
<form name="frmReg" method="post" action="lotteItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('wconcept1010');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="카테고리" onclick="pop_CateManager();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
				실제상품 등록 :
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="wconceptSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP1)" onClick="wconceptSelectRegStepProcess(1);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP2)" onClick="wconceptSelectRegStepProcess(2);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP3)" onClick="wconceptSelectRegStepProcess(3);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP4)" onClick="wconceptSelectRegStepProcess(4);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP5)" onClick="wconceptSelectRegStepProcess(5);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP6)" onClick="wconceptSelectRegStepProcess(6);">&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" value="수정" onClick="wconceptEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="가격" onClick="wconceptSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="컨텐츠" onClick="wconceptSelectContentEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지" onClick="wconceptSelectImageEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="추가이미지" onClick="wconceptSelectAddImageEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="wconceptSelectOptionEditProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" value="고시" onClick="wconceptSelectInfoCodeProcess();">&nbsp;&nbsp;
			<% End If %>
				<br><br>
				실제상품 조회 :
				<input class="button" type="button" value="조회" onClick="wconceptChkStatProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="상품" onClick="wconceptChkItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="wconceptChkOptProcess();">
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				공통코드 검색 :
				<Select name="CommCD" id="CommCD" class="select">
					<option value="BRAND">브랜드</option>
					<option value="ORIGIN">원산지</option>
					<option value="CATEGORY">카테고리</option>
					<% If (session("ssBctID")="kjy8517") Then %>
					<option value="PRODUCT_NOTICE_INFO">카테고리품목</option>
					<% End If %>
				</Select>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="fnwconceptCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="wconceptSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= FormatNumber(oWconcept.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oWconcept.FTotalPage,0) %></b>
	</td>
	<td align="right">
		<input type="button" class="button" value="엑셀받기" onClick="popXL()">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
<% If (xl <> "Y") Then %>
	<td width="50">이미지</td>
<% End If %>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">wconcept등록일<br>wconcept최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">wconcept<br>가격및판매</td>
	<td width="70">wconcept<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oWconcept.FResultCount - 1 %>
<tr align="center" <% If oWconcept.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oWconcept.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oWconcept.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oWconcept.FItemList(i).FItemID %>','wconcept1010','<%=oWconcept.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oWconcept.FItemList(i).FItemID%>" target="_blank"><%= oWconcept.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oWconcept.FItemList(i).FwconceptStatCd <> 7 Then
	%>
		<br><%= oWconcept.FItemList(i).getwconceptStatName %>
	<%
			End If
			response.write oWconcept.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oWconcept.FItemList(i).FMakerid %> <%= oWconcept.FItemList(i).getDeliverytypeName %><br><%= oWconcept.FItemList(i).FItemName %></td>
	<td align="center"><%= oWconcept.FItemList(i).FRegdate %><br><%= oWconcept.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oWconcept.FItemList(i).FwconceptRegdate %><br><%= oWconcept.FItemList(i).FwconceptLastUpdate %></td>
	<td align="right">
		<% If oWconcept.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oWconcept.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oWconcept.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oWconcept.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oWconcept.FItemList(i).Fsellcash = 0 Then
		elseif (oWconcept.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oWconcept.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oWconcept.FItemList(i).FOrgSuplycash/oWconcept.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oWconcept.FItemList(i).Fbuycash/oWconcept.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oWconcept.FItemList(i).Fbuycash/oWconcept.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oWconcept.FItemList(i).IsSoldOut Then
			If oWconcept.FItemList(i).FSellyn = "N" Then
	%>
		<font color="red">품절</font>
	<%
			Else
	%>
		<font color="red">일시<br>품절</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If oWconcept.FItemList(i).FItemdiv = "06" OR oWconcept.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oWconcept.FItemList(i).FwconceptStatCd > 0) Then
			If Not IsNULL(oWconcept.FItemList(i).FwconceptPrice) Then
				If (oWconcept.FItemList(i).Mustprice <> oWconcept.FItemList(i).FwconceptPrice) Then
	%>
					<strong><%= formatNumber(oWconcept.FItemList(i).FwconceptPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oWconcept.FItemList(i).FwconceptPrice,0)&"<br>"
				End If

				If Not IsNULL(oWconcept.FItemList(i).FSpecialPrice) Then
					If (now() >= oWconcept.FItemList(i).FStartDate) And (now() <= oWconcept.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oWconcept.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oWconcept.FItemList(i).FSellyn="Y" and oWconcept.FItemList(i).FwconceptSellYn<>"Y") or (oWconcept.FItemList(i).FSellyn<>"Y" and oWconcept.FItemList(i).FwconceptSellYn="Y") Then
	%>
					<strong><%= oWconcept.FItemList(i).FwconceptSellYn %></strong>
	<%
				Else
					response.write oWconcept.FItemList(i).FwconceptSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oWconcept.FItemList(i).FwconceptGoodNo <> "" Then %>
			<a target="_blank" href="https://www.wconcept.co.kr/product/<%=oWconcept.FItemList(i).FwconceptGoodNo%>"><%=oWconcept.FItemList(i).FwconceptGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oWconcept.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oWconcept.FItemList(i).FItemID%>','0');"><%= oWconcept.FItemList(i).FoptionCnt %>:<%= oWconcept.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oWconcept.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= Chkiif(oWconcept.FItemList(i).FCateMapCnt > 0, "매칭됨(카)", "<font color='darkred'>매칭안됨(카)</font>") %><br />
		<%= Chkiif(oWconcept.FItemList(i).FBrandMapCnt > 0, "매칭됨(브)", "<font color='darkred'>매칭안됨(브)</font>") %><br />
	</td>
	<td align="center">
		<%= oWconcept.FItemList(i).FinfoDiv %>
		<%
		If (oWconcept.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oWconcept.FItemList(i).FlastErrStr) &"'>ERR:"& oWconcept.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% wconceptGoodNoArray = wconceptGoodNoArray & oWconcept.FItemList(i).FwconceptGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= wconceptGoodNoArray %></textarea>
	<button onclick="copyArr();">Copy</button>
	<script>
	function copyArr() {
		var tt = document.getElementById("itemidArr");
		tt.select();
		document.execCommand("copy");
	}
	</script>
<% End If %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oWconcept.HasPreScroll then %>
		<a href="javascript:goPage('<%= oWconcept.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oWconcept.StartScrollPage to oWconcept.FScrollCount + oWconcept.StartScrollPage - 1 %>
    		<% if i>oWconcept.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oWconcept.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value= <%= page %>>
	<input type="hidden" name="research" value= <%= research %>>
	<input type="hidden" name="itemid" value= <%= itemid %>>
	<input type="hidden" name="makerid" value= <%= makerid %>>
	<input type="hidden" name="itemname" value= <%= itemname %>>
	<input type="hidden" name="bestOrd" value= <%= bestOrd %>>
	<input type="hidden" name="bestOrdMall" value= <%= bestOrdMall %>>
	<input type="hidden" name="sellyn" value= <%= sellyn %>>
	<input type="hidden" name="limityn" value= <%= limityn %>>
	<input type="hidden" name="sailyn" value= <%= sailyn %>>
	<input type="hidden" name="onlyValidMargin" value= <%= onlyValidMargin %>>
	<input type="hidden" name="startMargin" value= <%= startMargin %>>
	<input type="hidden" name="endMargin" value= <%= endMargin %>>
	<input type="hidden" name="isMadeHand" value= <%= isMadeHand %>>
	<input type="hidden" name="isOption" value= <%= isOption %>>
	<input type="hidden" name="infoDiv" value= <%= infoDiv %>>
	<input type="hidden" name="morningJY" value= <%= morningJY %>>
	<input type="hidden" name="extsellyn" value= <%= extsellyn %>>
	<input type="hidden" name="wconceptGoodNo" value= <%= wconceptGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchBrand" value= <%= MatchBrand %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="wconceptYes10x10No" value= <%= wconceptYes10x10No %>>
	<input type="hidden" name="wconceptNo10x10Yes" value= <%= wconceptNo10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oWconcept = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
