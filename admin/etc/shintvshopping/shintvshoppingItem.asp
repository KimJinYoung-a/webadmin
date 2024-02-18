<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/shintvshopping/shintvshoppingCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, shintvshoppingGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, shintvshoppingYes10x10No, shintvshoppingNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, isextusing, cisextusing, rctsellcnt
Dim page, i, research, reglevel
Dim oShintvshopping, xl, shintvshoppingGoodNoArray, scheduleNotInItemid
Dim startMargin, endMargin
Dim purchasetype
page    				= request("page")
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
shintvshoppingGoodNo	= request("shintvshoppingGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
shintvshoppingYes10x10No = request("shintvshoppingYes10x10No")
shintvshoppingNo10x10Yes = request("shintvshoppingNo10x10Yes")
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
reglevel				= request("reglevel")
purchasetype			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
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

'shintvshopping 상품코드 엔터키로 검색되게
If shintvshoppingGoodNo <> "" then
	Dim iA2, arrTemp2, arrshintvshoppingGoodNo
	shintvshoppingGoodNo = replace(shintvshoppingGoodNo,",",chr(10))
	shintvshoppingGoodNo = replace(shintvshoppingGoodNo,chr(13),"")
	arrTemp2 = Split(shintvshoppingGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrshintvshoppingGoodNo = arrshintvshoppingGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	shintvshoppingGoodNo = left(arrshintvshoppingGoodNo,len(arrshintvshoppingGoodNo)-1)
End If

Set oShintvshopping = new CShintvshopping
	oShintvshopping.FCurrPage				= page
	oShintvshopping.FPageSize				= 100
	oShintvshopping.FRectCDL				= request("cdl")
	oShintvshopping.FRectCDM				= request("cdm")
	oShintvshopping.FRectCDS				= request("cds")
	oShintvshopping.FRectItemID				= itemid
	oShintvshopping.FRectItemName			= itemname
	oShintvshopping.FRectSellYn				= sellyn
	oShintvshopping.FRectLimitYn			= limityn
	oShintvshopping.FRectSailYn				= sailyn
	oShintvshopping.FRectStartMargin		= startMargin
	oShintvshopping.FRectEndMargin			= endMargin
	oShintvshopping.FRectMakerid			= makerid
	oShintvshopping.FRectshintvshoppingGoodNo		= shintvshoppingGoodNo
	oShintvshopping.FRectMatchCate			= MatchCate
	oShintvshopping.FRectIsMadeHand			= isMadeHand
	oShintvshopping.FRectIsOption			= isOption
	oShintvshopping.FRectIsReged			= isReged
	oShintvshopping.FRectNotinmakerid		= notinmakerid
	oShintvshopping.FRectNotinitemid		= notinitemid
	oShintvshopping.FRectExcTrans			= exctrans
	oShintvshopping.FRectPriceOption		= priceOption
	oShintvshopping.FRectIsSpecialPrice		= isSpecialPrice
	oShintvshopping.FRectDeliverytype		= deliverytype
	oShintvshopping.FRectMwdiv				= mwdiv
	oShintvshopping.FRectScheduleNotInItemid	= scheduleNotInItemid
	oShintvshopping.FRectIsextusing			= isextusing
	oShintvshopping.FRectCisextusing		= cisextusing
	oShintvshopping.FRectRctsellcnt			= rctsellcnt
	oShintvshopping.FRectReglevel			= reglevel


	oShintvshopping.FRectExtNotReg			= ExtNotReg
	oShintvshopping.FRectExpensive10x10		= expensive10x10
	oShintvshopping.FRectdiffPrc			= diffPrc
	oShintvshopping.FRectshintvshoppingYes10x10No	= shintvshoppingYes10x10No
	oShintvshopping.FRectshintvshoppingNo10x10Yes	= shintvshoppingNo10x10Yes
	oShintvshopping.FRectExtSellYn			= extsellyn
	oShintvshopping.FRectInfoDiv			= infoDiv
	oShintvshopping.FRectFailCntOverExcept	= ""
	oShintvshopping.FRectFailCntExists		= failCntExists
	oShintvshopping.FRectReqEdit			= reqEdit
	oShintvshopping.FRectPurchasetype		= purchasetype
If (bestOrd = "on") Then
    oShintvshopping.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oShintvshopping.FRectOrdType = "BM"
End If

If isReged = "R" Then					'품절처리요망 상품보기 리스트
	oShintvshopping.getshintvshoppingreqExpireItemList
Else
	oShintvshopping.getshintvshoppingRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=shintvshoppingList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=shintvshopping","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=shintvshopping','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=shintvshopping','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=shintvshopping','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 배송비 코드
function popDeliveryCode(){
	var popwin2=window.open('/admin/etc/shintvshopping/popShintvshoppingDeliveryCodeItem.asp','popDeliveryCode','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}

//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/shintvshopping/popshintvshoppingCateList.asp","popCateshintvshoppingmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
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
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="shintvshoppingYes10x10No")&&(frm.shintvshoppingYes10x10No.checked)){ frm.shintvshoppingYes10x10No.checked=false }
	if ((comp.name!="shintvshoppingNo10x10Yes")&&(frm.shintvshoppingNo10x10Yes.checked)){ frm.shintvshoppingNo10x10Yes.checked=false }
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

    if ((comp.name=="shintvshoppingYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="shintvshoppingNo10x10Yes")&&(comp.checked)){
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
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.shintvshoppingYes10x10No.checked){
            comp.form.shintvshoppingYes10x10No.checked = false;
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
	if ((comp.name!="shintvshoppingYes10x10No")&&(frm.shintvshoppingYes10x10No.checked)){ frm.shintvshoppingYes10x10No.checked=false }
	if ((comp.name!="shintvshoppingNo10x10Yes")&&(frm.shintvshoppingNo10x10Yes.checked)){ frm.shintvshoppingNo10x10Yes.checked=false }
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
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//상품 삭제
function shintvshoppingDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\n11번가 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=shintvshopping&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 선택된 상품 일괄 등록
function shintvshoppingSelectRegProcess() {
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

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 승인요청 + 전시조회
function shintvshoppingSelectConfirmProcess() {
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

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품을 승인요청 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CONFIRM";
		document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
		document.frmSvArr.submit();
    }
}

// 선택된 상품 수기 등록
function shintvshoppingSelectSugiRegProcess(v) {
	var chkSel=0;
	var strAct;
	var apiAct;
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

	switch(v) {
		case "1": strAct="기초";apiAct="REGAddItem";break;
		case "2": strAct="기술서";apiAct="REGContent";break;
		case "3": strAct="단품";apiAct="REGOpt";break;
		case "4": strAct="이미지";apiAct="REGImg";break;
		case "5": strAct="고시";apiAct="REGGosi";break;
		case "6": strAct="안전인증";apiAct="REGCert";break;
		case "7": strAct="승인요청";apiAct="REGConfirm";break;
	}

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품을 '+strAct+' 등록 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = apiAct;
		document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
		document.frmSvArr.submit();
    }
}

// 선택된 상품 수기 수정
function shintvshoppingSelectSugiEditProcess(v) {
	var chkSel=0;
	var strAct;
	var apiAct;
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

	switch(v) {
		case "1": strAct="기술서";apiAct="EDITContent";break;
		case "2": strAct="이미지";apiAct="EDITImage";break;
		case "3": strAct="고시";apiAct="EDITGosi";break;
		case "4": strAct="안전인증";apiAct="EDITCert";break;
	}

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품을 '+strAct+' 수정 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = apiAct;
		document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
		document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function shintvshoppingSellYnProcess(chkYn) {
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
		case "X": strSell="판매종료(영구정지)";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
        document.frmSvArr.submit();
    }
}

//수정
function shintvshoppingSelectEditProcess() {
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

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 수정 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
        document.frmSvArr.submit();
    }
}

//가격 수정
function shintvshoppingSelectPriceEditProcess() {
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

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 단품 가격을 수정 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
        document.frmSvArr.submit();
    }
}

//재고 수정
function shintvshoppingSelectQtyProcess() {
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

    if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품의 재고를 수정 하시겠습니까?\n\n※shintvshopping과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITQTY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
        document.frmSvArr.submit();
    }
}

//상품 상세 조회
function shintvshoppingSelectViewProcess() {
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

	if (confirm('shintvshopping에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/shintvshopping/actShintvshoppingReq.asp";
        document.frmSvArr.submit();
    }
}

function popXL()
{
    frmXL.submit();
}

function popPassword(){
    var popwin = window.open("/admin/etc/shintvshopping/popPassword.asp","popPassword","width=400,height=50,scrollbars=yes,resizable=yes");
	popwin.focus();
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
		<a href="https://spo.shinsegaetvshopping.com/app/index.html" target="_blank">신세계홈쇼핑 Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ E419803 | ten101010* ]</font>"
			End If
		%>
		<input type="button" class="button" value="비밀번호" onclick="popPassword();">
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		shintvshopping 상품코드 : <textarea rows="2" cols="20" name="shintvshoppingGoodNo" id="itemid"><%= replace(replace(shintvshoppingGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >shintvshopping 등록실패
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >shintvshopping 등록예정
			<option value="E" <%= CHkIIF(ExtNotReg="E","selected","") %> >shintvshopping 취소
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >shintvshopping 반려
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >shintvshopping 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >shintvshopping 등록완료(전시)
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
		등록레벨 :
		<select name="reglevel" class="select">
			<option value="">선택
			<% For i = 1 to 7 %>
				<option value="<%= i %>" <%= CHkIIF(CSTR(i) = CSTR(reglevel),"selected","") %> ><%= i %>
			<% Next %>
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>shintvshopping 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="shintvshoppingYes10x10No" <%= ChkIIF(shintvshoppingYes10x10No="on","checked","") %> ><font color=red>shintvshopping판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="shintvshoppingNo10x10Yes" <%= ChkIIF(shintvshoppingNo10x10Yes="on","checked","") %> ><font color=red>shintvshopping품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 할인이 아닐 때 판매가 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 주문제작문구 상품, 옵션추가금액 있는 상품<br />
* 등록레벨 : 기초정보등록(1), 기술서등록(2), 단품등록(3), 이미지URL(4), 정보고시(5), 안전인증(Optional)(6), 승인요청(7)<br />
<p />
<form name="frmReg" method="post" action="shintvshoppingItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();">&nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">&nbsp;
				<input class="button" type="button" value="배송비코드" onclick="popDeliveryCode();">&nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('shintvshopping');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="shintvshoppingSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="승인요청" onClick="shintvshoppingSelectConfirmProcess('7');" style=color:red;font-weight:bold>&nbsp;
				<input class="button" type="button" value="기초" onClick="shintvshoppingSelectSugiRegProcess('1');">&nbsp;
				<input class="button" type="button" value="기술서" onClick="shintvshoppingSelectSugiRegProcess('2');">&nbsp;
				<input class="button" type="button" value="단품" onClick="shintvshoppingSelectSugiRegProcess('3');">&nbsp;
				<input class="button" type="button" value="이미지" onClick="shintvshoppingSelectSugiRegProcess('4');">&nbsp;
				<input class="button" type="button" value="고시" onClick="shintvshoppingSelectSugiRegProcess('5');">&nbsp;
				<input class="button" type="button" value="안전인증" onClick="shintvshoppingSelectSugiRegProcess('6');">&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="등록(승인요청)_not 조회" onClick="shintvshoppingSelectSugiRegProcess('7');">&nbsp;
			<% End If %>	
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSelPrice" value="가격" onClick="shintvshoppingSelectPriceEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="수정" onClick="shintvshoppingSelectEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="조회" onClick="shintvshoppingSelectViewProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="기술서" onClick="shintvshoppingSelectSugiEditProcess('1');">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지" onClick="shintvshoppingSelectSugiEditProcess('2');">&nbsp;&nbsp;
				<input class="button" type="button" value="고시" onClick="shintvshoppingSelectSugiEditProcess('3');">&nbsp;&nbsp;
				<input class="button" type="button" value="안전인증" onClick="shintvshoppingSelectSugiEditProcess('4');">&nbsp;&nbsp;
				<input class="button" type="button" value="재고" onClick="shintvshoppingSelectQtyProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<!-- <option value="Y">판매중</option> -->
					<option value="X">판매종료(영구정지))</option><!-- 영구정지하면 이후 수정 할 수 없음 / 삭제 재등록이 안 된다고 함 -->
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="shintvshoppingSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="grpVal" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= FormatNumber(oShintvshopping.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oShintvshopping.FTotalPage,0) %></b>
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
	<td width="140">shintvshopping등록일<br>shintvshopping최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">shintvshopping<br>가격및판매</td>
	<td width="70">shintvshopping<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oShintvshopping.FResultCount - 1 %>
<tr align="center" <% If oShintvshopping.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oShintvshopping.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oShintvshopping.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oShintvshopping.FItemList(i).FItemID %>','shintvshopping','<%=oShintvshopping.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oShintvshopping.FItemList(i).FItemID%>" target="_blank"><%= oShintvshopping.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oShintvshopping.FItemList(i).FshintvshoppingStatcd <> 7 Then
	%>
		<br><%= oShintvshopping.FItemList(i).getshintvshoppingStatName %>
	<%
			End If
			response.write oShintvshopping.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oShintvshopping.FItemList(i).FMakerid %> <%= oShintvshopping.FItemList(i).getDeliverytypeName %><br><%= oShintvshopping.FItemList(i).FItemName %></td>
	<td align="center"><%= oShintvshopping.FItemList(i).FRegdate %><br><%= oShintvshopping.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oShintvshopping.FItemList(i).FshintvshoppingRegdate %><br><%= oShintvshopping.FItemList(i).FshintvshoppingLastUpdate %></td>
	<td align="right">
		<% If oShintvshopping.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oShintvshopping.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oShintvshopping.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oShintvshopping.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oShintvshopping.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oShintvshopping.FItemList(i).Fbuycash/oShintvshopping.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oShintvshopping.FItemList(i).Fbuycash/oShintvshopping.FItemList(i).FshintvshoppingPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oShintvshopping.FItemList(i).Fbuycash/oShintvshopping.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oShintvshopping.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oShintvshopping.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oShintvshopping.FItemList(i).FOrgSuplycash/oShintvshopping.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oShintvshopping.FItemList(i).Fbuycash/oShintvshopping.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oShintvshopping.FItemList(i).Fbuycash/oShintvshopping.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oShintvshopping.FItemList(i).IsSoldOut Then
			If oShintvshopping.FItemList(i).FSellyn = "N" Then
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
		If oShintvshopping.FItemList(i).FItemdiv = "06" OR oShintvshopping.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oShintvshopping.FItemList(i).FshintvshoppingStatCd > 0) Then
			If Not IsNULL(oShintvshopping.FItemList(i).FshintvshoppingPrice) Then
				If (oShintvshopping.FItemList(i).Mustprice <> oShintvshopping.FItemList(i).FshintvshoppingPrice) Then
	%>
					<strong><%= formatNumber(oShintvshopping.FItemList(i).FshintvshoppingPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oShintvshopping.FItemList(i).FshintvshoppingPrice,0)&"<br>"
				End If

				If Not IsNULL(oShintvshopping.FItemList(i).FSpecialPrice) Then
					If (now() >= oShintvshopping.FItemList(i).FStartDate) And (now() <= oShintvshopping.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oShintvshopping.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oShintvshopping.FItemList(i).FSellyn="Y" and oShintvshopping.FItemList(i).FshintvshoppingSellYn<>"Y") or (oShintvshopping.FItemList(i).FSellyn<>"Y" and oShintvshopping.FItemList(i).FshintvshoppingSellYn="Y") Then
	%>
					<strong><%= oShintvshopping.FItemList(i).FshintvshoppingSellYn %></strong>
	<%
				Else
					response.write oShintvshopping.FItemList(i).FshintvshoppingSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oShintvshopping.FItemList(i).FshintvshoppingGoodNo <> "" Then %>
			<a target="_blank" href="https://www.shinsegaetvshopping.com/display/detail/<%=oShintvshopping.FItemList(i).FshintvshoppingGoodNo%>"><%=oShintvshopping.FItemList(i).FshintvshoppingGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oShintvshopping.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oShintvshopping.FItemList(i).FItemID%>','0');"><%= oShintvshopping.FItemList(i).FoptionCnt %>:<%= oShintvshopping.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oShintvshopping.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oShintvshopping.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If

		If oShintvshopping.FItemList(i).FReglevel < 5 Then
			response.write " <br />레벨 : " & oShintvshopping.FItemList(i).FReglevel
		End If
	%>
	</td>
	<td align="center">
		<%= oShintvshopping.FItemList(i).FinfoDiv %>
		<%
		If (oShintvshopping.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oShintvshopping.FItemList(i).FlastErrStr) &"'>ERR:"& oShintvshopping.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% shintvshoppingGoodNoArray = shintvshoppingGoodNoArray & oShintvshopping.FItemList(i).FshintvshoppingGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= shintvshoppingGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oShintvshopping.HasPreScroll then %>
		<a href="javascript:goPage('<%= oShintvshopping.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oShintvshopping.StartScrollPage to oShintvshopping.FScrollCount + oShintvshopping.StartScrollPage - 1 %>
    		<% if i>oShintvshopping.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oShintvshopping.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
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
	<input type="hidden" name="shintvshoppingGoodNo" value= <%= shintvshoppingGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="shintvshoppingYes10x10No" value= <%= shintvshoppingYes10x10No %>>
	<input type="hidden" name="shintvshoppingNo10x10Yes" value= <%= shintvshoppingNo10x10Yes %>>
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
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
	<input type="hidden" name="reglevel" value= <%= reglevel %>>
</form>
<% Set oShintvshopping = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->