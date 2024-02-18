<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/nvstoregift/nvstoregiftcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, nvstoregiftGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, nvstoregiftYes10x10No, nvstoregiftNo10x10Yes, reqEdit, reqExpire, failCntExists, scheduleNotInItemid, isextusing, cisextusing, rctsellcnt
Dim page, i, research
Dim onvstoregift
Dim startMargin, endMargin
Dim purchasetype
page    					= request("page")
research					= request("research")
itemid  					= request("itemid")
makerid						= request("makerid")
itemname					= request("itemname")
bestOrd						= request("bestOrd")
bestOrdMall					= request("bestOrdMall")
sellyn						= request("sellyn")
limityn						= request("limityn")
sailyn						= request("sailyn")
onlyValidMargin				= request("onlyValidMargin")
startMargin					= request("startMargin")
endMargin					= request("endMargin")
isMadeHand					= request("isMadeHand")
isOption					= request("isOption")
infoDiv						= request("infoDiv")
morningJY					= request("morningJY")
extsellyn					= request("extsellyn")
nvstoregiftGoodNo			= request("nvstoregiftGoodNo")
ExtNotReg					= request("ExtNotReg")
isReged						= request("isReged")
MatchCate					= request("MatchCate")
expensive10x10				= request("expensive10x10")
diffPrc						= request("diffPrc")
nvstoregiftYes10x10No		= request("nvstoregiftYes10x10No")
nvstoregiftNo10x10Yes		= request("nvstoregiftNo10x10Yes")
reqEdit						= request("reqEdit")
reqExpire					= request("reqExpire")
failCntExists				= request("failCntExists")
optAddPrcRegTypeNone		= request("optAddPrcRegTypeNone")
notinmakerid				= request("notinmakerid")
priceOption					= request("priceOption")
isSpecialPrice				= request("isSpecialPrice")
deliverytype				= request("deliverytype")
mwdiv						= request("mwdiv")
notinitemid					= requestCheckVar(request("notinitemid"), 1)
exctrans					= requestCheckVar(request("exctrans"), 1)
scheduleNotInItemid			= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing					= requestCheckVar(request("isextusing"), 1)
cisextusing					= requestCheckVar(request("cisextusing"), 1)
rctsellcnt					= requestCheckVar(request("rctsellcnt"), 1)
purchasetype				= request("purchasetype")

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
'스토어팜 상품코드 엔터키로 검색되게
If nvstoregiftGoodNo <> "" then
	Dim iA2, arrTemp2, arrnvstoregiftGoodNo
	nvstoregiftGoodNo = replace(nvstoregiftGoodNo,",",chr(10))
	nvstoregiftGoodNo = replace(nvstoregiftGoodNo,chr(13),"")
	arrTemp2 = Split(nvstoregiftGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrnvstoregiftGoodNo = arrnvstoregiftGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	nvstoregiftGoodNo = left(arrnvstoregiftGoodNo,len(arrnvstoregiftGoodNo)-1)
End If

Set oNvstoregift = new CNvstoregift
	oNvstoregift.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oNvstoregift.FPageSize					= 100
Else
	oNvstoregift.FPageSize					= 50
End If
	oNvstoregift.FRectCDL					= request("cdl")
	oNvstoregift.FRectCDM					= request("cdm")
	oNvstoregift.FRectCDS					= request("cds")
	oNvstoregift.FRectItemID				= itemid
	oNvstoregift.FRectItemName				= itemname
	oNvstoregift.FRectSellYn				= sellyn
	oNvstoregift.FRectLimitYn				= limityn
	oNvstoregift.FRectSailYn				= sailyn
	oNvstoregift.FRectStartMargin			= startMargin
	oNvstoregift.FRectEndMargin				= endMargin
	oNvstoregift.FRectMakerid				= makerid
	oNvstoregift.FRectnvstoregiftGoodNo		= nvstoregiftGoodNo
	oNvstoregift.FRectMatchCate				= MatchCate
	oNvstoregift.FRectIsMadeHand			= isMadeHand
	oNvstoregift.FRectIsOption				= isOption
	oNvstoregift.FRectIsReged				= isReged
	oNvstoregift.FRectNotinmakerid			= notinmakerid
	oNvstoregift.FRectNotinitemid			= notinitemid
	oNvstoregift.FRectExcTrans				= exctrans
	oNvstoregift.FRectPriceOption			= priceOption
	oNvstoregift.FRectIsSpecialPrice  	   	= isSpecialPrice
	oNvstoregift.FRectDeliverytype			= deliverytype
	oNvstoregift.FRectMwdiv					= mwdiv
	oNvstoregift.FRectScheduleNotInItemid	= scheduleNotInItemid
	oNvstoregift.FRectIsextusing			= isextusing
	oNvstoregift.FRectCisextusing			= cisextusing
	oNvstoregift.FRectRctsellcnt			= rctsellcnt

	oNvstoregift.FRectExtNotReg				= ExtNotReg
	oNvstoregift.FRectExpensive10x10		= expensive10x10
	oNvstoregift.FRectdiffPrc				= diffPrc
	oNvstoregift.FRectnvstoregiftYes10x10No = nvstoregiftYes10x10No
	oNvstoregift.FRectnvstoregiftNo10x10Yes = nvstoregiftNo10x10Yes
	oNvstoregift.FRectExtSellYn				= extsellyn
	oNvstoregift.FRectInfoDiv				= infoDiv
	oNvstoregift.FRectFailCntOverExcept		= ""
	oNvstoregift.FRectFailCntExists			= failCntExists
	oNvstoregift.FRectReqEdit				= reqEdit
	oNvstoregift.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oNvstoregift.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oNvstoregift.FRectOrdType = "BM"
End If

If isReged = "R" Then								'품절처리요망 상품보기 리스트
	oNvstoregift.getnvstoregiftreqExpireItemList
Else
	oNvstoregift.getnvstoregiftRegedItemList		'그 외 리스트
End If
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=nvstoregift","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=nvstoregift','popNotInItemid','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=nvstoregift','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 브랜드(EP)
function NotInMakeridEP(){
    var popwin = window.open("/admin/etc/potal/notinmakerid.asp?mallid=naverEP","popNotInMakeridEP","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품(EP)
function NotInItemidEP(){
	var popwin2=window.open('/admin/etc/potal/notinitemid.asp?mallid=naverEP','popNotInItemidEP','width=1200,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=nvstoregift','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
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

    if ((comp.name=="nvstoregiftYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="nvstoregiftNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.nvstoregiftYes10x10No.checked){
            comp.form.nvstoregiftYes10x10No.checked = false;
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
	if ((comp.name!="nvstoregiftYes10x10No")&&(frm.nvstoregiftYes10x10No.checked)){ frm.nvstoregiftYes10x10No.checked=false }
	if ((comp.name!="nvstoregiftNo10x10Yes")&&(frm.nvstoregiftNo10x10Yes.checked)){ frm.nvstoregiftNo10x10Yes.checked=false }
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
	if ((comp.name!="nvstoregiftYes10x10No")&&(frm.nvstoregiftYes10x10No.checked)){ frm.nvstoregiftYes10x10No.checked=false }
	if ((comp.name!="nvstoregiftNo10x10Yes")&&(frm.nvstoregiftNo10x10Yes.checked)){ frm.nvstoregiftNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
// 선택된 상품 등록
function NvstoregiftSelectRegItemProcess() {
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGITEM";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function NvstoregiftSelectRegProcess() {
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function NvstoregiftSelectEDITProcess() {
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 조회
function NvstoregiftSelectItemSearchProcess(){
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 조회
function NvstoregiftSelectOptionSearchProcess(){
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품 옵션을 조회 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 등록
function NvstoregiftSelectImageRegProcess() {
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품 이미지를 일괄 등록 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "Image";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 옵션 등록
function NvstoregiftSelectOPTRegProcess() {
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품 옵션을 일괄 등록 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}

function NvstoregiftSelectDelProcess(){
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

    if (confirm('스토어팜에 선택하신 ' + chkSel + '개 상품 옵션을 일괄 삭제 하시겠습니까?\n\n※스토어팜과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "DEL";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}

function NvstoregiftSellYnProcess(chkYn) {
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
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp"
        document.frmSvArr.submit();
    }
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// 스토어팜 카테고리 관리
function pop_CateManager() {
	var pCM = window.open("/admin/etc/nvstorefarm/popnvstorefarmCateList.asp","popnvstorefarm","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//공통코드 검색
function NvstorefarmCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('공통코드를 선택하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		xLink.location.href = "<%=apiURL%>/outmall/nvstoregift/actnvstoregiftReq.asp?cmdparam=nvstorefarmCommonCode&CommCD="+ccd+"";
	}
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=nvstoregift&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('http://scm.10x10.co.kr/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
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
		<a href="http://sell.storefarm.naver.com" target="_blank">스토어팜 Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenten | tenbytenstore ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		스토어팜 상품코드 : <textarea rows="2" cols="20" name="nvstoregiftGoodNo" id="itemid"><%= replace(replace(nvstoregiftGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >스토어팜 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >스토어팜 등록예정이상
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >스토어팜 전송시도중오류
			<option value="I" <%= CHkIIF(ExtNotReg="I","selected","") %> >스토어팜 이미지만 완료
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >스토어팜 등록완료(전시)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">미등록<font color="<%= CHKIIF(makerid="" and itemid="", "#000000", "#AAAAAA") %>">(최근 3개월 등록상품만)</font></label>&nbsp;
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>스토어팜 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvstoregiftYes10x10No" <%= ChkIIF(nvstoregiftYes10x10No="on","checked","") %> ><font color=red>스토어팜판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvstoregiftNo10x10Yes" <%= ChkIIF(nvstoregiftNo10x10Yes="on","checked","") %> ><font color=red>스토어팜품절&텐바이텐판매가능</font>(판매중,한정>5) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, <del>판매가(할인가) 1만원 미만</del>, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 주문제작상품, 주문제작문구상품, 판매가(할인가) 1천원 미만, 일부 품목(화장품, 식품(농수산물), 가공식품, 건강기능식품) 상품, 옵션가 0원 판매중 상품이 없음(옵션 한정수량 5개 이하 포함)

<p />

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
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">&nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('nvstoregift');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegImgSel" value="이미지" onClick="NvstoregiftSelectImageRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="yhj0613") OR (session("ssBctID")="hrkang97") Then %>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnRegSel" value="상품" onClick="NvstoregiftSelectRegItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegOptSel" value="옵션" onClick="NvstoregiftSelectOPTRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelSel" value="삭제" onClick="NvstoregiftSelectDelProcess();">&nbsp;&nbsp;
				<% Else %>
				<input class="button" type="button" id="btnDelSel" value="삭제" onClick="NvstoregiftSelectDelProcess();">&nbsp;&nbsp;
				<% End If %>
			<% End If %>
				<input class="button" type="button" id="btnReg" value="상품+옵션" onClick="NvstoregiftSelectRegProcess();" style=color:red;font-weight:bold>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnReg" value="수정" onClick="NvstoregiftSelectEDITProcess();" style=color:blue;font-weight:bold>
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				실제상품 조회 :
				<input class="button" type="button" id="btnSchitem" value="상품" onClick="NvstoregiftSelectItemSearchProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnSchOpt" value="옵션" onClick="NvstoregiftSelectOptionSearchProcess();">&nbsp;&nbsp;
				<br><br>
				공통코드 검색 :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="GetAddressBookList">판매자주소
				</select>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="NvstoregiftCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">판매중지</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="NvstoregiftSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oNvstoregift.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oNvstoregift.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">스토어팜등록일<br>스토어팜최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">스토어팜<br>가격및판매</td>
	<td width="70">스토어팜<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
	<td width="100">이미지<br>업로드</td>
</tr>
<% For i=0 to oNvstoregift.FResultCount - 1 %>
<tr align="center" <% If oNvstoregift.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oNvstoregift.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oNvstoregift.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oNvstoregift.FItemList(i).FItemID %>','nvstoregift','<%=oNvstoregift.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oNvstoregift.FItemList(i).FItemID%>" target="_blank"><%= oNvstoregift.FItemList(i).FItemID %></a>
		<% If oNvstoregift.FItemList(i).FNvstoregiftStatcd <> 7 Then %>
		<br><%= oNvstoregift.FItemList(i).getNvstoregiftStatName %>
		<% End If %>
		<%= oNvstoregift.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oNvstoregift.FItemList(i).FMakerid %> <%= oNvstoregift.FItemList(i).getDeliverytypeName %><br><%= oNvstoregift.FItemList(i).FItemName %></td>
	<td align="center"><%= oNvstoregift.FItemList(i).FRegdate %><br><%= oNvstoregift.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oNvstoregift.FItemList(i).FNvstoregiftRegdate %><br><%= oNvstoregift.FItemList(i).FNvstoregiftLastUpdate %></td>
	<td align="right">
		<% If oNvstoregift.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oNvstoregift.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oNvstoregift.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oNvstoregift.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oNvstoregift.FItemList(i).Fsellcash = 0 Then
		elseif (oNvstoregift.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oNvstoregift.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oNvstoregift.FItemList(i).FOrgSuplycash/oNvstoregift.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oNvstoregift.FItemList(i).Fbuycash/oNvstoregift.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oNvstoregift.FItemList(i).Fbuycash/oNvstoregift.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oNvstoregift.FItemList(i).IsSoldOut Then
			If oNvstoregift.FItemList(i).FSellyn = "N" Then
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
		If oNvstoregift.FItemList(i).FItemdiv = "06" OR oNvstoregift.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oNvstoregift.FItemList(i).FNvstoregiftStatCd > 0) Then
			If Not IsNULL(oNvstoregift.FItemList(i).FNvstoregiftPrice) Then
				If (oNvstoregift.FItemList(i).Fsellcash <> oNvstoregift.FItemList(i).FNvstoregiftPrice) Then
	%>
					<strong><%= formatNumber(oNvstoregift.FItemList(i).FNvstoregiftPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oNvstoregift.FItemList(i).FNvstoregiftPrice,0)&"<br>"
				End If

				If Not IsNULL(oNvstoregift.FItemList(i).FSpecialPrice) Then
					If (now() >= oNvstoregift.FItemList(i).FStartDate) And (now() <= oNvstoregift.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oNvstoregift.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oNvstoregift.FItemList(i).FSellyn="Y" and oNvstoregift.FItemList(i).FNvstoregiftSellYn<>"Y") or (oNvstoregift.FItemList(i).FSellyn<>"Y" and oNvstoregift.FItemList(i).FNvstoregiftSellYn="Y") Then
	%>
					<strong><%= oNvstoregift.FItemList(i).FNvstoregiftSellYn %></strong>
	<%
				Else
					response.write oNvstoregift.FItemList(i).FNvstoregiftSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oNvstoregift.FItemList(i).FNvstoregiftGoodNo)) Then
			Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytengift/products/"&oNvstoregift.FItemList(i).FNvstoregiftGoodNo&"'>"&oNvstoregift.FItemList(i).FNvstoregiftGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oNvstoregift.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oNvstoregift.FItemList(i).FItemID%>','0');"><%= oNvstoregift.FItemList(i).FoptionCnt %>:<%= oNvstoregift.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oNvstoregift.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oNvstoregift.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oNvstoregift.FItemList(i).FinfoDiv %>
		<%
		If (oNvstoregift.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oNvstoregift.FItemList(i).FlastErrStr) &"'>ERR:"& oNvstoregift.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oNvstoregift.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oNvstoregift.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>미등록</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oNvstoregift.HasPreScroll then %>
		<a href="javascript:goPage('<%= oNvstoregift.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oNvstoregift.StartScrollPage to oNvstoregift.FScrollCount + oNvstoregift.StartScrollPage - 1 %>
    		<% if i>oNvstoregift.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oNvstoregift.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oNvstoregift = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
