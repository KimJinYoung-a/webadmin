<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/nvstorefarmClass/nvClassCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, nvClassGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, nvClassYes10x10No, nvClassNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing
Dim page, i, research
Dim oNvclass
Dim startMargin, endMargin
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
nvClassGoodNo			= request("nvClassGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
nvClassYes10x10No		= request("nvClassYes10x10No")
nvClassNo10x10Yes		= request("nvClassNo10x10Yes")
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

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = ""
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
If nvClassGoodNo <> "" then
	Dim iA2, arrTemp2, arrnvClassGoodNo
	nvClassGoodNo = replace(nvClassGoodNo,",",chr(10))
	nvClassGoodNo = replace(nvClassGoodNo,chr(13),"")
	arrTemp2 = Split(nvClassGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrnvClassGoodNo = arrnvClassGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	nvClassGoodNo = left(arrnvClassGoodNo,len(arrnvClassGoodNo)-1)
End If

Set oNvclass = new CNvClass
	oNvclass.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oNvclass.FPageSize					= 100
Else
	oNvclass.FPageSize					= 50
End If
	oNvclass.FRectCDL					= request("cdl")
	oNvclass.FRectCDM					= request("cdm")
	oNvclass.FRectCDS					= request("cds")
	oNvclass.FRectItemID				= itemid
	oNvclass.FRectItemName				= itemname
	oNvclass.FRectSellYn				= sellyn
	oNvclass.FRectLimitYn				= limityn
	oNvclass.FRectSailYn				= sailyn
'	oNvclass.FRectonlyValidMargin		= onlyValidMargin
	oNvclass.FRectStartMargin			= startMargin
	oNvclass.FRectEndMargin				= endMargin
	oNvclass.FRectMakerid				= makerid
	oNvclass.FRectNvClassGoodNo			= nvClassGoodNo
	oNvclass.FRectMatchCate				= MatchCate
	oNvclass.FRectIsMadeHand			= isMadeHand
	oNvclass.FRectIsOption				= isOption
	oNvclass.FRectIsReged				= isReged
	oNvclass.FRectNotinmakerid			= notinmakerid
	oNvclass.FRectNotinitemid			= notinitemid
	oNvclass.FRectExcTrans				= exctrans
	oNvclass.FRectPriceOption			= priceOption
	oNvclass.FRectIsSpecialPrice     	= isSpecialPrice
	oNvclass.FRectDeliverytype			= deliverytype
	oNvclass.FRectMwdiv					= mwdiv
	oNvclass.FRectIsextusing			= isextusing

	oNvclass.FRectExtNotReg				= ExtNotReg
	oNvclass.FRectExpensive10x10		= expensive10x10
	oNvclass.FRectdiffPrc				= diffPrc
	oNvclass.FRectNvClassYes10x10No		= nvClassYes10x10No
	oNvclass.FRectNvClassNo10x10Yes		= nvClassNo10x10Yes
	oNvclass.FRectExtSellYn				= extsellyn
	oNvclass.FRectInfoDiv				= infoDiv
	oNvclass.FRectFailCntOverExcept		= ""
	oNvclass.FRectFailCntExists			= failCntExists
	oNvclass.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oNvclass.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oNvclass.FRectOrdType = "BM"
End If

If isReged = "R" Then								'품절처리요망 상품보기 리스트
	oNvclass.getNvClassreqExpireItemList
Else
	oNvclass.getNvClassRegedItemList		'그 외 리스트
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

// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=nvstorefarmclass','popNotInItemid','width=500,height=500,scrollbars=yes,resizable=yes');
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

    if ((comp.name=="nvClassYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="nvClassNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.nvClassYes10x10No.checked){
            comp.form.nvClassYes10x10No.checked = false;
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
	if ((comp.name!="nvClassYes10x10No")&&(frm.nvClassYes10x10No.checked)){ frm.nvClassYes10x10No.checked=false }
	if ((comp.name!="nvClassNo10x10Yes")&&(frm.nvClassNo10x10Yes.checked)){ frm.nvClassNo10x10Yes.checked=false }
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
	if ((comp.name!="nvClassYes10x10No")&&(frm.nvClassYes10x10No.checked)){ frm.nvClassYes10x10No.checked=false }
	if ((comp.name!="nvClassNo10x10Yes")&&(frm.nvClassNo10x10Yes.checked)){ frm.nvClassNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
// 선택된 상품 등록
function nvClassSelectRegItemProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmClass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function nvClassSelectRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function nvClassSelectEDITProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 조회
function nvClassSelectItemSearchProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 조회
function nvClassSelectOptionSearchProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 등록
function nvClassSelectImageRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 옵션 등록
function nvClassSelectOPTRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

function nvClassSelectDelProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

function nvClassSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=nvstorefarmclass&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
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
		스토어팜 상품코드 : <textarea rows="2" cols="20" name="nvClassGoodNo" id="itemid"><%=replace(nvClassGoodNo,",",chr(10))%></textarea>
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvClassYes10x10No" <%= ChkIIF(nvClassYes10x10No="on","checked","") %> ><font color=red>스토어팜판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvClassNo10x10Yes" <%= ChkIIF(nvClassNo10x10Yes="on","checked","") %> ><font color=red>스토어팜품절&텐바이텐판매가능</font>(판매중,한정>5) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
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
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();"> &nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('nvstorefarmclass');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegImgSel" value="이미지" onClick="nvClassSelectImageRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="yhj0613") OR (session("ssBctID")="hrkang97") Then %>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnRegSel" value="상품" onClick="nvClassSelectRegItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegOptSel" value="옵션" onClick="nvClassSelectOPTRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelSel" value="삭제" onClick="nvClassSelectDelProcess();">&nbsp;&nbsp;
				<% Else %>
				<input class="button" type="button" id="btnDelSel" value="삭제" onClick="nvClassSelectDelProcess();">&nbsp;&nbsp;
				<% End If %>
			<% End If %>
				<input class="button" type="button" id="btnReg" value="상품+옵션" onClick="nvClassSelectRegProcess();" style=color:red;font-weight:bold>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnReg" value="수정" onClick="nvClassSelectEDITProcess();" style=color:blue;font-weight:bold>
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				실제상품 조회 :
				<input class="button" type="button" id="btnSchitem" value="상품" onClick="nvClassSelectItemSearchProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnSchOpt" value="옵션" onClick="nvClassSelectOptionSearchProcess();">&nbsp;&nbsp;
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">판매중지</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="nvClassSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="17">
		검색결과 : <b><%= FormatNumber(oNvclass.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oNvclass.FTotalPage,0) %></b>
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
	<td width="70">스토어팜<br>가격및판매</td>
	<td width="70">스토어팜<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">품목</td>
	<td width="100">이미지<br>업로드</td>
</tr>
<% For i=0 to oNvclass.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oNvclass.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oNvclass.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oNvclass.FItemList(i).FItemID %>','nvstorefarmclass','<%=oNvclass.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oNvclass.FItemList(i).FItemID%>" target="_blank"><%= oNvclass.FItemList(i).FItemID %></a>
		<% If oNvclass.FItemList(i).FNvClassStatcd <> 7 Then %>
		<br><%= oNvclass.FItemList(i).getNvClassStatName %>
		<% End If %>
		<%= oNvclass.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oNvclass.FItemList(i).FMakerid %> <%= oNvclass.FItemList(i).getDeliverytypeName %><br><%= oNvclass.FItemList(i).FItemName %></td>
	<td align="center"><%= oNvclass.FItemList(i).FRegdate %><br><%= oNvclass.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oNvclass.FItemList(i).FNvClassRegdate %><br><%= oNvclass.FItemList(i).FNvClassLastUpdate %></td>
	<td align="right">
		<% If oNvclass.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oNvclass.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oNvclass.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oNvclass.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oNvclass.FItemList(i).Fsellcash = 0 Then
			'//
		elseIf (oNvclass.FItemList(i).FNvClassStatCd > 0) and Not IsNULL(oNvclass.FItemList(i).FNvClassPrice) Then
			If (oNvclass.FItemList(i).FSaleYn = "Y") and (oNvclass.FItemList(i).FSellcash < oNvclass.FItemList(i).FNvClassPrice) Then
				'// 제휴몰 정상가 판매중
		%>
		<strike><%= CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		<font color="#CC3333"><%= CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).FNvClassPrice*100*100)/100 & "%" %></font>
		<%
			else
				response.write CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%"
			end if
		else
			response.write CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oNvclass.FItemList(i).IsSoldOut Then
			If oNvclass.FItemList(i).FSellyn = "N" Then
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
		If (oNvclass.FItemList(i).FNvClassStatCd > 0) Then
			If Not IsNULL(oNvclass.FItemList(i).FNvClassPrice) Then
				If (oNvclass.FItemList(i).Fsellcash <> oNvclass.FItemList(i).FNvClassPrice) Then
	%>
					<strong><%= formatNumber(oNvclass.FItemList(i).FNvClassPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oNvclass.FItemList(i).FNvClassPrice,0)&"<br>"
				End If

				If (oNvclass.FItemList(i).FSellyn="Y" and oNvclass.FItemList(i).FNvClassSellYn<>"Y") or (oNvclass.FItemList(i).FSellyn<>"Y" and oNvclass.FItemList(i).FNvClassSellYn="Y") Then
	%>
					<strong><%= oNvclass.FItemList(i).FNvClassSellYn %></strong>
	<%
				Else
					response.write oNvclass.FItemList(i).FNvClassSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oNvclass.FItemList(i).FNvClassGoodNo)) Then
			Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytenclass/products/"&oNvclass.FItemList(i).FNvClassGoodNo&"'>"&oNvclass.FItemList(i).FNvClassGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oNvclass.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oNvclass.FItemList(i).FItemID%>','0');"><%= oNvclass.FItemList(i).FoptionCnt %>:<%= oNvclass.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oNvclass.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= oNvclass.FItemList(i).FinfoDiv %>
		<%
		If (oNvclass.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oNvclass.FItemList(i).FlastErrStr) &"'>ERR:"& oNvclass.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oNvclass.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oNvclass.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>미등록</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oNvclass.HasPreScroll then %>
		<a href="javascript:goPage('<%= oNvclass.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oNvclass.StartScrollPage to oNvclass.FScrollCount + oNvclass.StartScrollPage - 1 %>
    		<% if i>oNvclass.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oNvclass.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oNvclass = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
