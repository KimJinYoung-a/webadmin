<%@ language=vbscript %>
<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, EzwelGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchPrddiv, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, EzwelYes10x10No, EzwelNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isextusing, cisextusing, rctsellcnt
Dim page, i, research, getRegdate
Dim oEzwel
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
EzwelGoodNo				= request("EzwelGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
EzwelYes10x10No			= request("EzwelYes10x10No")
EzwelNo10x10Yes			= request("EzwelNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
getRegdate				= request("getRegdate")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
purchasetype			= request("purchasetype")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
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
'Ezwel 상품코드 엔터키로 검색되게
If EzwelGoodNo<>"" then
	Dim iA2, arrTemp2, arrEzwelGoodNo
	EzwelGoodNo = replace(EzwelGoodNo,",",chr(10))
	EzwelGoodNo = replace(EzwelGoodNo,chr(13),"")
	arrTemp2 = Split(EzwelGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrEzwelGoodNo = arrEzwelGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	EzwelGoodNo = left(arrEzwelGoodNo,len(arrEzwelGoodNo)-1)
End If

Set oEzwel = new CEzwel
	oEzwel.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oEzwel.FPageSize					= 100
Else
	oEzwel.FPageSize					= 50
End If
	oEzwel.FRectCDL					= request("cdl")
	oEzwel.FRectCDM					= request("cdm")
	oEzwel.FRectCDS					= request("cds")
	oEzwel.FRectItemID				= itemid
	oEzwel.FRectItemName			= itemname
	oEzwel.FRectSellYn				= sellyn
	oEzwel.FRectLimitYn				= limityn
	oEzwel.FRectSailYn				= sailyn
'	oEzwel.FRectonlyValidMargin		= onlyValidMargin
	oEzwel.FRectStartMargin			= startMargin
	oEzwel.FRectEndMargin			= endMargin
	oEzwel.FRectMakerid				= makerid
	oEzwel.FRectEzwelGoodNo			= EzwelGoodNo
	oEzwel.FRectMatchCate			= MatchCate
	oEzwel.FRectIsMadeHand			= isMadeHand
	oEzwel.FRectIsOption			= isOption
	oEzwel.FRectIsReged				= isReged
	oEzwel.FRectNotinmakerid		= notinmakerid
	oEzwel.FRectNotinitemid			= notinitemid
	oEzwel.FRectExcTrans			= exctrans
	oEzwel.FRectPriceOption			= priceOption
	oEzwel.FRectIsSpecialPrice     	= isSpecialPrice
	oEzwel.FRectDeliverytype		= deliverytype
	oEzwel.FRectMwdiv				= mwdiv
	oEzwel.FRectGetRegdate			= getRegdate
	oEzwel.FRectIsextusing			= isextusing
	oEzwel.FRectCisextusing			= cisextusing
	oEzwel.FRectRctsellcnt			= rctsellcnt

	oEzwel.FRectExtNotReg			= ExtNotReg
	oEzwel.FRectExpensive10x10		= expensive10x10
	oEzwel.FRectdiffPrc				= diffPrc
	oEzwel.FRectEzwelYes10x10No		= EzwelYes10x10No
	oEzwel.FRectEzwelNo10x10Yes		= EzwelNo10x10Yes
	oEzwel.FRectExtSellYn			= extsellyn
	oEzwel.FRectInfoDiv				= infoDiv
	oEzwel.FRectFailCntOverExcept	= ""
	oEzwel.FRectFailCntExists		= failCntExists
	oEzwel.FRectReqEdit				= reqEdit
	oEzwel.FRectPurchasetype		= purchasetype
If (bestOrd = "on") Then
    oEzwel.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oEzwel.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oEzwel.getEzwelreqExpireItemList
Else
	oEzwel.getEzwelRegedItemList			'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
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
function ezwelNotInMakerid(){
	var popwin=window.open('/admin/etc/ezwel/targetMall_Not_In_Makerid.asp?mallgubun=ezwel','notin','width=900,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 상품
function ezwelNotInItemid(){
	var popwin2=window.open('/admin/etc/ezwel/targetMall_Not_In_Itemid.asp?mallgubun=ezwel','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=ezwel','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//카테고리 관리
function pop_cateManager() {
	var pCM2 = window.open("/admin/etc/ezwel/popezwelcateList.asp","popCateezwelmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//New카테고리 관리
function pop_newcateManager() {
	var pCMNew = window.open("/admin/etc/ezwel/popNewezwelcateList.asp","popCateezwelmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMNew.focus();
}
//2022-11-03 김진영..전시카테고리로 매칭
function pop_dispcateManager() {
	var pCMDisp = window.open("/admin/etc/ezwel/popDispezwelcateList.asp","popdispcateManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMDisp.focus();
}
//승인목록 관리
function pop_statcdManager() {
	var pCMMng = window.open("/admin/etc/ezwel/popstatcdList.asp","popstatcdManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMMng.focus();
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
	if ((comp.name!="EzwelYes10x10No")&&(frm.EzwelYes10x10No.checked)){ frm.EzwelYes10x10No.checked=false }
	if ((comp.name!="EzwelNo10x10Yes")&&(frm.EzwelNo10x10Yes.checked)){ frm.EzwelNo10x10Yes.checked=false }
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

    if ((comp.name=="EzwelYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="EzwelNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.EzwelYes10x10No.checked){
            comp.form.EzwelYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="D"
	        comp.form.sellyn.value = "Y";
	        comp.form.onlyValidMargin.value="Y";
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
	if ((comp.name!="EzwelYes10x10No")&&(frm.EzwelYes10x10No.checked)){ frm.EzwelYes10x10No.checked=false }
	if ((comp.name!="EzwelNo10x10Yes")&&(frm.EzwelNo10x10Yes.checked)){ frm.EzwelNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}

//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
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

// 선택된 상품 일괄 등록
function EzwelSelectRegProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※Ezwel과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        //document.frmSvArr.action = "/admin/etc/ezwel/actezwelReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 판매여부 변경
function EzwelSellYnProcess(chkYn) {
	var chkSel=0, strSell;
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
		case "Y": strSell="판매중";break;
		case "N": strSell="품절";break;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Ezwel과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		if (chkYn == 'Y'){
			if (confirm('[중요]이지웰페어 담당자와 판매상태 협의 하였습니까?')){
		        document.frmSvArr.target = "xLink";
		        document.frmSvArr.cmdparam.value = "EditSellYn";
		        document.frmSvArr.chgSellYn.value = chkYn;
		        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
		        document.frmSvArr.submit();
			}
		}else{
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
	        document.frmSvArr.submit();
	    }
    }
}
// 선택된 상품정보 수정
function EzwelSelectEditProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 판매여부 변경
function EzwelSellYnNewProcess(chkYn) {
	var chkSel=0, strSell;
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
		case "Y": strSell="판매중";break;
		case "N": strSell="품절";break;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Ezwel과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		if (chkYn == 'Y'){
			if (confirm('[중요]이지웰페어 담당자와 판매상태 협의 하였습니까?')){
		        document.frmSvArr.target = "xLink";
		        document.frmSvArr.cmdparam.value = "EditSellYn";
		        document.frmSvArr.chgSellYn.value = chkYn;
		        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
		        document.frmSvArr.submit();
			}
		}else{
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
	        document.frmSvArr.submit();
	    }
    }
}

// 선택된 상품 일괄 등록
function EzwelSelectNewRegProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※Ezwel과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}

function EzwelSelectNewEditProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 조회
function EzwelSelectNewViewProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}


// 선택된 상품 상태 조회
function EzwelSelectViewProcess() {
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

    if (confirm('Ezwel에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 승인
function EzwelSelectStatCdProcess() {
	var conStr;
	if (document.getElementById("getRegdate").value != ""){
		conStr = "Ezwel에 "+document.getElementById("getRegdate").value+"일에 등록한 상품을 승인 받았습니까?";
	}else{
		conStr = "Ezwel에 모든 등록된 상품을 MD승인 받았습니까?";
	}

    if (confirm(conStr)){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "STAT";
        document.frmSvArr.getRegdate.value = document.getElementById("getRegdate").value;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=ezwel&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
<% if request("auto") = "Y" then %>
function EzwelSelectEditProcessAuto() {
	var cnt = <%= oEzwel.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oEzwel.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		EzwelSelectEditProcessAuto();
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

</script>
<!-- 검색 시작 -->
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
		<a href="http://shopadmin.ezwel.com" target="_blank">이지웰페어 Admin바로가기</a>
		<%
			If C_ADMIN_AUTH OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ CSP관리자 | 10x10 | Cube1010* ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		Ezwel 상품코드 : <textarea rows="2" cols="20" name="EzwelGoodNo" id="itemid"><%=replace(EzwelGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Ezwel 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >Ezwel 등록예정이상
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >Ezwel 전송시도중오류
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >Ezwel 승인예정
			<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >Ezwel 재판매예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Ezwel 등록완료(전시)
		</select>&nbsp;
		<label><input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">전체</label>&nbsp;
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Ezwel 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="EzwelYes10x10No" <%= ChkIIF(EzwelYes10x10No="on","checked","") %> ><font color=red>Ezwel판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="EzwelNo10x10Yes" <%= ChkIIF(EzwelNo10x10Yes="on","checked","") %> ><font color=red>Ezwel품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
		&nbsp;
		등록일 : <input id="getRegdate" name="getRegdate" value="<%= getRegdate %>" class="text" size="10" maxlength="10" />
		<img src="http://scm.10x10.co.kr/images/calicon.gif" id="gDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		(Ezwel승인여부 검색 및 실제상품 수정의 승인에서 사용하는 날짜 입니다)
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, <del>판매가(할인가) 1만원 미만</del>, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 옵션수(텐텐:제휴) 다른상품, 주문제작문구상품

<p />
<% end if %>
<!-- 액션 시작 -->
<form name="frmReg" method="post" action="ezwelItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="등록 제외 브랜드" onclick="ezwelNotInMakerid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="ezwelNotInItemid();">&nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('ezwel');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<!--
				<input class="button" type="button" value="카테고리" onclick="pop_cateManager();">
				-->
				<input class="button" type="button" value="카테고리" onclick="pop_newcateManager();">
				<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="카테고리" onclick="pop_dispcateManager();">
				<% End If %>
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
	    		<!-- <input class="button" type="button" id="btnRegSel" value="(구)상품 등록" onClick="EzwelSelectRegProcess();">&nbsp;&nbsp; -->
				<input class="button" type="button" id="btnViewSel" value="등록" onClick="EzwelSelectNewRegProcess();">&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<!--
			    <input class="button" type="button" id="btnEditSel" value="(구)수정" onClick="EzwelSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnViewSel" value="(구)조회" onClick="EzwelSelectViewProcess();">&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnViewSel" value="수정" onClick="EzwelSelectNewEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnViewSel" value="조회" onClick="EzwelSelectNewViewProcess();">&nbsp;&nbsp;
				<br><br>
				실제상품 승인 :
				<input class="button" type="button" value="승인목록" onClick="pop_statcdManager();">
			   <!-- <input class="button" type="button" id="btnStat" value="승인" onClick="EzwelSelectStatCdProcess();"> -->
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<!-- <option value="Y">판매중</option> -->
				</Select>(으)로
				<!-- <input class="button" type="button" id="btnSellYn" value="(구)변경" onClick="EzwelSellYnProcess(frmReg.chgSellYn.value);">&nbsp;&nbsp; -->
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="EzwelSellYnNewProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="getRegdate">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oEzwel.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEzwel.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">Ezwel등록일<br>Ezwel최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Ezwel<br>가격및판매</td>
	<td width="70">Ezwel<br>실가격</td>
	<td width="70">Ezwel<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oEzwel.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oEzwel.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oEzwel.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oEzwel.FItemList(i).FItemID %>','ezwel','')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oEzwel.FItemList(i).FItemID%>" target="_blank"><%= oEzwel.FItemList(i).FItemID %></a>
		<% If oEzwel.FItemList(i).FLimitYn= "Y" Then %><br><%= oEzwel.FItemList(i).getLimitHtmlStr %></font><% End If %>
		<%
			If oEzwel.FItemList(i).FEzwelStatcd = "3" Then
				response.write "<br />승인예정"
			ElseIf oEzwel.FItemList(i).FEzwelStatcd = "4" Then
				response.write "<br />재판매예정"
			End If
		%>
	</td>
	<td align="left"><%= oEzwel.FItemList(i).FMakerid %> <%= oEzwel.FItemList(i).getDeliverytypeName %><br><%= oEzwel.FItemList(i).FItemName %></td>
	<td align="center"><%= oEzwel.FItemList(i).FRegdate %><br><%= oEzwel.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oEzwel.FItemList(i).FEzwelRegdate %><br><%= oEzwel.FItemList(i).FEzwelLastUpdate %></td>
	<td align="right">
		<% If oEzwel.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oEzwel.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).Fsellcash = 0 Then
		elseif (oEzwel.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oEzwel.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oEzwel.FItemList(i).FOrgSuplycash/oEzwel.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).IsSoldOut Then
			If oEzwel.FItemList(i).FSellyn = "N" Then
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
		If oEzwel.FItemList(i).FItemdiv = "06" OR oEzwel.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oEzwel.FItemList(i).FEzwelStatCd > 0) Then
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				If (oEzwel.FItemList(i).Fsellcash<>oEzwel.FItemList(i).FEzwelPrice) Then
	%>
					<strong><%= formatNumber(oEzwel.FItemList(i).FEzwelPrice,0) %></strong>
	<%
				Else
					response.write formatNumber(oEzwel.FItemList(i).FEzwelPrice,0)
				End If
	%>
				<br>
	<%
				If Not IsNULL(oEzwel.FItemList(i).FSpecialPrice) Then
					If (now() >= oEzwel.FItemList(i).FStartDate) And (now() <= oEzwel.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oEzwel.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oEzwel.FItemList(i).FSellyn="Y" and oEzwel.FItemList(i).FEzwelSellYn<>"Y") or (oEzwel.FItemList(i).FSellyn<>"Y" and oEzwel.FItemList(i).FEzwelSellYn="Y") Then
	%>
					<strong><%= oEzwel.FItemList(i).FEzwelSellYn %></strong>
	<%
				Else
					response.write oEzwel.FItemList(i).FEzwelSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<%
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				response.write FormatNumber(Fix(oEzwel.FItemList(i).FEzwelPrice/100)*100,0)
		 	End If
		%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oEzwel.FItemList(i).FEzwelGoodNo)) Then
		    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://shop.ezwel.com/shopNew/goods/preview/goodsDetailView.ez?preview=yes&goodsBean.goodsCd="&oEzwel.FItemList(i).FEzwelGoodNo&"')>"&oEzwel.FItemList(i).FEzwelGoodNo&"</span><br>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oEzwel.FItemList(i).FEzwelStatCd="0","(등록예정)","")
		End If
	%>
	</td>
	<td align="center"><%= oEzwel.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oEzwel.FItemList(i).FItemID%>','0');"><%= oEzwel.FItemList(i).FoptionCnt %>:<%= oEzwel.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oEzwel.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oEzwel.FItemList(i).FCateMapCnt > 0 Then %>
		매칭됨
	<% Else %>
		<font color="darkred">매칭안됨</font>
	<% End If %>

	<% If (oEzwel.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oEzwel.FItemList(i).FlastErrStr %>">ERR:<%= oEzwel.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center"><%= oEzwel.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oEzwel.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEzwel.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEzwel.StartScrollPage to oEzwel.FScrollCount + oEzwel.StartScrollPage - 1 %>
    		<% if i>oEzwel.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEzwel.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "getRegdate", trigger    : "gDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<% SET oEzwel = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
