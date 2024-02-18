<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, gsshopgoodno, extsellyn, ExtNotReg, isReged, MatchCate, MatchPrddiv, notinmakerid, priceOption, deliverytype, mwdiv
Dim expensive10x10, diffPrc, gsshopYes10x10No, gsshopNo10x10Yes, reqEdit, reqExpire, failCntExists, diffName
Dim page, i, research
Dim ogsshop
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
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")
infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
gsshopgoodno			= request("gsshopgoodno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
gsshopYes10x10No		= request("gsshopYes10x10No")
gsshopNo10x10Yes		= request("gsshopNo10x10Yes")
reqEdit					= request("reqEdit")
diffName				= request("diffName")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")

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
'GSShop 상품코드 엔터키로 검색되게
If gsshopgoodno<>"" then
	Dim iA2, arrTemp2, arrgsshopgoodno
	gsshopgoodno = replace(gsshopgoodno,",",chr(10))
	gsshopgoodno = replace(gsshopgoodno,chr(13),"")
	arrTemp2 = Split(gsshopgoodno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrgsshopgoodno = arrgsshopgoodno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	gsshopgoodno = left(arrgsshopgoodno,len(arrgsshopgoodno)-1)
End If

SET oGSShop = new CGSShop
	oGSShop.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oGSShop.FPageSize					= 50
Else
	oGSShop.FPageSize					= 20
End If
	oGSShop.FRectCDL					= request("cdl")
	oGSShop.FRectCDM					= request("cdm")
	oGSShop.FRectCDS					= request("cds")
	oGSShop.FRectItemID					= itemid
	oGSShop.FRectItemName				= itemname
	oGSShop.FRectSellYn					= sellyn
	oGSShop.FRectLimitYn				= limityn
	oGSShop.FRectSailYn					= sailyn
	oGSShop.FRectonlyValidMargin		= onlyValidMargin
	oGSShop.FRectMakerid				= makerid
	oGSShop.FRectGSShopgoodno			= gsshopgoodno
	oGSShop.FRectMatchCate				= MatchCate
	oGSShop.FRectPrdDivMatch			= MatchPrddiv
	oGSShop.FRectIsMadeHand				= isMadeHand
	oGSShop.FRectIsOption				= isOption
	oGSShop.FRectIsReged				= isReged
	oGSShop.FRectNotinmakerid			= notinmakerid
	oGSShop.FRectPriceOption			= priceOption
	oGSShop.FRectDeliverytype			= deliverytype
	oGSShop.FRectMwdiv					= mwdiv

	oGSShop.FRectExtNotReg				= ExtNotReg
	oGSShop.FRectExpensive10x10			= expensive10x10
	oGSShop.FRectdiffPrc				= diffPrc
	oGSShop.FRectGSShopYes10x10No		= gsshopYes10x10No
	oGSShop.FRectGSShopNo10x10Yes		= gsshopNo10x10Yes
	oGSShop.FRectExtSellYn				= extsellyn
	oGSShop.FRectInfoDiv				= infoDiv
	oGSShop.FRectFailCntOverExcept		= ""
	oGSShop.FRectFailCntExists			= failCntExists
	oGSShop.FRectReqEdit				= reqEdit
	oGSShop.FRectDiffName				= diffName
If (bestOrd = "on") Then
    oGSShop.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oGSShop.FRectOrdType = "BM"
End If
	oGSShop.getGSShopAddOptionRegedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
//안전인증 필수 팝업
function pop_safecode(itemcd){
	var popwin=window.open('/admin/etc/gsshop/pop_safecode.asp?itemid='+itemcd+'','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
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
	if ((comp.name!="gsshopYes10x10No")&&(frm.gsshopYes10x10No.checked)){ frm.gsshopYes10x10No.checked=false }
	if ((comp.name!="gsshopNo10x10Yes")&&(frm.gsshopNo10x10Yes.checked)){ frm.gsshopNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
	if ((comp.name!="diffName")&&(frm.diffName.checked)){ frm.diffName.checked=false }
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "J"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "J"
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

    if ((comp.name=="gsshopYes10x10No")&&(comp.checked)){
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
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="gsshopNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.gsshopYes10x10No.checked){
            comp.form.gsshopYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="G"
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
			comp.form.ExtNotReg.value="G"
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
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}
	if (comp.name=="diffName"&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
             comp.form.expensive10x10.checked = false;
        }
		if (comp.checked){
        	document.getElementById("AR").checked=true;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.value="G"
			comp.form.ExtNotReg.disabled = false;
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gsshopYes10x10No")&&(frm.gsshopYes10x10No.checked)){ frm.gsshopYes10x10No.checked=false }
	if ((comp.name!="gsshopNo10x10Yes")&&(frm.gsshopNo10x10Yes.checked)){ frm.gsshopNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
	if ((comp.name!="diffName")&&(frm.diffName.checked)){ frm.diffName.checked=false }
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
function GSShopSelectRegProcess(isreal) {
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

    if (isreal){
        if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "REG";
            document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?\n\n※30분단위로 배치 등록됩니다.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/gsshop/actgsshopReq.asp"
            document.frmSvArr.submit();
        }
    }
}

// 선택된 상품 가격 수정
function GSShopPriceEditProcess() {
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품가격을 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditPrice";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

//수기로 regedoption 등록
function Sugi_regedoption() {
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
    document.frmSvArr.target = "xLink";
    document.frmSvArr.ckLimit.value = "<%= limityn %>";
    document.frmSvArr.cmdparam.value = "sugiRegedoption";
    document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
    document.frmSvArr.submit();
}

// 선택된 이미지(대표 및 썸네일) 수정
function GSShopImageEditProcess() {
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 이미지를 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditImage").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditImage";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품설명 수정
function GSShopContentsEditProcess() {
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품설명을 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditContents").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 재고 및 옵션 추가/수정
function GSShopOPTEditProcess() {
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 이미지를 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditOPT").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditOPT";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품명 수정
function GSShopItemnameEditProcess() {
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품명을 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditItemname";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

//정부고시항목 수정
function GSShopInfodivEditProcess(){
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 상품명을 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditInfoDiv").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditInfoDiv";
        document.frmSvArr.cmdparam.value = "INFODIV";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}


//기본정보 수정
function GSShopIteminfoEditProcess(){
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

    if (confirm('GSShop에 선택하신 ' + chkSel + '개 기본정보를 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITINFO";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}


// 선택된 상품 판매여부 변경
function GSShopSellYnProcess(chkYn) {
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

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※GSShop과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 GSShop에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=gsshop&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM6 = window.open("/admin/etc/que/popQueOptionLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function confirmOK(itemcd){
	if (confirm('텐바이텐 상품코드 : ' + itemcd + '\n승인 확인 하셨습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditStatCd";
        document.frmSvArr.chgStatItemCode.value = itemcd;
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
	}
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function chgItemname(idx, iname){
	document.frmUp.target = "xLink";
    document.frmUp.idx.value = idx
    document.frmUp.cName.value = document.getElementById(iname).value;
    document.frmUp.mode.value = "chgName"
    document.frmUp.action = "/admin/etc/optManager/optProc.asp"
    document.frmUp.submit();

}
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
		브랜드&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;<% 'OutmallAdminInfo("gsshop") %>
		&nbsp;
		<a href="https://withgs.gsshop.com/cmm/login" target="_blank">GSShop Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  1003890 | store101010* | 1010tenbyten* ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		GSShop 상품코드 : <textarea rows="2" cols="20" name="gsshopgoodno" id="itemid"><%=replace(gsshopgoodno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >GSShop 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >GSShop 등록예정이상
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >GSShop 등록예정
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >GSShop 전송시도중오류
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >GSShop 등록후 승인대기(임시)
			<option value="G" <%= CHkIIF(ExtNotReg="G","selected","") %> >GSShop 등록완료 승인대기이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >GSShop 등록완료(전시)
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
		텐바이텐 :
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>&nbsp;
		판매
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >전체
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
		</select>&nbsp;
		한정
		<select name="limityn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >한정
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >일반
		</select>&nbsp;
		세일
		<select name="sailyn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >세일Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >세일N
		</select>&nbsp;
		기준마진(<%= Chkiif(mallid="gsshop", "14.9", "") %>%)
		<select name="onlyValidMargin" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(onlyValidMargin="Y","selected","") %> >마진이상
			<option value="N" <%= CHkIIF(onlyValidMargin="N","selected","") %> >마진이하
		</select>&nbsp;
		주문제작
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>&nbsp;
		품목
		<select name="infodiv" class="select">
			<option value="" <%= CHkIIF(infoDiv="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >입력
			<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >미입력
		<%
			For infoLoop = 1 To 35
				If infoLoop < 10 Then
					infoDivValue = "0"&infoLoop
				Else
					infoDivValue = infoLoop
				End If
		%>
			<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
		<% Next %>
		</select>
		<br>
		제휴몰 &nbsp;&nbsp; :
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>&nbsp;
		판매
		<select name="extsellyn" class="select">
			<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
			<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >종료
			<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >종료제외
		</select>&nbsp;
		오류
		<select name="failCntExists" class="select">
			<option value="" <%= CHkIIF(failCntExists="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >등록수정오류1회이상
			<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >등록수정오류0회
		</select>&nbsp;
		카테고리
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
		상품분류
		<select name="MatchPrddiv" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>GSShop 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gsshopYes10x10No" <%= ChkIIF(gsshopYes10x10No="on","checked","") %> ><font color=red>GSShop판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="gsshopNo10x10Yes" <%= ChkIIF(gsshopNo10x10Yes="on","checked","") %> ><font color=red>GSShop품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffName" <%= ChkIIF(diffName="on","checked","") %> ><font color=red>상품/옵션명 이상</font>상품보기</label>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmReg" method="post" action="gsshopItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('gsshop');">&nbsp;&nbsp;
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
	    		<input class="button" type="button" id="btnRegSel" value="상품 등록" onClick="GSShopSelectRegProcess(true);">
				<br><br>
				실제상품 수정 :
			    <input class="button" type="button" id="btnEditSel" value="가격" onClick="GSShopPriceEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditImage" value="이미지(대표 및 썸네일)" onClick="GSShopImageEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditContents" value="상품설명" onClick="GSShopContentsEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditOPT" value="가격&재고&옵션&상태수정" onClick="GSShopOPTEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditName" value="상품명" onClick="GSShopItemnameEditProcess();">
   			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfoDiv" value="정부고시항목" onClick="GSShopInfodivEditProcess();">
				<% If (session("ssBctID")="kjy8517") Then %>
				 &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfo" value="기본정보" onClick="GSShopIteminfoEditProcess();">
				<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">일시중단</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="GSShopSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<form name="frmUp" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="idx" value="">
<input type="hidden" name="cName" value="">
<input type="hidden" name="mode" value="">
</form>

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
		검색결과 : <b><%= FormatNumber(oGSShop.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oGSShop.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">상품코드<br>옵션코드</td>
	<td>상품정보</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">GSShop등록일<br>GSShop최종수정일</td>
	<td width="70">판매가<br><font color="purple">옵션가</font></td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">GSShop<br>가격및판매</td>
	<td width="70">GSShop<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">3개월<br>판매량</td>
<!--	<td width="50">브랜드<br>매핑</td> -->
	<td width="60">카테고리<br>매칭여부</td>
	<td width="100"><font color="BLUE">GS상품분류</font><br><font color="Green">GS 안전인증</font><br><font color="BROWN">GS 과세여부</font></td>
	<td width="80">품목</td>
</tr>
<% For i = 0 To oGSShop.FResultCount - 1 %>
<% If (oGSShop.FItemList(i).FItemName <> oGSShop.FItemList(i).FRegedItemname) OR (oGSShop.FItemList(i).FOptionname <> oGSShop.FItemList(i).FRegedOptionname) Then %>
<tr align="center" bgcolor="GOLD">
<% Else %>
<tr align="center" bgcolor="#FFFFFF">
<% End If %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oGSShop.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oGSShop.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oGSShop.FItemList(i).FItemID %>','GSShop','<%=oGSShop.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oGSShop.FItemList(i).FItemID%>" target="_blank"><%= oGSShop.FItemList(i).FItemID %></a>
		<br><font color="gray"><%=oGSShop.FItemList(i).FItemOption%></font><br>
		<%
			If oGSShop.FItemList(i).getGSShopOptItemStatCd = "승인대기" Then
				response.write "<input type='button' class=button value="&oGSShop.FItemList(i).getGSShopOptItemStatCd&" onclick=confirmOK('"&oGSShop.FItemList(i).FItemID&"')><br>"
			Else
				response.write oGSShop.FItemList(i).getGSShopOptItemStatCd
			End If
		%>
	</td>
	<td align="left">
		<%= oGSShop.FItemList(i).FMakerid %><%= oGSShop.FItemList(i).getDeliverytypeName %>
		<br>(실) : <%= oGSShop.FItemList(i).FItemName %>
		<br>(등) : <%= oGSShop.FItemList(i).FRegedItemname %>
		<br>(실) : <%= oGSShop.FItemList(i).FOptionname %>
		<br>(등) : <%= oGSShop.FItemList(i).FRegedOptionname %>
		<br><input type="text" style="color:red" id="newitemname<%=oGSShop.FItemList(i).FIdx%>" size="50" value="<%= oGSShop.FItemList(i).getRealItemname %>">
		<input type="button" class="button" value="수정" onclick="chgItemname('<%= oGSShop.FItemList(i).FIdx %>', 'newitemname<%=oGSShop.FItemList(i).FIdx%>')">
	</td>
	<td align="center"><%= oGSShop.FItemList(i).FRegdate %><br><%= oGSShop.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oGSShop.FItemList(i).FGSShopRegdate %><br><%= oGSShop.FItemList(i).FGSShopLastUpdate %></td>
	<td align="right">
	<% If oGSShop.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oGSShop.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oGSShop.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oGSShop.FItemList(i).FSellcash,0) %>
	<% End If %>
	<br><font color="purple">+<%= oGSShop.FItemList(i).FOptaddprice %></font>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oGSShop.FItemList(i).Fbuycash/oGSShop.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).IsSoldOut Then
			If oGSShop.FItemList(i).FSellyn = "N" Then
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
		If oGSShop.FItemList(i).FItemdiv = "06" OR oGSShop.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oGSShop.FItemList(i).FGSShopStatCd > 0) Then
			If Not IsNULL(oGSShop.FItemList(i).FGSShopPrice) Then
				If (oGSShop.FItemList(i).Fsellcash + oGSShop.FItemList(i).FOptaddprice <> oGSShop.FItemList(i).FGSShopPrice) Then
	%>
					<strong><%= formatNumber(oGSShop.FItemList(i).FGSShopPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oGSShop.FItemList(i).FGSShopPrice,0)&"<br>"
				End If

				If (oGSShop.FItemList(i).FSellyn="Y" and oGSShop.FItemList(i).FGSShopSellYn<>"Y") or (oGSShop.FItemList(i).FSellyn<>"Y" and oGSShop.FItemList(i).FGSShopSellYn="Y") or (oGSShop.FItemList(i).IsSoldOutLimit5Sell and oGSShop.FItemList(i).FGSShopSellYn="Y") Then
	%>
					<strong><%= oGSShop.FItemList(i).FGSShopSellYn %></strong>
	<%
				Else
					response.write oGSShop.FItemList(i).FGSShopSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		'#실상품번호
		If Not(IsNULL(oGSShop.FItemList(i).FGSShopGoodNo)) Then
	    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.gsshop.com/prd/prd.gs?prdid="&oGSShop.FItemList(i).FGSShopGoodNo&"')>"&oGSShop.FItemList(i).FGSShopGoodNo&"</span>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oGSShop.FItemList(i).FGSShopStatCd="0","(등록예정)","")
		End If
	%>
	</td>
	<td align="center"><%= oGSShop.FItemList(i).Freguserid %></td>
	<td align="center"><%= oGSShop.FItemList(i).FrctSellCNT %></td>
<!--
	<td align="center">
	<%
		If (oGSShop.FItemList(i).FBrandCd = "") OR (oGSShop.FItemList(i).FDeliveryAddrCd = "") OR (oGSShop.FItemList(i).FDeliveryCd = "") Then
	%>
		<font color="darkred">매칭안됨</font>
	<%
		Else
			response.write "매칭됨"
		End If
	%>
	</td>
-->
	<td align="center">
	<% If oGSShop.FItemList(i).FCateMapCnt > 0 Then %>
	    매칭됨
	<% Else %>
		<font color="darkred">매칭안됨</font>
	<% End If %>

	<% If (oGSShop.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oGSShop.FItemList(i).FlastErrStr %>">ERR:<%= oGSShop.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).FDivcode = "" Then
			response.write "매칭안됨"
		Else
			rw "<font color='BLUE'>매칭됨</font>"
			Select Case oGSShop.FItemList(i).FSafeCode
				Case "1"	response.write "<input type='button' value='필수' onclick=pop_safecode('"&oGSShop.FItemList(i).FItemid&"'); class='button'>"
					If oGSShop.FItemList(i).FSafeCertGbnCd <> "" Then
						rw "<font color='BLUE'>( Y )</font>"
					Else
						rw "<font color='RED'>( N )</font>"
					End If
				Case "2"	response.write "<input type='button' value='선택' onclick=pop_safecode('"&oGSShop.FItemList(i).FItemid&"'); class='button'>"
					If oGSShop.FItemList(i).FSafeCertGbnCd <> "" Then
						rw "<font color='BLUE'>( Y )</font>"
					Else
						rw "<font color='RED'>( N )</font>"
					End If
				Case "3" 	rw "<font color='Green'>비대상</font>"
			End Select
			Select Case oGSShop.FItemList(i).FIsvat
				Case "01"	rw "<font color='Brown'>면세</font>"
				Case "02"	rw "<font color='Brown'>과세</font>"
				Case "03" 	rw "<font color='Brown'>영세</font>"
				Case "99" 	rw "<font color='Brown'>전체허용</font>"
			End Select
		End If
	%>
	</td>
	<td align="center"><%= oGSShop.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oGSShop.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGSShop.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oGSShop.StartScrollPage to oGSShop.FScrollCount + oGSShop.StartScrollPage - 1 %>
    		<% if i>oGSShop.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oGSShop.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oGSShop = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
