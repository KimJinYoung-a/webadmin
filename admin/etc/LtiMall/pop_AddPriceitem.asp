<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim mallid, infoLoop, infoDivValue
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, ltimallgoodno, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, priceOption, ltimalltmpgoodno, deliverytype, mwdiv
Dim expensive10x10, diffPrc, ltimallYes10x10No, ltimallNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oiMall

mallid					= CMALLNAME
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
ltimallgoodno			= request("ltimallgoodno")
ltimalltmpgoodno		= request("ltimalltmpgoodno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
ltimallYes10x10No		= request("ltimallYes10x10No")
ltimallNo10x10Yes		= request("ltimallNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
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
'롯데iMall 상품코드 엔터키로 검색되게
If ltimallgoodno <> "" then
	Dim iA2, arrTemp2, arrltimallgoodno
	ltimallgoodno = replace(ltimallgoodno,",",chr(10))
	ltimallgoodno = replace(ltimallgoodno,chr(13),"")
	arrTemp2 = Split(ltimallgoodno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrltimallgoodno = arrltimallgoodno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	ltimallgoodno = left(arrltimallgoodno,len(arrltimallgoodno)-1)
End If

'롯데iMall 승인전 상품코드 엔터키로 검색되게
If ltimalltmpgoodno <> "" then
	Dim iA3, arrTemp3, arrltimalltmpgoodno
	ltimalltmpgoodno = replace(ltimalltmpgoodno,",",chr(10))
	ltimalltmpgoodno = replace(ltimalltmpgoodno,chr(13),"")
	arrTemp3 = Split(ltimalltmpgoodno,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			If Not(isNumeric(trim(arrTemp3(iA3)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp3(iA3) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrltimalltmpgoodno = arrltimalltmpgoodno & trim(arrTemp3(iA3)) & ","
			End If
		End If
		iA3 = iA3 + 1
	Loop
	ltimalltmpgoodno = left(arrltimalltmpgoodno,len(arrltimalltmpgoodno)-1)
End If

Set oiMall = new CLotteiMall
If (session("ssBctID")="kjy8517") Then
	oiMall.FPageSize					= 50
Else
	oiMall.FPageSize					= 20
End If
	oiMall.FCurrPage					= page
	oiMall.FRectMakerid					= makerid
	oiMall.FRectItemID					= itemid
	oiMall.FRectItemName				= itemname
	oiMall.FRectLTiMallGoodNo			= ltimallgoodno
	oiMall.FRectLTimallTmpGoodNo		= ltimalltmpgoodno
	oiMall.FRectCDL						= request("cdl")
	oiMall.FRectCDM						= request("cdm")
	oiMall.FRectCDS						= request("cds")
	oiMall.FRectExtNotReg				= ExtNotReg
	oiMall.FRectIsReged					= isReged
	oiMall.FRectNotinmakerid			= notinmakerid
	oiMall.FRectPriceOption				= priceOption
	oiMall.FRectDeliverytype			= deliverytype
	oiMall.FRectMwdiv					= mwdiv

	oiMall.FRectSellYn					= sellyn
	oiMall.FRectLimitYn					= limityn
	oiMall.FRectSailYn					= sailyn
	oiMall.FRectonlyValidMargin			= onlyValidMargin
	oiMall.FRectIsMadeHand				= isMadeHand
	oiMall.FRectIsOption				= isOption
	oiMall.FRectInfoDiv					= infoDiv
	oiMall.FRectExtSellYn				= extsellyn
	oiMall.FRectFailCntExists			= failCntExists
	oiMall.FRectMatchCate				= MatchCate
	oiMall.FRectExpensive10x10			= expensive10x10
	oiMall.FRectdiffPrc					= diffPrc
	oiMall.FRectLtimallYes10x10No		= ltimallYes10x10No
	oiMall.FRectLtimallNo10x10Yes		= ltimallNo10x10Yes
	oiMall.FRectReqEdit					= reqEdit
If (bestOrd = "on") Then
    oiMall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oiMall.FRectOrdType = "BM"
End If
	oiMall.getLTiMallAddOptionRegedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteimall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteimall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteimall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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
	if ((comp.name!="ltimallYes10x10No")&&(frm.ltimallYes10x10No.checked)){ frm.ltimallYes10x10No.checked=false }
	if ((comp.name!="ltimallNo10x10Yes")&&(frm.ltimallNo10x10Yes.checked)){ frm.ltimallNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
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

    if ((comp.name=="ltimallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="ltimallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.ltimallYes10x10No.checked){
            comp.form.ltimallYes10x10No.checked = false;
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
	if ((comp.name!="ltimallYes10x10No")&&(frm.ltimallYes10x10No.checked)){ frm.ltimallYes10x10No.checked=false }
	if ((comp.name!="ltimallNo10x10Yes")&&(frm.ltimallNo10x10Yes.checked)){ frm.ltimallNo10x10Yes.checked=false }
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
	var pCM6 = window.open("/admin/etc/que/popQueOptionLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 선택된 상품 판매여부 변경
function LTiMallSellYnProcess(chkYn) {
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
		case "X": strSell="판매종료(삭제)";break;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 롯데iMall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 가격 수정
function LtimallPriceEditProcess() {
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

    if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품명 수정
function LtimallItemnameEditProcess() {
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

    if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품명을 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditItemNm";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function LtimallSelectRegProcess(isreal) {
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
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "REG";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?\n\n※30분단위로 배치 등록됩니다.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }
}

//선택 신규상품조회
function LtimallStatCheckProcess(){
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

    if (confirm('롯데아이몰에 선택하신 ' + chkSel + '개 상품을 판매 상태를 확인 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "CheckItemStat";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

//선택상품 전시상품조회
function LtimallDispViewProcess(){
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

    if (confirm('롯데아이몰에 선택하신 ' + chkSel + '개 상품을 전시상품 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
         document.frmSvArr.cmdparam.value = "DISPVIEW";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function LtimallEditProcess(v) {
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

    if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		if(v == ""){
			document.getElementById("btnEditSel").disabled=true;
		}else{
			document.getElementById("btnEditSel2").disabled=true;
		}
		document.frmSvArr.target = "xLink";
		//document.frmSvArr.cmdparam.value = "EditSelect" + v;
		document.frmSvArr.cmdparam.value = "EDIT" + v;
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 등록
function LtimallregIMSI(isreg) {
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

    if (isreg){
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?')){
			document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 삭제 하시겠습니까?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }
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
		<a href="https://partner.lotteimall.com/main/Login.lotte" target="_blank">롯데아이몰Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 011799LT | store101010* | 01063242821 ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		롯데iMall 상품코드 : <textarea rows="2" cols="20" name="ltimallgoodno" id="itemid"><%=replace(ltimallgoodno,",",chr(10))%></textarea>
		&nbsp;
		승인전 상품코드 : <textarea rows="2" cols="20" name="ltimalltmpgoodno" id="itemid"><%=replace(ltimalltmpgoodno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >롯데iMall 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >롯데iMall 등록예정이상
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >롯데iMall 등록예정
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >롯데iMall 전송시도중오류
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >롯데iMall 반려
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >롯데iMall 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >롯데iMall 등록완료(전시)
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
		기준마진(<%= Chkiif(mallid="lotteimall", "14.9", "") %>%)
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>롯데iMall 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="ltimallYes10x10No" <%= ChkIIF(ltimallYes10x10No="on","checked","") %> ><font color=red>롯데iMall판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="ltimallNo10x10Yes" <%= ChkIIF(ltimallNo10x10Yes="on","checked","") %> ><font color=red>롯데iMall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmUp" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="idx" value="">
<input type="hidden" name="cName" value="">
<input type="hidden" name="mode" value="">
</form>

<form name="frmReg" method="post" action="lotteimallItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lotteimall');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="LtimallSelectRegProcess(true);">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="LtimallPriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="정보&가격&옵션&재고&상태" onClick="LtimallEditProcess('');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditName" value="상품명" onClick="LtimallItemnameEditProcess();">
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="전시상품조회" onClick="LtimallDispViewProcess();">
				<% End If %>
				<br><br>
				승인예정 상품 :
				<!--
				<input class="button" type="button" id="btnEditSel2" value="수정" onClick="LtimallEditProcess('2');">
				&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnEditInfoDiv" value="신규상품조회" onClick="LtimallStatCheckProcess();">
				<br><br>
				등록예정 상품 :
				<input class="button" type="button" id="btnRegImsi" value="등록" onClick="LtimallregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="삭제" onClick="LtimallregIMSI(false);" >
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
					<option value="X">판매종료(삭제)</option><!-- 삭제하면 이후 수정 할 수 없음 -->
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="LTiMallSellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(oiMall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oiMall.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">상품코드<br>옵션코드</td>
	<td>상품정보</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">롯데iMall등록일<br>롯데iMall최종수정일</td>
	<td width="70">판매가<br><font color="purple">옵션가</font></td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">롯데iMall<br>가격및판매</td>
	<td width="70">롯데iMall<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i = 0 To oiMall.FResultCount - 1 %>
<% If (oiMall.FItemList(i).FItemName <> oiMall.FItemList(i).FRegedItemname) OR (oiMall.FItemList(i).FOptionname <> oiMall.FItemList(i).FRegedOptionname) Then %>
<tr align="center" bgcolor="GOLD">
<% Else %>
<tr align="center" bgcolor="#FFFFFF">
<% End If %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oiMall.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oiMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oiMall.FItemList(i).FItemID %>','lotteimall','<%=oiMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oiMall.FItemList(i).FItemID%>" target="_blank"><%= oiMall.FItemList(i).FItemID %></a>
		<br><font color="gray"><%=oiMall.FItemList(i).FItemOption%></font><br>
		<% if oiMall.FItemList(i).FLimitYn="Y" then %><br><%= oiMall.FItemList(i).getLimitHtmlStr %></font><% end if %>
	</td>
	<td align="left">
		<%= oiMall.FItemList(i).FMakerid %><%= oiMall.FItemList(i).getDeliverytypeName %>
		<br>(실) : <%= oiMall.FItemList(i).FItemName %>
		<br>(등) : <%= oiMall.FItemList(i).FRegedItemname %>
		<br>(실) : <%= oiMall.FItemList(i).FOptionname %>
		<br>(등) : <%= oiMall.FItemList(i).FRegedOptionname %>
		<br><input type="text" style="color:red" id="newitemname<%=oiMall.FItemList(i).FIdx%>" size="50" value="<%= oiMall.FItemList(i).getRealItemname %>">
		<input type="button" class="button" value="수정" onclick="chgItemname('<%= oiMall.FItemList(i).FIdx %>', 'newitemname<%=oiMall.FItemList(i).FIdx%>')">
	</td>
	<td align="center"><%= oiMall.FItemList(i).FRegdate %><br><%= oiMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oiMall.FItemList(i).FLtimallRegdate %><br><%= oiMall.FItemList(i).FLtimallLastUpdate %></td>

	<td align="right">
	<% If oiMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oiMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %>
	<% End If %>
		<br><font color="purple">+<%= oiMall.FItemList(i).FOptaddprice %></font>
	</td>
	<td align="center">
	<%
		If oiMall.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oiMall.FItemList(i).IsSoldOut Then
			If oiMall.FItemList(i).FSellyn = "N" Then
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
		If oiMall.FItemList(i).FItemdiv = "06" OR oiMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oiMall.FItemList(i).FLtimallStatCd > 0) Then
			If Not IsNULL(oiMall.FItemList(i).FLtimallPrice) Then
				If (oiMall.FItemList(i).Fsellcash + oiMall.FItemList(i).FOptaddprice <> oiMall.FItemList(i).FLtimallPrice) Then
	%>
					<strong><%= formatNumber(oiMall.FItemList(i).FLtimallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oiMall.FItemList(i).FLtimallPrice,0)&"<br>"
				End If

				If (oiMall.FItemList(i).FSellyn="Y" and oiMall.FItemList(i).FLtimallSellYn<>"Y") or (oiMall.FItemList(i).FSellyn<>"Y" and oiMall.FItemList(i).FLtimallSellYn="Y") Then
	%>
					<strong><%= oiMall.FItemList(i).FLtimallSellYn %></strong>
	<%
				Else
					response.write oiMall.FItemList(i).FLtimallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
    <%
    	'#실상품번호
    	if Not(IsNULL(oiMall.FItemList(i).FLtiMallGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotteimall.com/product/Product.jsp?i_code="&oiMall.FItemList(i).FLtiMallGoodNo&"'>"&oiMall.FItemList(i).FLtiMallGoodNo&"</a>"
		else
			'#임시상품번호
			if Not(IsNULL(oiMall.FItemList(i).FLtiMallTmpGoodNo)) then
				if oiMall.FItemList(i).FLTiMallStatCd<>"30" then
					Response.Write "<br>(" & oiMall.FItemList(i).FLtiMallTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if

		if (oiMall.FItemList(i).FLTiMallStatCd<>7) then
		    response.write "<br>"&oiMall.FItemList(i).getLTIMallStatCDName
		end if
	%>
	</td>
	<td align="center"><%= oiMall.FItemList(i).Freguserid %></td>
	<td align="center"><%= oiMall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oiMall.FItemList(i).FCateMapCnt > 0 Then %>
	    매칭됨
	<% Else %>
		<font color="darkred">매칭안됨</font>
	<% End If %>

	<% If (oiMall.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oiMall.FItemList(i).FlastErrStr %>">ERR:<%= oiMall.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
    <td align="center"><%= oiMall.FItemList(i).FinfoDiv %>
    <% if (oiMall.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oiMall.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oiMall.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">옵션금액</font>
	    <% If oiMall.FItemList(i).FoptAddPrcRegType<>0 Then %>
	    (<%=oiMall.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% End If %>
    </td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oiMall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
    		<% if i>oiMall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oiMall.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oiMall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->