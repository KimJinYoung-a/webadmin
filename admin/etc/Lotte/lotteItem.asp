<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/lotte/lotteitemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, lotteGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, lotteTmpGoodNo, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, lotteYes10x10No, lotteNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing
Dim page, i, research
Dim oLotteitem
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
lotteGoodNo				= request("lotteGoodNo")
lotteTmpGoodNo			= request("lotteTmpGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
lotteYes10x10No			= request("lotteYes10x10No")
lotteNo10x10Yes			= request("lotteNo10x10Yes")
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
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid="316889,934033,939944,940421,942435,948037,948220,948343,950570,952113,953810,958080,960587,964133,964871,969288,974333,976963,978955,979010,979436,980488,981438,983944,988755,988767,993408,993413"

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
'롯데닷컴 상품코드 엔터키로 검색되게
If lotteGoodNo <> "" then
	Dim iA2, arrTemp2, arrlotteGoodNo
	lotteGoodNo = replace(lotteGoodNo,",",chr(10))
	lotteGoodNo = replace(lotteGoodNo,chr(13),"")
	arrTemp2 = Split(lotteGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrlotteGoodNo = arrlotteGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	lotteGoodNo = left(arrlotteGoodNo,len(arrlotteGoodNo)-1)
End If

'롯데닷컴 승인전 상품코드 엔터키로 검색되게
If lotteTmpGoodNo <> "" then
	Dim iA3, arrTemp3, arrlotteTmpGoodNo
	lotteTmpGoodNo = replace(lotteTmpGoodNo,",",chr(10))
	lotteTmpGoodNo = replace(lotteTmpGoodNo,chr(13),"")
	arrTemp3 = Split(lotteTmpGoodNo,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			If Not(isNumeric(trim(arrTemp3(iA3)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp3(iA3) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrlotteTmpGoodNo = arrlotteTmpGoodNo & trim(arrTemp3(iA3)) & ","
			End If
		End If
		iA3 = iA3 + 1
	Loop
	lotteTmpGoodNo = left(arrlotteTmpGoodNo,len(arrlotteTmpGoodNo)-1)
End If

Set oLotteitem = new CLotte
If (session("ssBctID")="kjy8517") or (session("ssBctID")="icommang") Then
	oLotteitem.FPageSize					= 100
Else
	oLotteitem.FPageSize					= 50
End If
	oLotteitem.FCurrPage					= page
	oLotteitem.FRectMakerid					= makerid
	oLotteitem.FRectItemID					= itemid
	oLotteitem.FRectItemName				= itemname
	oLotteitem.FRectLotteGoodNo				= lotteGoodNo
	oLotteitem.FRectLotteTmpGoodNo			= lotteTmpGoodNo
	oLotteitem.FRectCDL						= request("cdl")
	oLotteitem.FRectCDM						= request("cdm")
	oLotteitem.FRectCDS						= request("cds")
	oLotteitem.FRectExtNotReg				= ExtNotReg
	oLotteitem.FRectIsReged					= isReged
	oLotteitem.FRectNotinmakerid			= notinmakerid
	oLotteitem.FRectNotinitemid				= notinitemid
	oLotteitem.FRectExcTrans				= exctrans
	oLotteitem.FRectPriceOption				= priceOption
	oLotteitem.FRectIsSpecialPrice     		= isSpecialPrice
	oLotteitem.FRectDeliverytype			= deliverytype
	oLotteitem.FRectMwdiv					= mwdiv
	oLotteitem.FRectIsextusing				= isextusing

	oLotteitem.FRectSellYn					= sellyn
	oLotteitem.FRectLimitYn					= limityn
	oLotteitem.FRectSailYn					= sailyn
'	oLotteitem.FRectonlyValidMargin			= onlyValidMargin
	oLotteitem.FRectStartMargin				= startMargin
	oLotteitem.FRectEndMargin				= endMargin
	oLotteitem.FRectIsMadeHand				= isMadeHand
	oLotteitem.FRectIsOption				= isOption
	oLotteitem.FRectInfoDiv					= infoDiv
	oLotteitem.FRectExtSellYn				= extsellyn
	oLotteitem.FRectFailCntExists			= failCntExists
	oLotteitem.FRectMatchCate				= MatchCate
	oLotteitem.FRectExpensive10x10			= expensive10x10
	oLotteitem.FRectdiffPrc					= diffPrc
	oLotteitem.FRectLotteYes10x10No			= lotteYes10x10No
	oLotteitem.FRectLotteNo10x10Yes			= lotteNo10x10Yes
	oLotteitem.FRectReqEdit					= reqEdit
If (bestOrd = "on") Then
    oLotteitem.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oLotteitem.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oLotteitem.getLottereqExpireItemList
Else
	oLotteitem.getLotteRegedItemList		'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteCom","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteCom','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=lotteCom','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//추가금액 상품관리
function optAddpriceItemList(){
	var optwin3=window.open('/admin/etc/lotte/pop_AddPriceitem.asp','optAddpriceItemList','width=1500,height=800,scrollbars=yes,resizable=yes');
	optwin3.focus();
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteCom&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

// 롯데닷컴 카테고리 관리
function pop_CateManager() {
	var pCM = window.open("/admin/etc/lotte/popLotteCateList.asp","popCateMan","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
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
	if ((comp.name!="lotteYes10x10No")&&(frm.lotteYes10x10No.checked)){ frm.lotteYes10x10No.checked=false }
	if ((comp.name!="lotteNo10x10Yes")&&(frm.lotteNo10x10Yes.checked)){ frm.lotteNo10x10Yes.checked=false }
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

    if ((comp.name=="lotteYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="lotteNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.lotteYes10x10No.checked){
            comp.form.lotteYes10x10No.checked = false;
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
	if ((comp.name!="lotteYes10x10No")&&(frm.lotteYes10x10No.checked)){ frm.lotteYes10x10No.checked=false }
	if ((comp.name!="lotteNo10x10Yes")&&(frm.lotteNo10x10Yes.checked)){ frm.lotteNo10x10Yes.checked=false }
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

// 선택된 상품 판매여부 변경
function LotteSellYnProcess(chkYn) {
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

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 롯데닷컴에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }

        //document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function LotteSelectRegProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		//document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 등록예정 상품 일괄 등록
function LotteregIMSI(isreg) {
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
        if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?')){
			//document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 삭제 하시겠습니까?')){
			//document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
			document.frmSvArr.submit();
        }
    }
}

//선택 신규상품조회
function LotteStatCheckProcess(){
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 판매 상태를 확인 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function LottePriceEditProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 수정
function LotteImageEditProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품 이미지를 일괄 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

//선택된 상품 상품명 수정
function LotteItemnameEditProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품명을 일괄 수정 요청 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
//선택된 상품 재고조회
function LottecheckStockProcess(){
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 재고 조회 하시겠습니까?')){
		//document.getElementById("btnStock").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTOCK";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
function LotteInfodivEditProcess(){
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품의 품목을 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "INFODIV";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 카테고리 수정
function LotteCateEditProcess(){
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품의 카테고리를 수정 하시겠습니까?')){
		//document.getElementById("btnEditCate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CATEGORY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function LotteEditProcess(v) {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		if(v == ""){
			//document.getElementById("btnEditSel").disabled=true;
		}else{
			//document.getElementById("btnEditSel2").disabled=true;
		}
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT" + v;
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
		document.frmSvArr.submit();
    }
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//등록상품 조회 (기간별/상품등록일)
function LotteSearchGoods(){
	var popwin = window.open("<%=apiURL%>/outmall/checkRegItemList.asp?sellsite=lotteCom","checkRegItemList","width=800,height=400,scrollbars=yes,resizable=yes")
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
		<a href="https://partner.lotte.com/main/Login.lotte" target="_blank">롯데닷컴Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 124072 | store101010** ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		롯데닷컴 상품코드 : <textarea rows="2" cols="20" name="lotteGoodNo" id="itemid"><%=replace(lotteGoodNo,",",chr(10))%></textarea>
		&nbsp;
		승인전 상품코드 : <textarea rows="2" cols="20" name="lotteTmpGoodNo" id="itemid"><%=replace(lotteTmpGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >롯데닷컴 등록실패
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >롯데닷컴 등록예정
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >롯데닷컴 반려
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >롯데닷컴 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >롯데닷컴 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>롯데닷컴 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="lotteYes10x10No" <%= ChkIIF(lotteYes10x10No="on","checked","") %> ><font color=red>롯데닷컴판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="lotteNo10x10Yes" <%= ChkIIF(lotteNo10x10Yes="on","checked","") %> ><font color=red>롯데닷컴품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리(롯데닷컴은 원단위 안씀)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 옵션가 있는 상품, 단품상품에서 옵션추가된 상품

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
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();">&nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="추가금액상품관리" onclick="optAddpriceItemList();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lotteCom');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="LotteSelectRegProcess();">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSel" value="수정" onClick="LotteEditProcess('');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="LottePriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditName" value="상품명" onClick="LotteItemnameEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="재고조회" onClick="LottecheckStockProcess();">
				&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfoDiv" value="정부고시항목" onClick="LotteInfodivEditProcess();">
   				&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditCate" value="카테고리변경" onClick="LotteCateEditProcess();">
   				&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditImage" value="이미지" onClick="LotteImageEditProcess();">
				<br><br>
				승인예정 상품 :
				<input class="button" type="button" id="btnEditSel2" value="수정" onClick="LotteEditProcess('2');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnchk" value="신규상품조회" onClick="LotteStatCheckProcess();">
				<br><br>
				등록예정 상품 :
				<input class="button" type="button" id="btnRegImsi" value="등록" onClick="LotteregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="삭제" onClick="LotteregIMSI(false);" >
				<br><br>
				상품 조회 :
				<input class="button" type="button" id="btnSearchGoods" value="기간조회" onClick="LotteSearchGoods();">
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
					<option value="X">판매종료(삭제)</option><!-- 삭제하면 이후 수정 할 수 없음 -->
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="LotteSellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(oLotteitem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oLotteitem.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">롯데닷컴등록일<br>롯데닷컴최종수정일<br><font color="purple">카테고리최종변경일</font></td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">롯데닷컴<br>가격및판매</td>
	<td width="70">롯데닷컴<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i = 0 To oLotteitem.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oLotteitem.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oLotteitem.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oLotteitem.FItemList(i).FItemID %>','lotteCom','<%=oLotteitem.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oLotteitem.FItemList(i).FItemID%>" target="_blank"><%= oLotteitem.FItemList(i).FItemID %></a>
		<% if oLotteitem.FItemList(i).FLimitYn="Y" then %><br><%= oLotteitem.FItemList(i).getLimitHtmlStr %></font><% end if %>
	</td>
	<td align="left"><%= oLotteitem.FItemList(i).FMakerid %><%= oLotteitem.FItemList(i).getDeliverytypeName %><br><%= oLotteitem.FItemList(i).FItemName %></td>
	<td align="center"><%= oLotteitem.FItemList(i).FRegdate %><br><%= oLotteitem.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oLotteitem.FItemList(i).FLotteRegdate %><br><%= oLotteitem.FItemList(i).FLotteLastUpdate %>
		<% If oLotteitem.FItemList(i).FLastcateChgDate <> "1900-01-01" Then %>
		<br><font color="purple"><%= oLotteitem.FItemList(i).FLastcateChgDate %></font>
		<% End If %>
	</td>

	<td align="right">
	<% If oLotteitem.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oLotteitem.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
		<%
		If oLotteitem.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (oLotteitem.FItemList(i).FLotteStatCd > 0) and Not IsNULL(oLotteitem.FItemList(i).FLottePrice) Then
		' 	If (oLotteitem.FItemList(i).FSaleYn = "Y") and (oLotteitem.FItemList(i).FSellcash < oLotteitem.FItemList(i).FLottePrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).FLottePrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oLotteitem.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oLotteitem.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oLotteitem.FItemList(i).FOrgSuplycash/oLotteitem.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oLotteitem.FItemList(i).IsSoldOut Then
			If oLotteitem.FItemList(i).FSellyn = "N" Then
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
		If oLotteitem.FItemList(i).FItemdiv = "06" OR oLotteitem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oLotteitem.FItemList(i).FLotteStatCd > 0) Then
			If Not IsNULL(oLotteitem.FItemList(i).FLottePrice) Then
				If (oLotteitem.FItemList(i).Fsellcash <> oLotteitem.FItemList(i).FLottePrice) Then
	%>
					<strong><%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oLotteitem.FItemList(i).FLottePrice,0)&"<br>"
				End If

				If Not IsNULL(oLotteitem.FItemList(i).FSpecialPrice) Then
					If (now() >= oLotteitem.FItemList(i).FStartDate) And (now() <= oLotteitem.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oLotteitem.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oLotteitem.FItemList(i).FSellyn="Y" and oLotteitem.FItemList(i).FLotteSellYn<>"Y") or (oLotteitem.FItemList(i).FSellyn<>"Y" and oLotteitem.FItemList(i).FLotteSellYn="Y") Then
	%>
					<strong><%= oLotteitem.FItemList(i).FLotteSellYn %></strong>
	<%
				Else
					response.write oLotteitem.FItemList(i).FLotteSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
    <%
    	'#실상품번호
    	if Not(IsNULL(oLotteitem.FItemList(i).FLotteGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotte.com/goods/viewGoodsDetail.lotte?goods_no="&oLotteitem.FItemList(i).FLotteGoodNo&"'>"&oLotteitem.FItemList(i).FLotteGoodNo&"</a>"
		else
			'#임시상품번호
			if Not(IsNULL(oLotteitem.FItemList(i).FLotteTmpGoodNo)) then
				if oLotteitem.FItemList(i).FLotteStatCd<>"30" then
					Response.Write oLotteitem.FItemList(i).getLotteItemStatCd & "<br>(" & oLotteitem.FItemList(i).FLotteTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if
	%>
	</td>
	<td align="center"><%= oLotteitem.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','0');"><%= oLotteitem.FItemList(i).FoptionCnt %>:<%= oLotteitem.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oLotteitem.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oLotteitem.FItemList(i).FCateMapCnt > 0 Then %>
	    매칭됨
	<% Else %>
		<font color="darkred">매칭안됨</font>
	<% End If %>

	<% If (oLotteitem.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oLotteitem.FItemList(i).FlastErrStr %>">ERR:<%= oLotteitem.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
    <td align="center"><%= oLotteitem.FItemList(i).FinfoDiv %>
    <% if (oLotteitem.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oLotteitem.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">옵션금액</font>
	    <% If oLotteitem.FItemList(i).FoptAddPrcRegType<>0 Then %>
	    (<%=oLotteitem.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% End If %>
    </td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oLotteitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotteitem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLotteitem.StartScrollPage to oLotteitem.FScrollCount + oLotteitem.StartScrollPage - 1 %>
    		<% if i>oLotteitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLotteitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oLotteitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
