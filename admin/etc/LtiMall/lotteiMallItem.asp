<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, ltimallgoodno, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, ltimalltmpgoodno, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, ltimallYes10x10No, ltimallNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing, cisextusing, rctsellcnt
Dim page, i, research, ltimallGoodNoArray
Dim oiMall
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
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
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
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid="1442823,1464139,1471535,1471538,1471539,1471617,1471618"
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
	oiMall.FPageSize					= 100
Else
	oiMall.FPageSize					= 50
End If
	oiMall.FCurrPage					= page
	oiMall.FRectMakerid					= makerid
	oiMall.FRectItemID					= itemid
	oiMall.FRectItemName				= itemname
	oiMall.FRectLTiMallGoodNo			= ltimallgoodno
	oiMall.FRectLTiMallTmpGoodNo		= ltimalltmpgoodno
	oiMall.FRectCDL						= request("cdl")
	oiMall.FRectCDM						= request("cdm")
	oiMall.FRectCDS						= request("cds")
	oiMall.FRectExtNotReg				= ExtNotReg
	oiMall.FRectIsReged					= isReged
	oiMall.FRectNotinmakerid			= notinmakerid
	oiMall.FRectNotinitemid				= notinitemid
	oiMall.FRectExcTrans				= exctrans
	oiMall.FRectPriceOption				= priceOption
	oiMall.FRectIsSpecialPrice     		= isSpecialPrice
	oiMall.FRectDeliverytype			= deliverytype
	oiMall.FRectMwdiv					= mwdiv
	oiMall.FRectIsextusing				= isextusing
	oiMall.FRectCisextusing				= cisextusing
	oiMall.FRectRctsellcnt				= rctsellcnt

	oiMall.FRectSellYn					= sellyn
	oiMall.FRectLimitYn					= limityn
	oiMall.FRectSailYn					= sailyn
	oiMall.FRectStartMargin				= startMargin
	oiMall.FRectEndMargin				= endMargin
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
	oiMall.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oiMall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oiMall.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oiMall.getLtiMallreqExpireItemList
Else
	oiMall.getLTiMallRegedItemList			'그 외 리스트
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteimall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteimall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=lotteimall','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//추가금액 상품관리
function optAddpriceItemList(){
	var optwin3=window.open('/admin/etc/Ltimall/pop_AddPriceitem.asp','optAddpriceItemList','width=1500,height=800,scrollbars=yes,resizable=yes');
	optwin3.focus();
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
			comp.form.notinmakerid.value = "";
			comp.form.notinitemid.value = "";
			comp.form.exctrans.value = "N";
			comp.form.failCntExists.value = "N";
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
// 롯데iMall 담당MD 목록
function pop_MDList() {
	var pMD = window.open("/admin/etc/Ltimall/popLTiMallMDList.asp","popMDListIMall","width=600,height=300,scrollbars=yes,resizable=yes");
	pMD.focus();
}

// 롯데iMall 카테고리 관리
function pop_CateManager() {
	var pCM = window.open("/admin/etc/Ltimall/popLTiMallCateList.asp","popCateManIMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
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
		//document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
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
        //document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
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
       //document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
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
			//document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
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
        //document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
//선택상품 재고조회
function LtimallcheckStockProcess(){
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

    if (confirm('롯데아이몰에 선택하신 ' + chkSel + '개 상품을 재고 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "ChkStockSelect";
         document.frmSvArr.cmdparam.value = "CHKSTOCK";
        //document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
//선택상품 재고조회
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
//상품 삭제
function LtimallDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
			document.frmSvArr.submit();
		}
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
		//document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
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
			//document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 삭제 하시겠습니까?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			//document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }
}

//등록상품 조회 (기간별/상품등록일)
function LtimallSearchGoods(){
	var popwin = window.open("<%=apiURL%>/outmall/checkRegItemList.asp?sellsite=lotteimall","checkRegItemList_ltimall","width=800,height=400,scrollbars=yes,resizable=yes")
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
		<a href="https://partners.lotteimall.com/" target="_blank">롯데아이몰Admin바로가기</a>
		<%
			If C_ADMIN_AUTH OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ 011799LT | store101010*! | 01068551098 ]</font>"
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
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
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

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리(롯데아이몰은 원단위 안씀)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 화장품, 식품류 제외, 단품이었다 옵션추가된 상품, 옵션가 있는 상품

<p />

<!-- 액션 시작 -->
<form name="frmReg" method="post" action="lotteimallItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="추가금액상품관리" onclick="optAddpriceItemList();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lotteimall');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="담당MD" onclick="pop_MDList();"> &nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="LtimallSelectRegProcess(true);">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="LtimallPriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="정보&가격&옵션&재고&상태" onClick="LtimallEditProcess('');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditName" value="상품명" onClick="LtimallItemnameEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="재고조회" onClick="LtimallcheckStockProcess();">
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="전시상품조회" onClick="LtimallDispViewProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnODelete" value="상품삭제" onClick="LtimallDeleteProcess();" style=font-weight:bold>
				<% End If %>
				<br><br>
				승인예정 상품 :
				<input class="button" type="button" id="btnEditSel2" value="수정" onClick="LtimallEditProcess('2');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditInfoDiv" value="신규상품조회" onClick="LtimallStatCheckProcess();">
				<br><br>
				등록예정 상품 :
				<input class="button" type="button" id="btnRegImsi" value="등록" onClick="LtimallregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="삭제" onClick="LtimallregIMSI(false);" >
				<br><br>
				상품 조회 :
				<input class="button" type="button" id="btnSearchGoods" value="기간조회" onClick="LtimallSearchGoods();">
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
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">롯데iMall등록일<br>롯데iMall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">롯데iMall<br>가격및판매</td>
	<td width="70">롯데iMall<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i = 0 To oiMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oiMall.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oiMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oiMall.FItemList(i).FItemID %>','lotteimall','<%=oiMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oiMall.FItemList(i).FItemID%>" target="_blank"><%= oiMall.FItemList(i).FItemID %></a>
		<% if oiMall.FItemList(i).FLimitYn="Y" then %><br><%= oiMall.FItemList(i).getLimitHtmlStr %></font><% end if %>
	</td>
	<td align="left"><%= oiMall.FItemList(i).FMakerid %><%= oiMall.FItemList(i).getDeliverytypeName %><br><%= oiMall.FItemList(i).FItemName %></td>
	<td align="center"><%= oiMall.FItemList(i).FRegdate %><br><%= oiMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oiMall.FItemList(i).FLtimallRegdate %><br><%= oiMall.FItemList(i).FLtimallLastUpdate %></td>

	<td align="right">
	<% If oiMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oiMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
		<%
		If oiMall.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (oiMall.FItemList(i).FLtimallStatCd > 0) and Not IsNULL(oiMall.FItemList(i).FLtimallPrice) Then
		' 	If (oiMall.FItemList(i).FSaleYn = "Y") then and (oiMall.FItemList(i).FSellcash < oiMall.FItemList(i).FLtimallPrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).FLtimallPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oiMall.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oiMall.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oiMall.FItemList(i).FOrgSuplycash/oiMall.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
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
				If (oiMall.FItemList(i).Fsellcash <> oiMall.FItemList(i).FLtimallPrice) Then
	%>
					<strong><%= formatNumber(oiMall.FItemList(i).FLtimallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oiMall.FItemList(i).FLtimallPrice,0)&"<br>"
				End If

				If Not IsNULL(oiMall.FItemList(i).FSpecialPrice) Then
					If (now() >= oiMall.FItemList(i).FStartDate) And (now() <= oiMall.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oiMall.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
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
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oiMall.FItemList(i).FItemID%>','0');"><%= oiMall.FItemList(i).FoptionCnt %>:<%= oiMall.FItemList(i).FregedOptCnt %></a></td>
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
<% ltimallGoodNoArray = ltimallGoodNoArray & oiMall.FItemList(i).FLtiMallGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= ltimallGoodNoArray %></textarea>
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
