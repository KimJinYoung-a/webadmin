<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/auction/auctioncls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, auctionGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, auctionYes10x10No, auctionNo10x10Yes, auctionKeepSell, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice
Dim page, i, research
Dim oAuction, AuctionGoodNoArray, isextusing, cisextusing, rctsellcnt, scheduleNotInItemid
dim startsell, stopsell, xl
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
auctionGoodNo			= request("auctionGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
auctionYes10x10No		= request("auctionYes10x10No")
auctionNo10x10Yes		= request("auctionNo10x10Yes")
auctionKeepSell			= request("auctionKeepSell")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
startsell				= requestCheckVar(request("startsell"), 1)
stopsell				= requestCheckVar(request("stopsell"), 1)
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
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

	if (stopsell = "Y") then
		'// 판매중지 대상 상품목록
		ExtNotReg = "D"
		sellyn = "N"
		extsellyn = "Y"
		auctionYes10x10No = "on"
		onlyValidMargin = ""
	elseif (startsell = "Y") then
		'// 판매전환 대상 상품목록
		ExtNotReg = "D"
		sellyn = "Y"
		extsellyn = "N"
		onlyValidMargin = "Y"
		notinmakerid = "N"
		notinitemid = "N"
		auctionNo10x10Yes = "on"
	end if
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
'옥션 상품코드 엔터키로 검색되게
If auctionGoodNo <> "" then
	Dim iA2, arrTemp2, arrauctionGoodNo
	auctionGoodNo = replace(auctionGoodNo,",",chr(10))
	auctionGoodNo = replace(auctionGoodNo,chr(13),"")
	arrTemp2 = Split(auctionGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrauctionGoodNo = arrauctionGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	auctionGoodNo = left(arrauctionGoodNo,len(arrauctionGoodNo)-1)
End If

Set oAuction = new CAuction
	oAuction.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oAuction.FPageSize					= 100

Else
	oAuction.FPageSize					= 50
End If
	oAuction.FRectCDL					= request("cdl")
	oAuction.FRectCDM					= request("cdm")
	oAuction.FRectCDS					= request("cds")
	oAuction.FRectItemID				= itemid
	oAuction.FRectItemName				= itemname
	oAuction.FRectSellYn				= sellyn
	oAuction.FRectLimitYn				= limityn
	oAuction.FRectSailYn				= sailyn
'	oAuction.FRectonlyValidMargin		= onlyValidMargin
	oAuction.FRectStartMargin			= startMargin
	oAuction.FRectEndMargin				= endMargin
	oAuction.FRectMakerid				= makerid
	oAuction.FRectAuctionGoodNo			= auctionGoodNo
	oAuction.FRectMatchCate				= MatchCate
	oAuction.FRectIsMadeHand			= isMadeHand
	oAuction.FRectIsOption				= isOption
	oAuction.FRectIsReged				= isReged
	oAuction.FRectNotinmakerid			= notinmakerid
	oAuction.FRectNotinitemid			= notinitemid
	oAuction.FRectExcTrans				= exctrans
	oAuction.FRectPriceOption			= priceOption
	oAuction.FRectIsSpecialPrice       	= isSpecialPrice
	oAuction.FRectDeliverytype			= deliverytype
	oAuction.FRectMwdiv					= mwdiv
	oAuction.FRectScheduleNotInItemid	= scheduleNotInItemid
	oAuction.FRectIsextusing			= isextusing
	oAuction.FRectCisextusing			= cisextusing
	oAuction.FRectRctsellcnt			= rctsellcnt

	oAuction.FRectExtNotReg				= ExtNotReg
	oAuction.FRectExpensive10x10		= expensive10x10
	oAuction.FRectdiffPrc				= diffPrc
	oAuction.FRectAuctionYes10x10No		= auctionYes10x10No
	oAuction.FRectAuctionNo10x10Yes		= auctionNo10x10Yes
	oAuction.FRectAuctionKeepSell		= auctionKeepSell
	oAuction.FRectExtSellYn				= extsellyn
	oAuction.FRectInfoDiv				= infoDiv
	oAuction.FRectFailCntOverExcept		= ""
	oAuction.FRectFailCntExists			= failCntExists
	oAuction.FRectReqEdit				= reqEdit
	oAuction.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oAuction.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oAuction.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oAuction.getAuctionreqExpireItemList
Else
	oAuction.getAuctionRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=auction1010List"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=auction1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=auction1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=auction1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=auction1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="auctionKeepSell")&&(frm.auctionKeepSell.checked)){ frm.auctionKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="auctionYes10x10No")&&(frm.auctionYes10x10No.checked)){ frm.auctionYes10x10No.checked=false }
	if ((comp.name!="auctionNo10x10Yes")&&(frm.auctionNo10x10Yes.checked)){ frm.auctionNo10x10Yes.checked=false }
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

    if ((comp.name=="auctionYes10x10No")&&(comp.checked)){
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
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="auctionNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.auctionYes10x10No.checked){
            comp.form.auctionYes10x10No.checked = false;
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

    if ((comp.name=="auctionKeepSell")&&(comp.checked)){
        if (comp.form.auctionYes10x10No.checked){
            comp.form.auctionYes10x10No.checked = false;
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
	if ((comp.name!="auctionKeepSell")&&(frm.auctionKeepSell.checked)){ frm.auctionKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="auctionYes10x10No")&&(frm.auctionYes10x10No.checked)){ frm.auctionYes10x10No.checked=false }
	if ((comp.name!="auctionNo10x10Yes")&&(frm.auctionNo10x10Yes.checked)){ frm.auctionNo10x10Yes.checked=false }
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
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/auction/popauctioncateList.asp","popCateAuctionmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// 선택된 상품 기본정보 등록
function AuctionSelectRegProcess() {
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품의 기본정보를 등록 하시겠습니까?\n\n옥션코드를 리턴받기 위한 기본정보 등록입니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddItem";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 옵션정보 등록
function AuctionSelectOPTProcess() {
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품의 옵션정보를 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 정보고시 등록
function AuctionSelectInfoCdRegProcess() {
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품의 정보고시를 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGInfoCd";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 기본정보 + 옵션 + 고시정보 등록
function AuctionREGProcess() {
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 정보고시 등록
function AuctionOnSaleEditProcess() {
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품상태를 판매중으로 변경 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOnSale";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 상태 변경
function AuctionSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//기본정보 수정
function AuctionEditInfoProcess(){
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 상품정보를 수정 하시겠습니까?\n\n※상품명, 가격, 이미지, 상품상세등이 수정됩니다')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditInfo";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//옵션 조회
function AuctionViewProcess(){
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 옵션을 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//옵션 수정
function AuctionEditOPTProcess(){
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 옵션을 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//기본정보 + 옵션정보 수정
function AuctionEditProcess(){
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

    if (confirm('Auction에 선택하신 ' + chkSel + '개 기본정보 + 옵션 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 삭제
function AuctionDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\n옥션 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
			document.frmSvArr.submit();
		}
    }
}

//공통코드 검색
function AuctionCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('공통코드를 선택하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		xLink.location.href = "<%=apiURL%>/outmall/auction/actauctionReq.asp?cmdparam=auctionCommonCode&CommCD="+ccd+"";
	}
}


//공통코드 검색
function fnAuctionCommCDSubmit() {
	var ccd;
	var goodsGrpCd;
	ccd = document.getElementById('CommCD2').value;
	//goodsGrpCd = $("#goodsGrpCd option:selected").val();

	goodsGrpCd = $("#goodsGrpCd").val();
	if(ccd == ''){
		alert('공통코드를 선택하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		xLink.location.href = "/admin/etc/auction/actauctionReq.asp?cmdparam=ebayCommonCode&CommCD="+ccd+"&goodsGrpCd="+goodsGrpCd;
	}
}
function jsByValue(s){
	if((s == "brand") || (s == "maker") || (s =="placepolicy" || s == "infocodedtl" || s == "mastercode" || s == "sitecode")) {
		$("#goodsGrpCd_span").show();
	}else{
		$("#goodsGrpCd_span").hide();
	}
}

//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=auction1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
//제조, 유효일 등록 팝업
function popAuctionDate(iitemid){
    var pdate = window.open("/admin/etc/auction/popAuctionDate.asp?itemid="+iitemid+'&mallid=auction1010',"popAuctionDate","width=500,height=200,scrollbars=yes,resizable=yes");
	pdate.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
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
		<a href="http://www.esmplus.com/Home/Home" target="_blank">옥션Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ 10x10store | Cube1010!* ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		옥션 상품코드 : <textarea rows="2" cols="20" name="auctionGoodNo" id="itemid"><%= replace(replace(auctionGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >옥션 등록성공_OnSale전
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >옥션 전송시도 중 오류
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >옥션 등록예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >옥션 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>옥션 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionYes10x10No" <%= ChkIIF(auctionYes10x10No="on","checked","") %> ><font color=red>옥션판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionNo10x10Yes" <%= ChkIIF(auctionNo10x10Yes="on","checked","") %> ><font color=red>옥션품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionKeepSell" <%= ChkIIF(auctionKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 상품설명에 IFRAME TAG 사용한 상품

<p />

<form name="frmReg" method="post" action="auctionitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="등록 제외 카테고리" onclick="NotInCategory();">
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('auction1010');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="기본정보" onClick="AuctionSelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOPTSel" value="옵션정보" onClick="AuctionSelectOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnInfocdSel" value="상품고시" onClick="AuctionSelectInfoCdRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnREG" value="기본+옵션+고시" onClick="AuctionREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnOnSale" value="OnSale변경" onClick="AuctionOnSaleEditProcess();" style=color:red;font-weight:bold>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditInfo" value="기본정보(가격)" onClick="AuctionEditInfoProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEditOPT" value="옵션정보" onClick="AuctionEditOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEdit" value="기본+옵션" onClick="AuctionEditProcess();" style=color:blue;font-weight:bold>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="nys1006") OR (session("ssBctID")="z0516") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="상품삭제" onClick="AuctionDeleteProcess();" style=font-weight:bold>
			<% End If %>
				<br><br>
				실제상품 조회 :
				<input class="button" type="button" id="btnViewItem" value="옵션조회" onClick="AuctionViewProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				<br><br>
				공통코드 검색 :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="GetShippingPlaceCode">출하지코드
					<option value="GetNationCode">원산지코드
					<option value="GetDeliveryList">배송사(택배)조회
					<option value="GetPaidOrderList">주문Test
				</select>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="AuctionCommCDSubmit();" >

				2.0 : 
				<select name="CommCD2" class="select" id="CommCD2" onChange="jsByValue(this.value);">
					<option value="">- Choice -
					<option value="mastercode">마스터코드조회(옥션코드로)</option>
					<option value="sitecode">옥션코드조회(마스터코드로)</option>
				</select>
				<span id="goodsGrpCd_span" style="display:none;">
					<input type="text" name="goodsGrpCd" id="goodsGrpCd">
				</span>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="fnAuctionCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">판매중지</option>
					<option value="Y">판매</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="AuctionSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<!-- 리스트 시작 -->
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
		검색결과 : <b><%= FormatNumber(oAuction.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oAuction.FTotalPage,0) %></b>
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
	<td width="140">Auction등록일<br>Auction최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Auction<br>가격및판매</td>
	<td width="70">Auction<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
	<td width="100">기|옵|상<br>OnSale변경일</td>
</tr>
<% For i=0 to oAuction.FResultCount - 1 %>
<tr align="center" <% If oAuction.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oAuction.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oAuction.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oAuction.FItemList(i).FItemID %>','auction1010','<%=oAuction.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oAuction.FItemList(i).FItemID%>" target="_blank"><%= oAuction.FItemList(i).FItemID %></a>
		<%
			If (xl <> "Y") Then
				If oAuction.FItemList(i).FAuctionStatcd <> 7 Then
		%>
		<br><%= oAuction.FItemList(i).getAuctionStatName %>
		<%
				End If
				response.write oAuction.FItemList(i).getLimitHtmlStr
			End If
		%>
	</td>
	<td align="left"><%= oAuction.FItemList(i).FMakerid %> <%= oAuction.FItemList(i).getDeliverytypeName %><br><%= oAuction.FItemList(i).FItemName %></td>
	<td align="center"><%= oAuction.FItemList(i).FRegdate %><br><%= oAuction.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oAuction.FItemList(i).FAuctionRegdate %><br><%= oAuction.FItemList(i).FAuctionLastUpdate %></td>
	<td align="right">
		<% If oAuction.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oAuction.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oAuction.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oAuction.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oAuction.FItemList(i).Fsellcash = 0 Then
		elseif (oAuction.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oAuction.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oAuction.FItemList(i).FOrgSuplycash/oAuction.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oAuction.FItemList(i).Fbuycash/oAuction.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oAuction.FItemList(i).Fbuycash/oAuction.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oAuction.FItemList(i).IsSoldOut Then
			If oAuction.FItemList(i).FSellyn = "N" Then
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
		If oAuction.FItemList(i).FItemdiv = "06" OR oAuction.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oAuction.FItemList(i).FAuctionStatCd > 0) Then
			If Not IsNULL(oAuction.FItemList(i).FAuctionPrice) Then
				If (oAuction.FItemList(i).Mustprice <> oAuction.FItemList(i).FAuctionPrice) Then
	%>
					<strong><%= formatNumber(oAuction.FItemList(i).FAuctionPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oAuction.FItemList(i).FAuctionPrice,0)&"<br>"
				End If

				If Not IsNULL(oAuction.FItemList(i).FSpecialPrice) Then
					If (now() >= oAuction.FItemList(i).FStartDate) And (now() <= oAuction.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oAuction.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oAuction.FItemList(i).FSellyn="Y" and oAuction.FItemList(i).FAuctionSellYn<>"Y") or (oAuction.FItemList(i).FSellyn<>"Y" and oAuction.FItemList(i).FAuctionSellYn="Y") Then
	%>
					<strong><%= oAuction.FItemList(i).FAuctionSellYn %></strong>
	<%
				Else
					response.write oAuction.FItemList(i).FAuctionSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oAuction.FItemList(i).FAuctionGoodNo)) Then
			Response.Write "<a target='_blank' href='http://itempage3.auction.co.kr/detailview.aspx?itemNo="&oAuction.FItemList(i).FAuctionGoodNo&"'>"&oAuction.FItemList(i).FAuctionGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oAuction.FItemList(i).Freguserid %></td>
	<td align="center">
		<a href="javascript:popManageOptAddPrc('<%=oAuction.FItemList(i).FItemID%>','0');"><%= oAuction.FItemList(i).FoptionCnt %>:<%= oAuction.FItemList(i).FregedOptCnt %></a>
		<br>
		<input type="button" class="button" value="일자" onclick="javascript:popAuctionDate('<%=oAuction.FItemList(i).FItemID%>');">
	</td>
	<td align="center"><%= oAuction.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oAuction.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oAuction.FItemList(i).FinfoDiv %>
		<%
		If (oAuction.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oAuction.FItemList(i).FlastErrStr) &"'>ERR:"& oAuction.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oAuction.FItemList(i).FAPIadditem="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIadditem&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIadditem&"</font>") %>&nbsp;|
		<%= Chkiif(oAuction.FItemList(i).FAPIaddopt="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIaddopt&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIaddopt&"</font>") %>&nbsp;|
		<%= Chkiif(oAuction.FItemList(i).FAPIaddgosi="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIaddgosi&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIaddgosi&"</font>") %>
		<br>
		<%= oAuction.FItemList(i).FOnSaleRegdate %>
	</td>
</tr>
<% AuctionGoodNoArray = AuctionGoodNoArray & oAuction.FItemList(i).FAuctionGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= AuctionGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oAuction.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAuction.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oAuction.StartScrollPage to oAuction.FScrollCount + oAuction.StartScrollPage - 1 %>
    		<% if i>oAuction.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oAuction.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
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
	<input type="hidden" name="auctionGoodNo" value= <%= auctionGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="auctionYes10x10No" value= <%= auctionYes10x10No %>>
	<input type="hidden" name="auctionNo10x10Yes" value= <%= auctionNo10x10Yes %>>
	<input type="hidden" name="auctionKeepSell" value= <%= auctionKeepSell %>>
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
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
	<input type="hidden" name="startsell" value= <%= startsell %>>
	<input type="hidden" name="stopsell" value= <%= stopsell %>>
	<input type="hidden" name="rctsellcnt" value= <%= rctsellcnt %>>
	<input type="hidden" name="purchasetype" value= <%= purchasetype %>>
</form>
<% SET oAuction = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
