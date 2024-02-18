<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/lfmall/lfmallCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, lfmallKeepSell, isSpecialPrice
Dim bestOrdMall, lfmallGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans, MatchDiv
Dim expensive10x10, diffPrc, lfmallYes10x10No, lfmallNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, scheduleNotInItemid
Dim page, i, research, isextusing, cisextusing, rctsellcnt
Dim oLfmall, xl
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
lfmallGoodNo			= request("lfmallGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchDiv				= request("MatchDiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
lfmallYes10x10No		= request("lfmallYes10x10No")
lfmallNo10x10Yes		= request("lfmallNo10x10Yes")
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
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
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
	MatchDiv = ""
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

'LFmall 상품코드 엔터키로 검색되게
If lfmallGoodNo <> "" then
	Dim iA2, arrTemp2, arrlfmallGoodNo
	lfmallGoodNo = replace(lfmallGoodNo,",",chr(10))
	lfmallGoodNo = replace(lfmallGoodNo,chr(13),"")
	arrTemp2 = Split(lfmallGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrlfmallGoodNo = arrlfmallGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	lfmallGoodNo = left(arrlfmallGoodNo,len(arrlfmallGoodNo)-1)
End If

Set oLfmall = new CLfmall
	oLfmall.FCurrPage					= page
	oLfmall.FPageSize					= 100
	oLfmall.FRectCDL					= request("cdl")
	oLfmall.FRectCDM					= request("cdm")
	oLfmall.FRectCDS					= request("cds")
	oLfmall.FRectItemID					= itemid
	oLfmall.FRectItemName				= itemname
	oLfmall.FRectSellYn					= sellyn
	oLfmall.FRectLimitYn				= limityn
	oLfmall.FRectSailYn					= sailyn
	oLfmall.FRectStartMargin			= startMargin
	oLfmall.FRectEndMargin				= endMargin
	oLfmall.FRectMakerid				= makerid
	oLfmall.FRectlfmallGoodNo			= lfmallGoodNo
	oLfmall.FRectMatchCate				= MatchCate
	oLfmall.FRectMatchDiv				= MatchDiv
	oLfmall.FRectIsMadeHand				= isMadeHand
	oLfmall.FRectIsOption				= isOption
	oLfmall.FRectIsReged				= isReged
	oLfmall.FRectNotinmakerid			= notinmakerid
	oLfmall.FRectNotinitemid			= notinitemid
	oLfmall.FRectScheduleNotInItemid	= scheduleNotInItemid
	oLfmall.FRectIsextusing				= isextusing
	oLfmall.FRectCisextusing			= cisextusing
	oLfmall.FRectRctsellcnt				= rctsellcnt

	oLfmall.FRectExcTrans				= exctrans
	oLfmall.FRectPriceOption			= priceOption
	oLfmall.FRectIsSpecialPrice     	= isSpecialPrice
	oLfmall.FRectDeliverytype			= deliverytype
	oLfmall.FRectMwdiv					= mwdiv

	oLfmall.FRectExtNotReg				= ExtNotReg
	oLfmall.FRectExpensive10x10			= expensive10x10
	oLfmall.FRectdiffPrc				= diffPrc
	oLfmall.FRectlfmallYes10x10No		= lfmallYes10x10No
	oLfmall.FRectlfmallNo10x10Yes		= lfmallNo10x10Yes
	oLfmall.FRectlfmallKeepSell			= lfmallKeepSell
	oLfmall.FRectExtSellYn				= extsellyn
	oLfmall.FRectInfoDiv				= infoDiv
	oLfmall.FRectFailCntOverExcept		= ""
	oLfmall.FRectFailCntExists			= failCntExists
	oLfmall.FRectReqEdit				= reqEdit
	oLfmall.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oLfmall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oLfmall.FRectOrdType = "BM"
End If

If isReged = "R" Then					'품절처리요망 상품보기 리스트
	oLfmall.getLfmallreqExpireItemList
Else
	oLfmall.getLfmallRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=lfmallList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lfmall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lfmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=lfmall','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=lfmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 전시제외 상품
function DisplayNotInItemid(){
	var popwin=window.open('/admin/etc/display_Not_In_Itemid.asp?mallgubun=lfmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//키워드 관리
function popKeywordItemList(){
	var popwin=window.open('/admin/etc/common/popKeywordList.asp?mallgubun=lfmall','popKeywordItemList','width=1300,height=700,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="lfmallKeepSell")&&(frm.lfmallKeepSell.checked)){ frm.lfmallKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="lfmallYes10x10No")&&(frm.lfmallYes10x10No.checked)){ frm.lfmallYes10x10No.checked=false }
	if ((comp.name!="lfmallNo10x10Yes")&&(frm.lfmallNo10x10Yes.checked)){ frm.lfmallNo10x10Yes.checked=false }
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

    if ((comp.name=="lfmallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="lfmallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.lfmallYes10x10No.checked){
            comp.form.lfmallYes10x10No.checked = false;
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
	if ((comp.name!="lfmallYes10x10No")&&(frm.lfmallYes10x10No.checked)){ frm.lfmallYes10x10No.checked=false }
	if ((comp.name!="lfmallNo10x10Yes")&&(frm.lfmallNo10x10Yes.checked)){ frm.lfmallNo10x10Yes.checked=false }
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
//상품분류 관리
function pop_NotiManager() {
	var pCM1 = window.open("/admin/etc/lfmall/poplfmallnotiList.asp","popNotilfmallmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM1.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/lfmall/poplfmallcateList.asp","popCatelfmallmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// 선택된 상품 일괄 등록
function lfmallSelectRegProcess() {
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

    if (confirm('LFmall에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "/admin/etc/lfmall/actlfmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function lfmallSellYnProcess(chkYn) {
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
		case "N": strSell="일시품절";break;
		case "X": strSell="상품정지(삭제)";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        //document.frmSvArr.action = "/admin/etc/lfmall/actlfmallReq.asp"
		document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 조회
function lfmallcheckStatProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 재고를 조회 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/lfmall/actlfmallReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function lfmallEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "/admin/etc/lfmall/actlfmallReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 등록
function lfmallNewSelectRegProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function lfmallNewcheckStatProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 조회 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function lfmallNewEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※LFmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lfmall/actlfmallReq.asp"
        document.frmSvArr.submit();
    }
}

//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_lfmall.asp?itemid="+iitemid+'&mallid=lfmall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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
		<a href="https://partner.lfmall.co.kr/front/login?redirect=%2Fhome" target="_blank">LFmallAdmin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenten | Ten1010*!! ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		LFmall 상품코드 : <textarea rows="2" cols="20" name="lfmallGoodNo" id="itemid"><%= replace(replace(lfmallGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >LFmall 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >LFmall 등록예정이상
            <option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >LFmall 등록후 승인대기
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >LFmall 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>LFmall 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="lfmallYes10x10No" <%= ChkIIF(lfmallYes10x10No="on","checked","") %> ><font color=red>LFmall판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="lfmallNo10x10Yes" <%= ChkIIF(lfmallNo10x10Yes="on","checked","") %> ><font color=red>LFmall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="lfmallKeepSell" <%= ChkIIF(lfmallKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />
* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
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
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">&nbsp;
				<input class="button" type="button" value="키워드" onclick="popKeywordItemList();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lfmall');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="품목분류" onclick="pop_NotiManager();">&nbsp;
				<input class="button" type="button" value="표준카테" onclick="pop_CateManager();">
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
				<input class="button" type="button" id="btnRegSel2" value="등록" onClick="lfmallNewSelectRegProcess();">&nbsp;&nbsp;
		<%
			If (session("ssBctID")="kjy8517") Then
		%>
				<input class="button" type="button" id="btnRegSel" value="등록(구)" onClick="lfmallSelectRegProcess();">&nbsp;&nbsp;
		<%
			End If
		%>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEdit2" value="수정" onClick="lfmallNewEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock2" value="조회" onClick="lfmallNewcheckStatProcess();">&nbsp;&nbsp;
		<%
			If (session("ssBctID")="kjy8517") Then
		%>
				<input class="button" type="button" id="btnEdit" value="수정(구)" onClick="lfmallEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="조회(구)" onClick="lfmallcheckStatProcess();">&nbsp;&nbsp;
		<%
			End If
		%>

			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="lfmallSellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(oLfmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oLfmall.FTotalPage,0) %></b>
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
	<td width="140">LFmall등록일<br>LFmall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">LFmall<br>가격및판매</td>
	<td width="70">LFmall<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oLfmall.FResultCount - 1 %>
<tr align="center" <% If oLfmall.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %> >
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oLfmall.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oLfmall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oLfmall.FItemList(i).FItemID %>','lfmall','<%=oLfmall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oLfmall.FItemList(i).FItemID%>" target="_blank"><%= oLfmall.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oLfmall.FItemList(i).FLfmallStatcd <> 7 Then
	%>
		<br><%= oLfmall.FItemList(i).getLfmallStatName %>
	<%
			End If
			response.write oLfmall.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oLfmall.FItemList(i).FMakerid %> <%= oLfmall.FItemList(i).getDeliverytypeName %><br><%= oLfmall.FItemList(i).FItemName %></td>
	<td align="center"><%= oLfmall.FItemList(i).FRegdate %><br><%= oLfmall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oLfmall.FItemList(i).FLfmallRegdate %><br><%= oLfmall.FItemList(i).FLfmallLastUpdate %></td>
	<td align="right">
		<% If oLfmall.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oLfmall.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oLfmall.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oLfmall.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oLfmall.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (oLfmall.FItemList(i).FLfmallStatCd > 0) and Not IsNULL(oLfmall.FItemList(i).FLfmallPrice) Then
		' 	If (oLfmall.FItemList(i).FSaleYn = "Y") and (CLng((1.0*oLfmall.FItemList(i).FSellcash/10)*10) < oLfmall.FItemList(i).FLfmallPrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-oLfmall.FItemList(i).Fbuycash/oLfmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oLfmall.FItemList(i).Fbuycash/oLfmall.FItemList(i).FLfmallPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oLfmall.FItemList(i).Fbuycash/oLfmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oLfmall.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oLfmall.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oLfmall.FItemList(i).FOrgSuplycash/oLfmall.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oLfmall.FItemList(i).Fbuycash/oLfmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oLfmall.FItemList(i).Fbuycash/oLfmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oLfmall.FItemList(i).IsSoldOut Then
			If oLfmall.FItemList(i).FSellyn = "N" Then
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
		If oLfmall.FItemList(i).FItemdiv = "06" OR oLfmall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oLfmall.FItemList(i).FLfmallStatCd > 0) Then
			If Not IsNULL(oLfmall.FItemList(i).FLfmallPrice) Then
				If (oLfmall.FItemList(i).Mustprice <> oLfmall.FItemList(i).FLfmallPrice) Then
	%>
					<strong><%= formatNumber(oLfmall.FItemList(i).FLfmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oLfmall.FItemList(i).FLfmallPrice,0)&"<br>"
				End If

				If Not IsNULL(oLfmall.FItemList(i).FSpecialPrice) Then
					If (now() >= oLfmall.FItemList(i).FStartDate) And (now() <= oLfmall.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oLfmall.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oLfmall.FItemList(i).FSellyn="Y" and oLfmall.FItemList(i).FLfmallSellYn<>"Y") or (oLfmall.FItemList(i).FSellyn<>"Y" and oLfmall.FItemList(i).FLfmallSellYn="Y") Then
	%>
					<strong><%= oLfmall.FItemList(i).FLfmallSellYn %></strong>
	<%
				Else
					response.write oLfmall.FItemList(i).FLfmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oLfmall.FItemList(i).FLfmallGoodNo <> "" Then %>
			<a target="_blank" href="https://www.lfmall.co.kr/product.do?cmd=getProductDetail&PROD_CD=<%=oLfmall.FItemList(i).FLfmallGoodNo%>"><%=oLfmall.FItemList(i).FLfmallGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oLfmall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oLfmall.FItemList(i).FItemID%>','0');"><%= oLfmall.FItemList(i).FoptionCnt %>:<%= oLfmall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oLfmall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oLfmall.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oLfmall.FItemList(i).FinfoDiv %>
		<%
		If (oLfmall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oLfmall.FItemList(i).FlastErrStr) &"'>ERR:"& oLfmall.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oLfmall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLfmall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLfmall.StartScrollPage to oLfmall.FScrollCount + oLfmall.StartScrollPage - 1 %>
    		<% if i>oLfmall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLfmall.HasNextScroll then %>
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
	<input type="hidden" name="lfmallGoodNo" value= <%= lfmallGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchDiv" value= <%= MatchDiv %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="lfmallYes10x10No" value= <%= lfmallYes10x10No %>>
	<input type="hidden" name="lfmallNo10x10Yes" value= <%= lfmallNo10x10Yes %>>
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
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
