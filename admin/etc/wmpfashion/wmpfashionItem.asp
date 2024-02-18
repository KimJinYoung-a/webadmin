<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wmpfashion/wmpfashionCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, wfwemakeKeepSell, isSpecialPrice
Dim bestOrdMall, wfwemakeGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, wfwemakeYes10x10No, wfwemakeNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, scheduleNotInItemid
Dim page, i, research, isextusing
Dim oWmpfashion
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
wfwemakeGoodNo			= request("wfwemakeGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
wfwemakeYes10x10No		= request("wfwemakeYes10x10No")
wfwemakeNo10x10Yes		= request("wfwemakeNo10x10Yes")
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

'위메프 상품코드 엔터키로 검색되게
If wfwemakeGoodNo <> "" then
	Dim iA2, arrTemp2, arrwfwemakeGoodNo
	wfwemakeGoodNo = replace(wfwemakeGoodNo,",",chr(10))
	wfwemakeGoodNo = replace(wfwemakeGoodNo,chr(13),"")
	arrTemp2 = Split(wfwemakeGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrwfwemakeGoodNo = arrwfwemakeGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	wfwemakeGoodNo = left(arrwfwemakeGoodNo,len(arrwfwemakeGoodNo)-1)
End If

Set oWmpfashion = new CWmpfashion
	oWmpfashion.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oWmpfashion.FPageSize					= 100
Else
	oWmpfashion.FPageSize					= 50
End If
	oWmpfashion.FRectCDL					= request("cdl")
	oWmpfashion.FRectCDM					= request("cdm")
	oWmpfashion.FRectCDS					= request("cds")
	oWmpfashion.FRectItemID					= itemid
	oWmpfashion.FRectItemName				= itemname
	oWmpfashion.FRectSellYn					= sellyn
	oWmpfashion.FRectLimitYn				= limityn
	oWmpfashion.FRectSailYn					= sailyn
	oWmpfashion.FRectStartMargin			= startMargin
	oWmpfashion.FRectEndMargin				= endMargin
	oWmpfashion.FRectMakerid				= makerid
	oWmpfashion.FRectWfWemakeGoodNo			= wfwemakeGoodNo
	oWmpfashion.FRectMatchCate				= MatchCate
	oWmpfashion.FRectIsMadeHand				= isMadeHand
	oWmpfashion.FRectIsOption				= isOption
	oWmpfashion.FRectIsReged				= isReged
	oWmpfashion.FRectNotinmakerid			= notinmakerid
	oWmpfashion.FRectNotinitemid			= notinitemid
	oWmpfashion.FRectScheduleNotInItemid	= scheduleNotInItemid
	oWmpfashion.FRectIsextusing				= isextusing

	oWmpfashion.FRectExcTrans				= exctrans
	oWmpfashion.FRectPriceOption			= priceOption
	oWmpfashion.FRectIsSpecialPrice     	= isSpecialPrice
	oWmpfashion.FRectDeliverytype			= deliverytype
	oWmpfashion.FRectMwdiv					= mwdiv

	oWmpfashion.FRectExtNotReg				= ExtNotReg
	oWmpfashion.FRectExpensive10x10			= expensive10x10
	oWmpfashion.FRectdiffPrc				= diffPrc
	oWmpfashion.FRectWfWemakeYes10x10No		= wfwemakeYes10x10No
	oWmpfashion.FRectWfWemakeNo10x10Yes		= wfwemakeNo10x10Yes
	oWmpfashion.FRectWfWemakeKeepSell		= wfwemakeKeepSell
	oWmpfashion.FRectExtSellYn				= extsellyn
	oWmpfashion.FRectInfoDiv				= infoDiv
	oWmpfashion.FRectFailCntOverExcept		= ""
	oWmpfashion.FRectFailCntExists			= failCntExists
	oWmpfashion.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oWmpfashion.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oWmpfashion.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oWmpfashion.getWmpreqExpireItemList
Else
	oWmpfashion.getWmpfashionRegedItemList		'그 외 리스트
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=wmpfashion","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=wmpfashion','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=wmpfashion','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=wmpfashion','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록 가능 상품
function RegAvailItem(){
	var popwin=window.open('/admin/etc/reg_avail_Itemid.asp?mallgubun=wmpfashion','regAvailItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 딜 관리
function DealItem(){
	var popwin=window.open('/admin/etc/wmpfashion/popDealItemList.asp','deal','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//키워드 관리
function popKeywordItemList(){
	var popwin=window.open('/admin/etc/common/popKeywordList.asp?mallgubun=wmpfashion','popKeywordItemList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function deleteItem(){
	if(confirm("먼저 상품이 판매중지로 되어있는 지 확인 해주세요\n\n삭제하시겠습니까?")){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.mode.value = "deal";
		document.frmSvArr.auto.value = "Y";
		document.frmSvArr.action = "procDeal.asp"
		document.frmSvArr.submit();
	}
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
	if ((comp.name!="wfwemakeKeepSell")&&(frm.wfwemakeKeepSell.checked)){ frm.wfwemakeKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="wfwemakeYes10x10No")&&(frm.wfwemakeYes10x10No.checked)){ frm.wfwemakeYes10x10No.checked=false }
	if ((comp.name!="wfwemakeNo10x10Yes")&&(frm.wfwemakeNo10x10Yes.checked)){ frm.wfwemakeNo10x10Yes.checked=false }
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

    if ((comp.name=="wfwemakeYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="wfwemakeNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.wfwemakeYes10x10No.checked){
            comp.form.wfwemakeYes10x10No.checked = false;
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
	if ((comp.name!="wfwemakeYes10x10No")&&(frm.wfwemakeYes10x10No.checked)){ frm.wfwemakeYes10x10No.checked=false }
	if ((comp.name!="wfwemakeNo10x10Yes")&&(frm.wfwemakeNo10x10Yes.checked)){ frm.wfwemakeNo10x10Yes.checked=false }
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
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/wmpfashion/popWmpfashioncateList.asp","popCateWmpfashionmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// 선택된 상품 일괄 등록
function wemakeSelectRegProcess() {
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

    if (confirm('위메프에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※위메프와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 상태 변경
function wemakeSellYnProcess(chkYn) {
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
		case "X": strSell="삭제";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 DB에서 삭제됩니다.\n\n반드시 위메프 어드민에서 판매종료 시켜야합니다.\n\n계속 하시겠습니까?')) return;
        }
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 가격 수정
function wemakePriceEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※위메프와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		//document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 재고 조회
function wemakecheckStatProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 재고를 조회 하시겠습니까?\n\n※위메프와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 일괄 수정
function wemakeEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※위메프와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
		document.frmSvArr.submit();
    }
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=wmpfashion&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function wemakeEditProcessAuto() {
	var cnt = <%= oWmpfashion.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/wmpfashion/actWmpfashionReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oWmpfashion.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		wemakeEditProcessAuto();
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
		<a href="https://wpartner.wemakeprice.com/login" target="_blank">위메프Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ wf00010x10 | store10x10 ]</font>"
			End If

			If (session("ssBctID")="kjy8517") then
				response.write "&nbsp;&nbsp;VPN PW : kjy8517 | cube101010"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		위메프 상품코드 : <textarea rows="2" cols="20" name="wfwemakeGoodNo" id="itemid"><%=replace(wfwemakeGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >위메프 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >위메프 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >위메프 등록완료(전시)
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
		스케줄제외상품
		<select name="scheduleNotInItemid" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(scheduleNotInItemid="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(scheduleNotInItemid="N","selected","") %> >N
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>위메프 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wfwemakeYes10x10No" <%= ChkIIF(wfwemakeYes10x10No="on","checked","") %> ><font color=red>위메프판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="wfwemakeNo10x10Yes" <%= ChkIIF(wfwemakeNo10x10Yes="on","checked","") %> ><font color=red>위메프품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wfwemakeKeepSell" <%= ChkIIF(wfwemakeKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
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
<form name="frmReg" method="post" action="wmpfashionItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="등록 가능 상품" onclick="RegAvailItem();">&nbsp;
			<!-- 2019-05-21 김진영 주석처리..사용안함
				<input class="button" type="button" value="딜 상품 삭제" onclick="deleteItem();">
				<input class="button" type="button" value="전시 제외 상품" onclick="DisplayNotInItemid();">&nbsp;

				<input class="button" type="button" value="딜 관리" onclick="DealItem();">&nbsp;
				<input class="button" type="button" value="키워드" onclick="popKeywordItemList();">
			-->
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('wmpfashion');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="wemakeSelectRegProcess();">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="wemakePriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEdit" value="수정" onClick="wemakeEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="조회" onClick="wemakecheckStatProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="yhj0613") Then %>
					<option value="X">삭제(관리자용)</option>
				<% End If %>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="wemakeSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oWmpfashion.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oWmpfashion.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">위메프등록일<br>위메프최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">위메프<br>가격및판매</td>
	<td width="70">위메프<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oWmpfashion.FResultCount - 1 %>
<tr align="center" <% If oWmpfashion.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %> >
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oWmpfashion.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oWmpfashion.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oWmpfashion.FItemList(i).FItemID %>','wmpfashion','<%=oWmpfashion.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oWmpfashion.FItemList(i).FItemID%>" target="_blank"><%= oWmpfashion.FItemList(i).FItemID %></a>
		<% If oWmpfashion.FItemList(i).FWfWemakeStatcd <> 7 Then %>
		<br><%= oWmpfashion.FItemList(i).getWfwemakeStatName %>
		<% End If %>
		<%= oWmpfashion.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oWmpfashion.FItemList(i).FMakerid %> <%= oWmpfashion.FItemList(i).getDeliverytypeName %><br><%= oWmpfashion.FItemList(i).FItemName %></td>
	<td align="center"><%= oWmpfashion.FItemList(i).FRegdate %><br><%= oWmpfashion.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oWmpfashion.FItemList(i).FWfWemakeRegdate %><br><%= oWmpfashion.FItemList(i).FWfWemakeLastUpdate %></td>
	<td align="right">
		<% If oWmpfashion.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oWmpfashion.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oWmpfashion.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oWmpfashion.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oWmpfashion.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (oWmpfashion.FItemList(i).FWemakeStatCd > 0) and Not IsNULL(oWmpfashion.FItemList(i).FWfWemakePrice) Then
		' 	If (oWmpfashion.FItemList(i).FSaleYn = "Y") and (CLng((1.0*oWmpfashion.FItemList(i).FSellcash/10)*10) < oWmpfashion.FItemList(i).FWfWemakePrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-oWmpfashion.FItemList(i).Fbuycash/oWmpfashion.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oWmpfashion.FItemList(i).Fbuycash/oWmpfashion.FItemList(i).FWfWemakePrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oWmpfashion.FItemList(i).Fbuycash/oWmpfashion.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oWmpfashion.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oWmpfashion.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oWmpfashion.FItemList(i).FOrgSuplycash/oWmpfashion.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oWmpfashion.FItemList(i).Fbuycash/oWmpfashion.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oWmpfashion.FItemList(i).Fbuycash/oWmpfashion.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oWmpfashion.FItemList(i).IsSoldOut Then
			If oWmpfashion.FItemList(i).FSellyn = "N" Then
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
		If oWmpfashion.FItemList(i).FItemdiv = "06" OR oWmpfashion.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oWmpfashion.FItemList(i).FWfWemakeStatCd > 0) Then
			If Not IsNULL(oWmpfashion.FItemList(i).FWfWemakePrice) Then
				If (oWmpfashion.FItemList(i).Mustprice <> oWmpfashion.FItemList(i).FWfWemakePrice) Then
	%>
					<strong><%= formatNumber(oWmpfashion.FItemList(i).FWfWemakePrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oWmpfashion.FItemList(i).FWfWemakePrice,0)&"<br>"
				End If

				If Not IsNULL(oWmpfashion.FItemList(i).FSpecialPrice) Then
					If (now() >= oWmpfashion.FItemList(i).FStartDate) And (now() <= oWmpfashion.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oWmpfashion.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oWmpfashion.FItemList(i).FSellyn="Y" and oWmpfashion.FItemList(i).FWfWemakeSellYn<>"Y") or (oWmpfashion.FItemList(i).FSellyn<>"Y" and oWmpfashion.FItemList(i).FWfWemakeSellYn="Y") Then
	%>
					<strong><%= oWmpfashion.FItemList(i).FWfWemakeSellYn %></strong>
	<%
				Else
					response.write oWmpfashion.FItemList(i).FWfWemakeSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oWmpfashion.FItemList(i).FWfWemakeGoodNo <> "" Then %>
			<a target="_blank" href="https://front.wemakeprice.com/product/<%=oWmpfashion.FItemList(i).FWfWemakeGoodNo%>"><%=oWmpfashion.FItemList(i).FWfWemakeGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oWmpfashion.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oWmpfashion.FItemList(i).FItemID%>','0');"><%= oWmpfashion.FItemList(i).FoptionCnt %>:<%= oWmpfashion.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oWmpfashion.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oWmpfashion.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oWmpfashion.FItemList(i).FinfoDiv %>
		<%
		If (oWmpfashion.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oWmpfashion.FItemList(i).FlastErrStr) &"'>ERR:"& oWmpfashion.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oWmpfashion.HasPreScroll then %>
		<a href="javascript:goPage('<%= oWmpfashion.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oWmpfashion.StartScrollPage to oWmpfashion.FScrollCount + oWmpfashion.StartScrollPage - 1 %>
    		<% if i>oWmpfashion.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oWmpfashion.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>

<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
