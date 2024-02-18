<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/benepia/benepiacls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, benepiaKeepSell, isSpecialPrice
Dim bestOrdMall, benepiaGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, benepiaYes10x10No, benepiaNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, benepiaGoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim obenepia, xl, kjypageSize
Dim startMargin, endMargin
Dim purchasetype
page    				= request("page")
kjypageSize				= request("kjypageSize")
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
benepiaGoodNo			= request("benepiaGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
benepiaYes10x10No		= request("benepiaYes10x10No")
benepiaNo10x10Yes		= request("benepiaNo10x10Yes")
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
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
purchasetype			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
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

'benepia 상품코드 엔터키로 검색되게
If benepiaGoodNo <> "" then
	Dim iA2, arrTemp2, arrbenepiaGoodNo
	benepiaGoodNo = replace(benepiaGoodNo,",",chr(10))
	benepiaGoodNo = replace(benepiaGoodNo,chr(13),"")
	arrTemp2 = Split(benepiaGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrbenepiaGoodNo = arrbenepiaGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	benepiaGoodNo = left(arrbenepiaGoodNo,len(arrbenepiaGoodNo)-1)
End If

Set obenepia = new CBenepia
	obenepia.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	obenepia.FPageSize					= kjypageSize
Else
	obenepia.FPageSize					= 100
End If
	obenepia.FRectCDL					= request("cdl")
	obenepia.FRectCDM					= request("cdm")
	obenepia.FRectCDS					= request("cds")
	obenepia.FRectItemID				= itemid
	obenepia.FRectItemName				= itemname
	obenepia.FRectSellYn				= sellyn
	obenepia.FRectLimitYn				= limityn
	obenepia.FRectSailYn				= sailyn
	obenepia.FRectStartMargin			= startMargin
	obenepia.FRectEndMargin				= endMargin
	obenepia.FRectMakerid				= makerid
	obenepia.FRectbenepiaGoodNo			= benepiaGoodNo
	obenepia.FRectMatchCate				= MatchCate
	obenepia.FRectIsMadeHand			= isMadeHand
	obenepia.FRectIsOption				= isOption
	obenepia.FRectIsReged				= isReged
	obenepia.FRectNotinmakerid			= notinmakerid
	obenepia.FRectNotinitemid			= notinitemid
	obenepia.FRectExcTrans				= exctrans
	obenepia.FRectPriceOption			= priceOption
	obenepia.FRectIsSpecialPrice     	= isSpecialPrice
	obenepia.FRectDeliverytype			= deliverytype
	obenepia.FRectMwdiv					= mwdiv
	obenepia.FRectScheduleNotInItemid	= scheduleNotInItemid
	obenepia.FRectIsextusing			= isextusing
	obenepia.FRectCisextusing			= cisextusing
	obenepia.FRectRctsellcnt			= rctsellcnt

	obenepia.FRectExtNotReg				= ExtNotReg
	obenepia.FRectExpensive10x10		= expensive10x10
	obenepia.FRectdiffPrc				= diffPrc
	obenepia.FRectbenepiaYes10x10No	= benepiaYes10x10No
	obenepia.FRectbenepiaNo10x10Yes	= benepiaNo10x10Yes
	obenepia.FRectbenepiaKeepSell		= benepiaKeepSell
	obenepia.FRectExtSellYn				= extsellyn
	obenepia.FRectInfoDiv				= infoDiv
	obenepia.FRectFailCntOverExcept		= ""
	obenepia.FRectFailCntExists			= failCntExists
	obenepia.FRectReqEdit				= reqEdit
	obenepia.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    obenepia.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    obenepia.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	obenepia.getbenepiareqExpireItemList
Else
	obenepia.getbenepiaRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=benepiaList"& replace(DATE(), "-", "") &"_xl.xls"
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

function jsByValue(s){
	if((s == "category")){
		$("#goodsGrpCd_span").show();
	}else{
		$("#goodsGrpCd_span").hide();
	}
}


// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=benepia1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=benepia1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=benepia1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=benepia1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="benepiaKeepSell")&&(frm.benepiaKeepSell.checked)){ frm.benepiaKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="benepiaYes10x10No")&&(frm.benepiaYes10x10No.checked)){ frm.benepiaYes10x10No.checked=false }
	if ((comp.name!="benepiaNo10x10Yes")&&(frm.benepiaNo10x10Yes.checked)){ frm.benepiaNo10x10Yes.checked=false }
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

    if ((comp.name=="benepiaYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="benepiaNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.benepiaYes10x10No.checked){
            comp.form.benepiaYes10x10No.checked = false;
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
	if ((comp.name!="benepiaYes10x10No")&&(frm.benepiaYes10x10No.checked)){ frm.benepiaYes10x10No.checked=false }
	if ((comp.name!="benepiaNo10x10Yes")&&(frm.benepiaNo10x10Yes.checked)){ frm.benepiaNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/benepia/popbenepiaCateList.asp","popCatebenepiamanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//공통코드 검색
function fnbenepiaCommCDSubmit() {
	var ccd;
	var goodsGrpCd;
	ccd = document.getElementById('CommCD').value;
	goodsGrpCd = $("#goodsGrpCd option:selected").val();

	if(ccd == ''){
		alert('공통코드를 선택하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		
		xLink.location.href = "/admin/etc/benepia/actbenepiaReq.asp?cmdparam=benepiaCommonCode&CommCD="+ccd+"&goodsGrpCd="+goodsGrpCd;
	}
}

// 선택된 상품 이미지 등록
function benepiaSelectImageRegProcess() {
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

    if (confirm('benepia에 선택하신 ' + chkSel + '개 이미지를 등록 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 등록
function benepiaSelectRegProcess() {
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

    if (confirm('benepia에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품만 등록
function benepiaSelectItemRegProcess() {
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

    if (confirm('benepia에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGITEM";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function benepiaSellYnProcess(chkYn) {
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
       	document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function benepiaEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 카테고리 수정
function benepiaCategoryEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITCATE";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function benepiaPriceEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function benepiaQuantityEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "QTY";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 정보 수정
function benepiaItemInfoEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITINFO";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 배송정보 수정
function benepiaItemDeliveryEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITDELIVERY";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 수정
function benepiaImageEditProcess(v) {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
		if (v == 1) {
        	document.frmSvArr.cmdparam.value = "EDITIMAGEPC";
		} else {
			document.frmSvArr.cmdparam.value = "EDITIMAGEMOB";
		}
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 설명 수정
function benepiaContentEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 안전인증정보 수정
function benepiaSafeInfoEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SAFEINFO";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 정보고시 수정
function benepiaInfoCodeEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "INFOCODE";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 옵션 수정
function benepiaOptionEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 + 옵션 조회
function benepiaChkStatProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
		document.frmSvArr.submit();
    }
}

//상품 조회
function benepiaChkItemProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKITEM";
		document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
		document.frmSvArr.submit();
    }
}

//옵션 조회
function benepiaChkOptProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※benepia와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKOPT";
		document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
		document.frmSvArr.submit();
    }
}

//상품 삭제
function benepiaDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\nbenepia 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=benepia1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function benepiaEditProcessAuto() {
	var cnt = <%= obenepia.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "/admin/etc/benepia/actbenepiaReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= obenepia.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		benepiaEditProcessAuto();
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
		<a href="https://newmallvenadm.benepia.co.kr/login/loginView.do" target="_blank">benepiaAdmin바로가기</a>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		benepia 상품코드 : <textarea rows="2" cols="20" name="benepiaGoodNo" id="itemid"><%=replace(benepiaGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >benepia 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >benepia 등록예정이상
			<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >benepia 승인예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >benepia 등록완료(전시)
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
		<% If (session("ssBctID")="kjy8517") Then %>
			<input class="text" size="5" type="text" name="kjypageSize" value="<%= kjypageSize %>">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>benepia 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="benepiaYes10x10No" <%= ChkIIF(benepiaYes10x10No="on","checked","") %> ><font color=red>benepia판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="benepiaNo10x10Yes" <%= ChkIIF(benepiaNo10x10Yes="on","checked","") %> ><font color=red>benepia품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="benepiaKeepSell" <%= ChkIIF(benepiaKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 배송방법이 택배(일반) 아닌것, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 옵션가 0원 판매중 상품이 없음(옵션 한정수량 5개 이하 포함), 옵션가가 판매가 100% 이상인 상품

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
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('benepia1010');">&nbsp;&nbsp;
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
				<input class="button" type="button" value="등록" onClick="benepiaSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="상품" onClick="benepiaSelectItemRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지" onClick="benepiaSelectImageRegProcess();">&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" value="수정" onClick="benepiaEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="가격" onClick="benepiaPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="benepiaOptionEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="재고" onClick="benepiaQuantityEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="정보" onClick="benepiaItemInfoEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="배송" onClick="benepiaItemDeliveryEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지(PC)" onClick="benepiaImageEditProcess(1);">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지(MOB)" onClick="benepiaImageEditProcess(2);">&nbsp;&nbsp;
				<input class="button" type="button" value="설명" onClick="benepiaContentEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="인증" onClick="benepiaSafeInfoEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="고시" onClick="benepiaInfoCodeEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="카테고리" onClick="benepiaCategoryEditProcess();">&nbsp;&nbsp;
				<br><br>
				실제상품 조회 :
				<input class="button" type="button" value="조회" onClick="benepiaChkStatProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="상품" onClick="benepiaChkItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="benepiaChkOptProcess();">
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="nys1006") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="상품삭제" onClick="benepiaDeleteProcess();" style=font-weight:bold>
			<% End If %>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<br><br>
				공통코드 검색 :
				<select name="CommCD" class="select" id="CommCD" onChange="jsByValue(this.value);">
					<option value="">- Choice -</option>
					<option value="category">카테고리</option>
					<option value="md">MD</option>
					<option value="area">배송가능지역</option>
					<option value="safe">인증유형</option>
					<option value="infocode">정보고시</option>
					<option value="infocodedtl">정보고시항목</option>
					<option value="casedelivery">벤더사 조건부 배송조회</option>
					<option value="parcel">택배사</option>
					<option value="brand">브랜드</option>
					<option value="locaddress">출고/반품지조회</option>
				</select>
				<span id="goodsGrpCd_span" style="display:none;">
					<select class="select" name="goodsGrpCd" id="goodsGrpCd">
						<option value="1">1Depth</option>
						<option value="2">2Depth</option>
						<option value="3">3Depth</option>
						<option value="4">4Depth</option>
						<option value="e">카테고리 재귀호출</option>
					</select>
				</span>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="fnbenepiaCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="benepiaSellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(obenepia.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(obenepia.FTotalPage,0) %></b>
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
	<td width="140">benepia등록일<br>benepia최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">benepia<br>가격및판매</td>
	<td width="70">benepia<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to obenepia.FResultCount - 1 %>
<tr align="center" <% If obenepia.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= obenepia.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= obenepia.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= obenepia.FItemList(i).FItemID %>','benepia','<%=obenepia.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=obenepia.FItemList(i).FItemID%>" target="_blank"><%= obenepia.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If obenepia.FItemList(i).FbenepiaStatCd <> 7 Then
	%>
		<br><%= obenepia.FItemList(i).getbenepiaStatName %>
	<%
			End If
			response.write obenepia.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= obenepia.FItemList(i).FMakerid %> <%= obenepia.FItemList(i).getDeliverytypeName %><br><%= obenepia.FItemList(i).FItemName %></td>
	<td align="center"><%= obenepia.FItemList(i).FRegdate %><br><%= obenepia.FItemList(i).FLastupdate %></td>
	<td align="center"><%= obenepia.FItemList(i).FbenepiaRegdate %><br><%= obenepia.FItemList(i).FbenepiaLastUpdate %></td>
	<td align="right">
		<% If obenepia.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(obenepia.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(obenepia.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(obenepia.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If obenepia.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (obenepia.FItemList(i).FbenepiaStatCd > 0) and Not IsNULL(obenepia.FItemList(i).FbenepiaPrice) Then
		' 	If (obenepia.FItemList(i).FSaleYn = "Y") and (CLng((1.0*obenepia.FItemList(i).FSellcash/10)*10) < obenepia.FItemList(i).FbenepiaPrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-obenepia.FItemList(i).Fbuycash/obenepia.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-obenepia.FItemList(i).Fbuycash/obenepia.FItemList(i).FbenepiaPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-obenepia.FItemList(i).Fbuycash/obenepia.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (obenepia.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (obenepia.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-obenepia.FItemList(i).FOrgSuplycash/obenepia.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-obenepia.FItemList(i).Fbuycash/obenepia.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-obenepia.FItemList(i).Fbuycash/obenepia.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If obenepia.FItemList(i).IsSoldOut Then
			If obenepia.FItemList(i).FSellyn = "N" Then
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
		If obenepia.FItemList(i).FItemdiv = "06" OR obenepia.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (obenepia.FItemList(i).FbenepiaStatCd > 0) Then
			If Not IsNULL(obenepia.FItemList(i).FbenepiaPrice) Then
				If (obenepia.FItemList(i).Mustprice <> obenepia.FItemList(i).FbenepiaPrice) Then
	%>
					<strong><%= formatNumber(obenepia.FItemList(i).FbenepiaPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(obenepia.FItemList(i).FbenepiaPrice,0)&"<br>"
				End If

				If Not IsNULL(obenepia.FItemList(i).FSpecialPrice) Then
					If (now() >= obenepia.FItemList(i).FStartDate) And (now() <= obenepia.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(obenepia.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (obenepia.FItemList(i).FSellyn="Y" and obenepia.FItemList(i).FbenepiaSellYn<>"Y") or (obenepia.FItemList(i).FSellyn<>"Y" and obenepia.FItemList(i).FbenepiaSellYn="Y") Then
	%>
					<strong><%= obenepia.FItemList(i).FbenepiaSellYn %></strong>
	<%
				Else
					response.write obenepia.FItemList(i).FbenepiaSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If obenepia.FItemList(i).FbenepiaGoodNo <> "" Then %>
			<a target="_blank" href="https://newmall.benepia.co.kr/disp/storeMain.bene?prdId=<%=obenepia.FItemList(i).FbenepiaGoodNo%>"><%=obenepia.FItemList(i).FbenepiaGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= obenepia.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=obenepia.FItemList(i).FItemID%>','0');"><%= obenepia.FItemList(i).FoptionCnt %>:<%= obenepia.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= obenepia.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If obenepia.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= obenepia.FItemList(i).FinfoDiv %>
		<%
		If (obenepia.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(obenepia.FItemList(i).FlastErrStr) &"'>ERR:"& obenepia.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% benepiaGoodNoArray = benepiaGoodNoArray & obenepia.FItemList(i).FbenepiaGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= benepiaGoodNoArray %></textarea>
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
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if obenepia.HasPreScroll then %>
		<a href="javascript:goPage('<%= obenepia.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + obenepia.StartScrollPage to obenepia.FScrollCount + obenepia.StartScrollPage - 1 %>
    		<% if i>obenepia.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if obenepia.HasNextScroll then %>
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
	<input type="hidden" name="benepiaGoodNo" value= <%= benepiaGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="benepiaYes10x10No" value= <%= benepiaYes10x10No %>>
	<input type="hidden" name="benepiaNo10x10Yes" value= <%= benepiaNo10x10Yes %>>
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
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET obenepia = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
