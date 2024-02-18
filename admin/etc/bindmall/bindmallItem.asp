<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/bindmall/bindmallcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, bindmallKeepSell, isSpecialPrice
Dim bestOrdMall, bindmallGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchBrand, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, bindmallYes10x10No, bindmallNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, bindmallGoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim obindmall, xl, kjypageSize, purchasetype
Dim startMargin, endMargin
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
bindmallGoodNo			= request("bindmallGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchBrand				= request("MatchBrand")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
bindmallYes10x10No		= request("bindmallYes10x10No")
bindmallNo10x10Yes		= request("bindmallNo10x10Yes")
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
purchasetype 			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchBrand= ""
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

'bindmall 상품코드 엔터키로 검색되게
If bindmallGoodNo <> "" then
	Dim iA2, arrTemp2, arrbindmallGoodNo
	bindmallGoodNo = replace(bindmallGoodNo,",",chr(10))
	bindmallGoodNo = replace(bindmallGoodNo,chr(13),"")
	arrTemp2 = Split(bindmallGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrbindmallGoodNo = arrbindmallGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	bindmallGoodNo = left(arrbindmallGoodNo,len(arrbindmallGoodNo)-1)
End If

Set obindmall = new CBindmall
	obindmall.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	obindmall.FPageSize					= kjypageSize
Else
	obindmall.FPageSize					= 100
End If
	obindmall.FRectCDL					= request("cdl")
	obindmall.FRectCDM					= request("cdm")
	obindmall.FRectCDS					= request("cds")
	obindmall.FRectItemID				= itemid
	obindmall.FRectItemName				= itemname
	obindmall.FRectSellYn				= sellyn
	obindmall.FRectLimitYn				= limityn
	obindmall.FRectSailYn				= sailyn
	obindmall.FRectStartMargin			= startMargin
	obindmall.FRectEndMargin			= endMargin
	obindmall.FRectMakerid				= makerid
	obindmall.FRectbindmallGoodNo		= bindmallGoodNo
	obindmall.FRectMatchCate			= MatchCate
	obindmall.FRectMatchBrand			= MatchBrand
	obindmall.FRectIsMadeHand			= isMadeHand
	obindmall.FRectIsOption				= isOption
	obindmall.FRectIsReged				= isReged
	obindmall.FRectNotinmakerid			= notinmakerid
	obindmall.FRectNotinitemid			= notinitemid
	obindmall.FRectExcTrans				= exctrans
	obindmall.FRectPriceOption			= priceOption
	obindmall.FRectIsSpecialPrice     	= isSpecialPrice
	obindmall.FRectDeliverytype			= deliverytype
	obindmall.FRectMwdiv				= mwdiv
	obindmall.FRectScheduleNotInItemid	= scheduleNotInItemid
	obindmall.FRectIsextusing			= isextusing
	obindmall.FRectCisextusing			= cisextusing
	obindmall.FRectRctsellcnt			= rctsellcnt
	obindmall.FRectPurchasetype 		= purchasetype

	obindmall.FRectExtNotReg			= ExtNotReg
	obindmall.FRectExpensive10x10		= expensive10x10
	obindmall.FRectdiffPrc				= diffPrc
	obindmall.FRectbindmallYes10x10No	= bindmallYes10x10No
	obindmall.FRectbindmallNo10x10Yes	= bindmallNo10x10Yes
	obindmall.FRectbindmallKeepSell		= bindmallKeepSell
	obindmall.FRectExtSellYn			= extsellyn
	obindmall.FRectInfoDiv				= infoDiv
	obindmall.FRectFailCntOverExcept	= ""
	obindmall.FRectFailCntExists		= failCntExists
	obindmall.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    obindmall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    obindmall.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	obindmall.getbindmallreqExpireItemList
Else
	obindmall.getbindmallRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=bindmallList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=bindmall1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=bindmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=bindmall1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=bindmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="bindmallKeepSell")&&(frm.bindmallKeepSell.checked)){ frm.bindmallKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="bindmallYes10x10No")&&(frm.bindmallYes10x10No.checked)){ frm.bindmallYes10x10No.checked=false }
	if ((comp.name!="bindmallNo10x10Yes")&&(frm.bindmallNo10x10Yes.checked)){ frm.bindmallNo10x10Yes.checked=false }
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

    if ((comp.name=="bindmallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="bindmallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.bindmallYes10x10No.checked){
            comp.form.bindmallYes10x10No.checked = false;
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
	if ((comp.name!="bindmallYes10x10No")&&(frm.bindmallYes10x10No.checked)){ frm.bindmallYes10x10No.checked=false }
	if ((comp.name!="bindmallNo10x10Yes")&&(frm.bindmallNo10x10Yes.checked)){ frm.bindmallNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/bindmall/popbindmallCateList.asp","popCatebindmallmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

//공통코드 검색
function fnbindmallCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;

	if(ccd == ''){
		alert('공통코드를 입력하세요');
		return;
	}
	if (confirm('선택하신 코드를 검색 하시겠습니까?')){
		xLink.location.href = "/admin/etc/bindmall/actbindmallReq.asp?cmdparam=bindmallCommonCode&CommCD="+ccd+"";
	}
}

// 선택된 상품 등록
function bindmallSelectRegProcess() {
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

    if (confirm('bindmall에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

function bindmallSelectRegStepProcess(v) {
	var chkSel=0;
	var strStep;
	switch(v) {
		case 1 :
			strStep="REGSTEP1";
			break;
		case 2 :
			strStep="REGSTEP2";
			break;
	}
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

    if (confirm('bindmall에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = strStep;
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function bindmallSellYnProcess(chkYn) {
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
		case "X": strSell="승인전 상품삭제";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
       	document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function bindmallEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function bindmallSelectPriceEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 가격 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

//컨텐츠 수정
function bindmallSelectContentEditProcess(){
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

    if (confirm('선택하신 ' + chkSel + '개 상품 컨텐츠를 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

//이미지 수정
function bindmallSelectImageEditProcess(){
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

    if (confirm('선택하신 ' + chkSel + '개 상품 이미지를 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

//출고, 교환, 반품지 수정
function bindmallSelectDeliveryEditProcess(){
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

    if (confirm('선택하신 ' + chkSel + '개 상품 출고/교환/반품지를 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "DELIVERY";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

//옵션 수정
function bindmallSelectOptionEditProcess(){
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

    if (confirm('선택하신 ' + chkSel + '개 상품 옵션을 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}

//옵션 수정
function bindmallSelectOptionAddProcess(){
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

    if (confirm('선택하신 ' + chkSel + '개 상품 옵션을 수정 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTADD";
        document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
        document.frmSvArr.submit();
    }
}


//상품 + 옵션 조회
function bindmallChkStatProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
		document.frmSvArr.submit();
    }
}

//상품 조회
function bindmallChkItemProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 조회 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKITEM";
		document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
		document.frmSvArr.submit();
    }
}

//옵션 조회
function bindmallChkOptProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 옵션을 조회 하시겠습니까?\n\n※bindmall와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKOPT";
		document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
		document.frmSvArr.submit();
    }
}


//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=bindmall1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function bindmallEditProcessAuto() {
	var cnt = <%= obindmall.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "/admin/etc/bindmall/actbindmallReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= obindmall.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		bindmallEditProcessAuto();
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
		<a href="https://mall.bindseller.co.kr/info/index.php" target="_blank">bindmallAdmin바로가기</a>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		bindmall 상품코드 : <textarea rows="2" cols="20" name="bindmallGoodNo" id="itemid"><%=replace(bindmallGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >bindmall 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >bindmall 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >bindmall 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>bindmall 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="bindmallYes10x10No" <%= ChkIIF(bindmallYes10x10No="on","checked","") %> ><font color=red>bindmall판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="bindmallNo10x10Yes" <%= ChkIIF(bindmallNo10x10Yes="on","checked","") %> ><font color=red>bindmall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="bindmallKeepSell" <%= ChkIIF(bindmallKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('bindmall1010');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="카테고리" onclick="pop_CateManager();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<!-- 오픈시 이 조건 제거하기 -->
<% If (session("ssBctID")="kjy8517") Then %>
<tr bgcolor="#FFFFFF">
    <td style="padding:5 0 5 0">
	    <table width="100%" class="a">
	    <tr>
	    	<td valign="top">
				실제상품 등록 :
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="bindmallSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="등록(STEP1)" onClick="bindmallSelectRegStepProcess(1);">&nbsp;&nbsp;
				<input class="button" type="button" value="등록(STEP2)" onClick="bindmallSelectRegStepProcess(2);">
				<% End If %>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" value="수정" onClick="bindmallEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="가격" onClick="bindmallSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="컨텐츠" onClick="bindmallSelectContentEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="이미지" onClick="bindmallSelectImageEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="출고지" onClick="bindmallSelectDeliveryEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="bindmallSelectOptionEditProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" value="옵션추가" onClick="bindmallSelectOptionAddProcess();">&nbsp;&nbsp;
			<% End If %>
				<br><br>
				실제상품 조회 :
				<input class="button" type="button" value="조회" onClick="bindmallChkStatProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="상품" onClick="bindmallChkItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="옵션" onClick="bindmallChkOptProcess();">
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				공통코드 검색 :
				<input class="text" type="text" name="CommCD" id="CommCD" value="">&nbsp;&nbsp;
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="fnbindmallCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
					<option value="X">승인전 상품삭제</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="bindmallSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
<% End If %>
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
		검색결과 : <b><%= FormatNumber(obindmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(obindmall.FTotalPage,0) %></b>
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
	<td width="140">bindmall등록일<br>bindmall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">bindmall<br>가격및판매</td>
	<td width="70">bindmall<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to obindmall.FResultCount - 1 %>
<tr align="center" <% If obindmall.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= obindmall.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= obindmall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= obindmall.FItemList(i).FItemID %>','bindmall1010','<%=obindmall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=obindmall.FItemList(i).FItemID%>" target="_blank"><%= obindmall.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If obindmall.FItemList(i).FbindmallStatCd <> 7 Then
	%>
		<br><%= obindmall.FItemList(i).getbindmallStatName %>
	<%
			End If
			response.write obindmall.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= obindmall.FItemList(i).FMakerid %> <%= obindmall.FItemList(i).getDeliverytypeName %><br><%= obindmall.FItemList(i).FItemName %></td>
	<td align="center"><%= obindmall.FItemList(i).FRegdate %><br><%= obindmall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= obindmall.FItemList(i).FbindmallRegdate %><br><%= obindmall.FItemList(i).FbindmallLastUpdate %></td>
	<td align="right">
		<% If obindmall.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(obindmall.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(obindmall.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(obindmall.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If obindmall.FItemList(i).Fsellcash = 0 Then
		elseif (obindmall.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (obindmall.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-obindmall.FItemList(i).FOrgSuplycash/obindmall.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-obindmall.FItemList(i).Fbuycash/obindmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-obindmall.FItemList(i).Fbuycash/obindmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If obindmall.FItemList(i).IsSoldOut Then
			If obindmall.FItemList(i).FSellyn = "N" Then
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
		If obindmall.FItemList(i).FItemdiv = "06" OR obindmall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (obindmall.FItemList(i).FbindmallStatCd > 0) Then
			If Not IsNULL(obindmall.FItemList(i).FbindmallPrice) Then
				If (obindmall.FItemList(i).Mustprice <> obindmall.FItemList(i).FbindmallPrice) Then
	%>
					<strong><%= formatNumber(obindmall.FItemList(i).FbindmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(obindmall.FItemList(i).FbindmallPrice,0)&"<br>"
				End If

				If Not IsNULL(obindmall.FItemList(i).FSpecialPrice) Then
					If (now() >= obindmall.FItemList(i).FStartDate) And (now() <= obindmall.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(obindmall.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (obindmall.FItemList(i).FSellyn="Y" and obindmall.FItemList(i).FbindmallSellYn<>"Y") or (obindmall.FItemList(i).FSellyn<>"Y" and obindmall.FItemList(i).FbindmallSellYn="Y") Then
	%>
					<strong><%= obindmall.FItemList(i).FbindmallSellYn %></strong>
	<%
				Else
					response.write obindmall.FItemList(i).FbindmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If obindmall.FItemList(i).FbindmallGoodNo <> "" Then %>
			<a target="_blank" href="https://www.bindmall.co.kr/product/<%=obindmall.FItemList(i).FbindmallGoodNo%>"><%=obindmall.FItemList(i).FbindmallGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= obindmall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=obindmall.FItemList(i).FItemID%>','0');"><%= obindmall.FItemList(i).FoptionCnt %>:<%= obindmall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= obindmall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= Chkiif(obindmall.FItemList(i).FCateMapCnt > 0, "매칭됨(카)", "<font color='darkred'>매칭안됨(카)</font>") %>
	</td>
	<td align="center">
		<%= obindmall.FItemList(i).FinfoDiv %>
		<%
		If (obindmall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(obindmall.FItemList(i).FlastErrStr) &"'>ERR:"& obindmall.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% bindmallGoodNoArray = bindmallGoodNoArray & obindmall.FItemList(i).FbindmallGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= bindmallGoodNoArray %></textarea>
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
        <% if obindmall.HasPreScroll then %>
		<a href="javascript:goPage('<%= obindmall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + obindmall.StartScrollPage to obindmall.FScrollCount + obindmall.StartScrollPage - 1 %>
    		<% if i>obindmall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if obindmall.HasNextScroll then %>
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
	<input type="hidden" name="bindmallGoodNo" value= <%= bindmallGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchBrand" value= <%= MatchBrand %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="bindmallYes10x10No" value= <%= bindmallYes10x10No %>>
	<input type="hidden" name="bindmallNo10x10Yes" value= <%= bindmallNo10x10Yes %>>
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
<% SET obindmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
