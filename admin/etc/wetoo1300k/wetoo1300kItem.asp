<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wetoo1300k/wetoo1300kcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, wetoo1300kKeepSell, isSpecialPrice
Dim bestOrdMall, wetoo1300kGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, wetoo1300kYes10x10No, wetoo1300kNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, wetoo1300kGoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt, MatchBrand
Dim o1300k, xl, kjypageSize
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
wetoo1300kGoodNo		= request("wetoo1300kGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
wetoo1300kYes10x10No	= request("wetoo1300kYes10x10No")
wetoo1300kNo10x10Yes	= request("wetoo1300kNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
MatchBrand				= request("MatchBrand")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
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

'1300k 상품코드 엔터키로 검색되게
If wetoo1300kGoodNo <> "" then
	Dim iA2, arrTemp2, arrwetoo1300kGoodNo
	wetoo1300kGoodNo = replace(wetoo1300kGoodNo,",",chr(10))
	wetoo1300kGoodNo = replace(wetoo1300kGoodNo,chr(13),"")
	arrTemp2 = Split(wetoo1300kGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrwetoo1300kGoodNo = arrwetoo1300kGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	wetoo1300kGoodNo = left(arrwetoo1300kGoodNo,len(arrwetoo1300kGoodNo)-1)
End If

Set o1300k = new C1300k
	o1300k.FCurrPage				= page
If (session("ssBctID")="kjy8517") Then
	o1300k.FPageSize				= kjypageSize
Else
	o1300k.FPageSize				= 100
End If
	o1300k.FRectCDL					= request("cdl")
	o1300k.FRectCDM					= request("cdm")
	o1300k.FRectCDS					= request("cds")
	o1300k.FRectItemID				= itemid
	o1300k.FRectItemName			= itemname
	o1300k.FRectSellYn				= sellyn
	o1300k.FRectLimitYn				= limityn
	o1300k.FRectSailYn				= sailyn
	o1300k.FRectStartMargin			= startMargin
	o1300k.FRectEndMargin			= endMargin
	o1300k.FRectMakerid				= makerid
	o1300k.FRectWetoo1300kGoodNo	= wetoo1300kGoodNo
	o1300k.FRectMatchCate			= MatchCate
	o1300k.FRectIsMadeHand			= isMadeHand
	o1300k.FRectIsOption			= isOption
	o1300k.FRectIsReged				= isReged
	o1300k.FRectNotinmakerid		= notinmakerid
	o1300k.FRectNotinitemid			= notinitemid
	o1300k.FRectExcTrans			= exctrans
	o1300k.FRectPriceOption			= priceOption
	o1300k.FRectIsSpecialPrice     	= isSpecialPrice
	o1300k.FRectDeliverytype		= deliverytype
	o1300k.FRectMwdiv				= mwdiv
	o1300k.FRectScheduleNotInItemid	= scheduleNotInItemid
	o1300k.FRectIsextusing			= isextusing
	o1300k.FRectCisextusing			= cisextusing
	o1300k.FRectRctsellcnt			= rctsellcnt
	o1300k.FRectMatchBrand			= MatchBrand
	
	o1300k.FRectExtNotReg			= ExtNotReg
	o1300k.FRectExpensive10x10		= expensive10x10
	o1300k.FRectdiffPrc				= diffPrc
	o1300k.FRectWetoo1300kYes10x10No= wetoo1300kYes10x10No
	o1300k.FRectWetoo1300kNo10x10Yes= wetoo1300kNo10x10Yes
	o1300k.FRectWetoo1300kKeepSell	= wetoo1300kKeepSell
	o1300k.FRectExtSellYn			= extsellyn
	o1300k.FRectInfoDiv				= infoDiv
	o1300k.FRectFailCntOverExcept	= ""
	o1300k.FRectFailCntExists		= failCntExists
	o1300k.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    o1300k.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    o1300k.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	o1300k.getWetoo1300kreqExpireItemList
Else
	o1300k.getWetoo1300kRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=wetoo1300kList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=wetoo1300k","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=wetoo1300k','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=wetoo1300k','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=wetoo1300k','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="wetoo1300kKeepSell")&&(frm.wetoo1300kKeepSell.checked)){ frm.wetoo1300kKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="wetoo1300kYes10x10No")&&(frm.wetoo1300kYes10x10No.checked)){ frm.wetoo1300kYes10x10No.checked=false }
	if ((comp.name!="wetoo1300kNo10x10Yes")&&(frm.wetoo1300kNo10x10Yes.checked)){ frm.wetoo1300kNo10x10Yes.checked=false }
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

    if ((comp.name=="wetoo1300kYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="wetoo1300kNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.wetoo1300kYes10x10No.checked){
            comp.form.wetoo1300kYes10x10No.checked = false;
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
	if ((comp.name!="wetoo1300kYes10x10No")&&(frm.wetoo1300kYes10x10No.checked)){ frm.wetoo1300kYes10x10No.checked=false }
	if ((comp.name!="wetoo1300kNo10x10Yes")&&(frm.wetoo1300kNo10x10Yes.checked)){ frm.wetoo1300kNo10x10Yes.checked=false }
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
//브랜드 관리
function pop_BrandManager() {
	var pCM2 = window.open("/admin/etc/wetoo1300k/popwetoo1300kBrandList.asp","popBrandwetoo1300kmanager","width=1200,height=400,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/wetoo1300k/popwetoo1300kcateList.asp","popCatewetoo1300kmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// 선택된 상품 일괄 등록
function wetoo1300kSelectRegProcess() {
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

    if (confirm('1300k에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※1300k와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 상태 변경
function wetoo1300kSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 가격 수정
function wetoo1300kPriceEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※1300k와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function wetoo1300kEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※1300k와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
		document.frmSvArr.submit();
    }
}
//상품 삭제
function wetoo1300kDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\n1300k 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=wetoo1300k&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function wetoo1300kEditProcessAuto() {
	var cnt = <%= o1300k.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/wetoo1300k/actwetoo1300kReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= o1300k.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		wetoo1300kEditProcessAuto();
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
		<a href="https://kscm.1300k.com" target="_blank">1300kAdmin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenbyten10 | ten101010!* ]</font>"
			End If
		%>
		&nbsp;
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		1300k 상품코드 : <textarea rows="2" cols="20" name="wetoo1300kGoodNo" id="itemid"><%=replace(wetoo1300kGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >1300k 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >1300k 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >1300k 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>1300k 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wetoo1300kYes10x10No" <%= ChkIIF(wetoo1300kYes10x10No="on","checked","") %> ><font color=red>1300k판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="wetoo1300kNo10x10Yes" <%= ChkIIF(wetoo1300kNo10x10Yes="on","checked","") %> ><font color=red>1300k품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wetoo1300kKeepSell" <%= ChkIIF(wetoo1300kKeepSell="on","checked","") %> ><font color=red>판매유지</font> 해야할 상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리(1300k는 원단위 안씀)<br />
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('wetoo1300k');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="브랜드코드" onclick="pop_BrandManager();">
				&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="wetoo1300kSelectRegProcess();">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="wetoo1300kPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="수정" onClick="wetoo1300kEditProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="상품삭제" onClick="wetoo1300kDeleteProcess();" style=font-weight:bold>
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="wetoo1300kSellYnProcess(frmReg.chgSellYn.value);">
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
		검색결과 : <b><%= FormatNumber(o1300k.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(o1300k.FTotalPage,0) %></b>
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
	<td width="140">wetoo1300k등록일<br>wetoo1300k최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">wetoo1300k<br>가격및판매</td>
	<td width="70">wetoo1300k<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to o1300k.FResultCount - 1 %>
<tr align="center" <% If o1300k.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= o1300k.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= o1300k.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= o1300k.FItemList(i).FItemID %>','wetoo1300k','<%=o1300k.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=o1300k.FItemList(i).FItemID%>" target="_blank"><%= o1300k.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If o1300k.FItemList(i).FWetoo1300kStatcd <> 7 Then
	%>
		<br><%= o1300k.FItemList(i).getWetoo1300kStatName %>
	<%
			End If
			response.write o1300k.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= o1300k.FItemList(i).FMakerid %> <%= o1300k.FItemList(i).getDeliverytypeName %><br><%= o1300k.FItemList(i).FItemName %></td>
	<td align="center"><%= o1300k.FItemList(i).FRegdate %><br><%= o1300k.FItemList(i).FLastupdate %></td>
	<td align="center"><%= o1300k.FItemList(i).FWetoo1300kRegdate %><br><%= o1300k.FItemList(i).FWetoo1300kLastUpdate %></td>
	<td align="right">
		<% If o1300k.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(o1300k.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(o1300k.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(o1300k.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If o1300k.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (o1300k.FItemList(i).Fwetoo1300kStatCd > 0) and Not IsNULL(o1300k.FItemList(i).Fwetoo1300kPrice) Then
		' 	If (o1300k.FItemList(i).FSaleYn = "Y") and (CLng((1.0*o1300k.FItemList(i).FSellcash/10)*10) < o1300k.FItemList(i).Fwetoo1300kPrice) Then
		' 		'// 제휴몰 정상가 판매중
		' %>
		' <strike><%= CLng(10000-o1300k.FItemList(i).Fbuycash/o1300k.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-o1300k.FItemList(i).Fbuycash/o1300k.FItemList(i).FWetoo1300kPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-o1300k.FItemList(i).Fbuycash/o1300k.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (o1300k.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (o1300k.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-o1300k.FItemList(i).FOrgSuplycash/o1300k.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-o1300k.FItemList(i).Fbuycash/o1300k.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-o1300k.FItemList(i).Fbuycash/o1300k.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If o1300k.FItemList(i).IsSoldOut Then
			If o1300k.FItemList(i).FSellyn = "N" Then
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
		If o1300k.FItemList(i).FItemdiv = "06" OR o1300k.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (o1300k.FItemList(i).FWetoo1300kStatCd > 0) Then
			If Not IsNULL(o1300k.FItemList(i).FWetoo1300kPrice) Then
				If (o1300k.FItemList(i).Mustprice <> o1300k.FItemList(i).FWetoo1300kPrice) Then
	%>
					<strong><%= formatNumber(o1300k.FItemList(i).FWetoo1300kPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(o1300k.FItemList(i).FWetoo1300kPrice,0)&"<br>"
				End If

				If Not IsNULL(o1300k.FItemList(i).FSpecialPrice) Then
					If (now() >= o1300k.FItemList(i).FStartDate) And (now() <= o1300k.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(o1300k.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (o1300k.FItemList(i).FSellyn="Y" and o1300k.FItemList(i).FWetoo1300kSellYn<>"Y") or (o1300k.FItemList(i).FSellyn<>"Y" and o1300k.FItemList(i).FWetoo1300kSellYn="Y") Then
	%>
					<strong><%= o1300k.FItemList(i).FWetoo1300kSellYn %></strong>
	<%
				Else
					response.write o1300k.FItemList(i).FWetoo1300kSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If o1300k.FItemList(i).Fwetoo1300kGoodNo <> "" Then %>
			<a target="_blank" href="http://www.1300k.com/shop/goodsDetail.html?f_goodsno=<%=o1300k.FItemList(i).Fwetoo1300kGoodNo%>"><%=o1300k.FItemList(i).Fwetoo1300kGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= o1300k.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=o1300k.FItemList(i).FItemID%>','0');"><%= o1300k.FItemList(i).FoptionCnt %>:<%= o1300k.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= o1300k.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If o1300k.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨(카)"
		Else
			response.write "<font color='darkred'>매칭안됨(카)</font>"
		End If

		If o1300k.FItemList(i).FBrandCode <> "" Then
			response.write "<br />매칭됨(브)"
		Else
			response.write "<br /><font color='darkred'>매칭안됨(브)</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= o1300k.FItemList(i).FinfoDiv %>
		<%
		If (o1300k.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(o1300k.FItemList(i).FlastErrStr) &"'>ERR:"& o1300k.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% wetoo1300kGoodNoArray = wetoo1300kGoodNoArray & o1300k.FItemList(i).Fwetoo1300kGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= wetoo1300kGoodNoArray %></textarea>
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
        <% if o1300k.HasPreScroll then %>
		<a href="javascript:goPage('<%= o1300k.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + o1300k.StartScrollPage to o1300k.FScrollCount + o1300k.StartScrollPage - 1 %>
    		<% if i>o1300k.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if o1300k.HasNextScroll then %>
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
	<input type="hidden" name="wetoo1300kGoodNo" value= <%= wetoo1300kGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="wetoo1300kYes10x10No" value= <%= wetoo1300kYes10x10No %>>
	<input type="hidden" name="wetoo1300kNo10x10Yes" value= <%= wetoo1300kNo10x10Yes %>>
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
<% SET o1300k = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
