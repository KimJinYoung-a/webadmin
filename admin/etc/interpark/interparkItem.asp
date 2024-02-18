<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/interpark/interparkcls.asp"-->
<%
Server.ScriptTimeOut = 60 * 5		' 5분
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, interparkPrdno, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, interparkYes10x10No, interparkNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing, cisextusing, rctsellcnt
Dim page, i, research, InterparkGoodNoArray, scheduleNotInItemid
Dim oInterpark, xl
Dim startMargin, endMargin
Dim sSdate, sEdate
Dim purchasetype

sSdate					= requestCheckVar(request("iSD"),10)
sEdate					= requestCheckVar(request("iED"),10)

If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = Date()

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
interparkPrdno			= request("interparkPrdno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
interparkYes10x10No		= request("interparkYes10x10No")
interparkNo10x10Yes		= request("interparkNo10x10Yes")
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
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid="1548380,1707616,1854914,1824546,1827550,1827550,1824546,1754245,932141,1428639,932141,1818430,1755622,1824546,419748,1622289"

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
'인터파크 상품코드 엔터키로 검색되게
If interparkPrdno <> "" then
	Dim iA2, arrTemp2, arrinterparkPrdno
	interparkPrdno = replace(interparkPrdno,",",chr(10))
	interparkPrdno = replace(interparkPrdno,chr(13),"")
	arrTemp2 = Split(interparkPrdno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrinterparkPrdno = arrinterparkPrdno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	interparkPrdno = left(arrinterparkPrdno,len(arrinterparkPrdno)-1)
End If

Set oInterpark = new CInterpark


If (session("ssBctID")="kjy8517") Then
	oInterpark.FPageSize					= 100
Else
	oInterpark.FPageSize					= 50
End If
	oInterpark.FCurrPage					= page
	oInterpark.FRectMakerid					= makerid
	oInterpark.FRectItemID					= itemid
	oInterpark.FRectItemName				= itemname
	oInterpark.FRectInterparkPrdno			= interparkPrdno
	oInterpark.FRectCDL						= request("cdl")
	oInterpark.FRectCDM						= request("cdm")
	oInterpark.FRectCDS						= request("cds")
	oInterpark.FRectExtNotReg				= ExtNotReg
	oInterpark.FRectIsReged					= isReged
	oInterpark.FRectNotinmakerid			= notinmakerid
	oInterpark.FRectNotinitemid				= notinitemid
	oInterpark.FRectExcTrans				= exctrans
	oInterpark.FRectPriceOption				= priceOption
	oInterpark.FRectIsSpecialPrice     		= isSpecialPrice
	oInterpark.FRectDeliverytype			= deliverytype
	oInterpark.FRectMwdiv					= mwdiv
	oInterpark.FRectIsextusing				= isextusing
	oInterpark.FRectCisextusing				= cisextusing
	oInterpark.FRectRctsellcnt				= rctsellcnt

	oInterpark.FRectSellYn					= sellyn
	oInterpark.FRectLimitYn					= limityn
	oInterpark.FRectSailYn					= sailyn
'	oInterpark.FRectonlyValidMargin			= onlyValidMargin
	oInterpark.FRectStartMargin				= startMargin
	oInterpark.FRectEndMargin				= endMargin
	oInterpark.FRectIsMadeHand				= isMadeHand
	oInterpark.FRectIsOption				= isOption
	oInterpark.FRectInfoDiv					= infoDiv
	oInterpark.FRectExtSellYn				= extsellyn
	oInterpark.FRectFailCntExists			= failCntExists
	oInterpark.FRectMatchCate				= MatchCate
	oInterpark.FRectExpensive10x10			= expensive10x10
	oInterpark.FRectdiffPrc					= diffPrc
	oInterpark.FRectInterparklYes10x10No	= interparkYes10x10No
	oInterpark.FRectInterparkNo10x10Yes		= interparkNo10x10Yes
	oInterpark.FRectReqEdit					= reqEdit
	oInterpark.FRectScheduleNotInItemid		= scheduleNotInItemid
	oInterpark.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oInterpark.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oInterpark.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oInterpark.getInterParkreqExpireItemList
Else
	oInterpark.getInterParkRegedItemList			'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=interparkList"& replace(DATE(), "-", "") &"_xl.xls"
Else
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
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

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=interpark","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=interpark','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=interpark','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function checkNDelItem(){
	alert('강제삭제 막음..운영기획팀 진영대리에게 문의');
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
	if ((comp.name!="interparkYes10x10No")&&(frm.interparkYes10x10No.checked)){ frm.interparkYes10x10No.checked=false }
	if ((comp.name!="interparkNo10x10Yes")&&(frm.interparkNo10x10Yes.checked)){ frm.interparkNo10x10Yes.checked=false }
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

    if ((comp.name=="interparkYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="interparkNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.interparkYes10x10No.checked){
            comp.form.interparkYes10x10No.checked = false;
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
	if ((comp.name!="interparkYes10x10No")&&(frm.interparkYes10x10No.checked)){ frm.interparkYes10x10No.checked=false }
	if ((comp.name!="interparkNo10x10Yes")&&(frm.interparkNo10x10Yes.checked)){ frm.interparkNo10x10Yes.checked=false }
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
// function pop_CateManager() {
// 	var pCM2 = window.open("/admin/etc/interpark/popinterparkcateList.asp","popinterparkcateList","width=1100,height=700,scrollbars=yes,resizable=yes");
// 	pCM2.focus();
// }

//카테고리New 관리
function pop_NewCateManager() {
	var pCate = window.open("/admin/etc/interpark/popNewinterparkcateList.asp","popCateinterparkmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCate.focus();
}
function popItem2CategoryRedirect(itemid){
    var popwin = window.open('/admin/etc/interpark/InterParkMatcheDispCateByitemRedirect.asp?itemid=' + itemid,'MatcheDispCate','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=interpark&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function EditIParkSupplyCtrtSeq(iitemid){
    var popwin = window.open('/admin/etc/interpark/EditIParkSupplyCtrtSeq.asp?itemid=' + iitemid,'EditIParkSupplyCtrtSeq','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=interpark','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 선택된 상품 일괄 등록
function interparkregIMSI(isreg) {
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
        if (confirm('인터파크에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?')){
			document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('인터파크에 선택하신 ' + chkSel + '개 상품을 삭제 하시겠습니까?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
        }
    }
}
//선택 상품 조회
function interparkStatCheckProcess(){
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

    if (confirm('인터파크에 선택하신 ' + chkSel + '개 상품을 판매 상태를 조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
        document.frmSvArr.submit();
    }
}
//상품 삭제
function interparkDeleteProcess(){
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
    if (confirm('API로 삭제하는 기능이 아닙니다.\n\n인터파크 어드민에서 삭제후 처리해야 합니다.\n\n ' + chkSel + '개 삭제 하시겠습니까?')){
		if (confirm('정말 삭제하시겠습니까? 확인버튼 클릭시 DB에서 상품이 삭제됩니다.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
		}
    }
}
function interparkCateProcess(){
	if (confirm('인터파크 카테고리를 가져 오시겠습니까?')){
		document.frmSvArr.target = "xLink";iSD
		document.frmSvArr.param1.value = $("#iSD").val();
		document.frmSvArr.param2.value = $("#iED").val();
		document.frmSvArr.cmdparam.value = "CATE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
	}
}
// 선택된 상품 판매여부 변경
function InterParkSellYnProcess(chkYn) {
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

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※인터파크와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 인터파크에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function interparkSelectRegProcess() {
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

    if (confirm('인터파크에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※인터파크와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function interparkEditProcess() {
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

    if (confirm('인터파크에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※인터파크의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

<% if request("auto") = "Y" then %>
function interparkEditProcessAuto() {
	var cnt = <%= oInterpark.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oInterpark.FResultCount %>;
	if (cnt === 0) {
		// 45분뒤 새로고침
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		interparkEditProcessAuto();
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
		<a href="http://ipss.interpark.com/member/login.do?_method=initial&GNBLogin=Y&wid1=wgnb&wid2=wel_login&wid3=seller" target="_blank">(구)인터파크Admin바로가기</a>
		<a href="https://seller.interpark.com/login" target="_blank">(신)인터파크Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ coolhass | store10x10 ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		인터파크 상품코드 : <textarea rows="2" cols="20" name="interparkPrdno" id="itemid"><%=replace(interparkPrdno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >인터파크 등록예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >인터파크 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>인터파크 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="interparkYes10x10No" <%= ChkIIF(interparkYes10x10No="on","checked","") %> ><font color=red>인터파크판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="interparkNo10x10Yes" <%= ChkIIF(interparkNo10x10Yes="on","checked","") %> ><font color=red>인터파크품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리(인터파크는 원단위 안씀)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : <br />

<p />
<% end if %>
<form name="frmReg" method="post" action="interparkItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('interpark');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="카테고리" onclick="pop_NewCateManager();" style=color:blue;font-weight:bold>
				<!--
				<input class="button" type="button" value="카테고리" onclick="pop_CateManager();">
				-->
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
				예정상품 관련 :
				<input class="button" type="button" id="btnRegImsi" value="예정등록" onClick="interparkregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="예정삭제" onClick="interparkregIMSI(false);" >
				<br><br>
				실제상품 관련 :
				<input class="button" type="button" id="btnRegSel" value="전시등록" onClick="interparkSelectRegProcess(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="전시수정" onClick="interparkEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="전시조회" onClick="interparkStatCheckProcess();">
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="상품삭제" onClick="interparkDeleteProcess();" style=font-weight:bold>
			<% End If %>
				<br><br>
				카테고리 조회 :
				<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				<input class="button" type="button" id="btnCate" value="조회" onClick="interparkCateProcess();">
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
					<option value="X">판매종료(삭제)</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="InterParkSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="param1">
<input type="hidden" name="param2">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= FormatNumber(oInterpark.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oInterpark.FTotalPage,0) %></b>
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
	<td width="140">상품예정등록일<br>상품최종수정일</td>
	<td width="140">인터파크등록일<br>인터파크최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">인터파크<br>가격및판매</td>
	<!-- <td width="60">구분</td> -->
	<td width="70">인터파크<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="100">카테고리<br>매칭여부</td>
	<td width="50">품목</td>
</tr>
<% For i=0 to oInterpark.FResultCount - 1 %>
<tr align="center" <% If oInterpark.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oInterpark.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oInterpark.FItemList(i).Fsmallimage %>" width="50"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oInterpark.FItemList(i).FItemID%>" target="_blank"><%= oInterpark.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oInterpark.FItemList(i).FLimitYn="Y" Then
	%>
			<br><%= oInterpark.FItemList(i).getLimitHtmlStr %></font>
	<%
			End If
			response.write "<br>"&oInterpark.FItemList(i).getiParkRegStateName
		End If
	%>
	</td>
	<td align="left"><%= oInterpark.FItemList(i).FMakerid %><%= oInterpark.FItemList(i).getDeliverytypeName %><br><%= oInterpark.FItemList(i).FItemName %></td>
	<td align="center"><%= oInterpark.FItemList(i).FiparkTmpregdate %><br><%= oInterpark.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oInterpark.FItemList(i).FinterparkRegdate %><br><%= oInterpark.FItemList(i).FinterparkLastUpdate %></td>

	<td align="right">
	<% If oInterpark.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oInterpark.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oInterpark.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oInterpark.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
		<%
		If oInterpark.FItemList(i).Fsellcash = 0 Then
		elseif (oInterpark.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oInterpark.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oInterpark.FItemList(i).FOrgSuplycash/oInterpark.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oInterpark.FItemList(i).Fbuycash/oInterpark.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oInterpark.FItemList(i).Fbuycash/oInterpark.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oInterpark.FItemList(i).IsSoldOut Then
			If oInterpark.FItemList(i).FSellyn = "N" Then
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
		If oInterpark.FItemList(i).FItemdiv = "06" OR oInterpark.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
    <% if Not IsNULL(oInterpark.FItemList(i).FmayiParkPrice) then %>
        <% if (oInterpark.FItemList(i).Fsellcash<>oInterpark.FItemList(i).FmayiParkPrice) then %>
        <strong><%= formatNumber(oInterpark.FItemList(i).FmayiParkPrice,0) %></strong>
        <% else %>
        <%= formatNumber(oInterpark.FItemList(i).FmayiParkPrice,0) %>
        <% end if %>
        <br>

		<%
			If Not IsNULL(oInterpark.FItemList(i).FSpecialPrice) Then
				If (now() >= oInterpark.FItemList(i).FStartDate) And (now() <= oInterpark.FItemList(i).FEndDate) Then
					response.write "<font color='orange'><strong>(특)" & formatNumber(oInterpark.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
				End If
			End If
		%>

        <% if (oInterpark.FItemList(i).FmayiParkSellYn="X") then %>
        <a href="javascript:checkNDelItem('<%= oInterpark.FItemList(i).FItemID %>')">
        <% end if %>

        <% if (oInterpark.FItemList(i).FSellyn<>oInterpark.FItemList(i).FmayiParkSellYn) then %>
        <strong><%= oInterpark.FItemList(i).FmayiParkSellYn %></strong>
        <% else %>
        <%= oInterpark.FItemList(i).FmayiParkSellYn %>
        <% end if %>

        <% if (oInterpark.FItemList(i).FmayiParkSellYn="X") then %>
        </a>
        <% end if %>
    <% end if %>
	</td>
<!--
	<td align="center"><a href="javascript:EditIParkSupplyCtrtSeq('oInterpark.FItemList(i).FItemID')"> oInterpark.FItemList(i).GetExtStoreSeqName ( oInterpark.FItemList(i).FSupplyCtrtSeq )</a></td>
-->
	<td align="center">
		<a target=_blank href="https://shopping.interpark.com/product/productInfo.do?prdNo=<%= oInterpark.FItemList(i).FInterparkPrdno %>"><%= oInterpark.FItemList(i).FInterparkPrdno %></a>
		<% If IsNULL(oInterpark.FItemList(i).FInterparkPrdno) Then %>
			<a href="javascript:DelTenIparkItem('<%= oInterpark.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
		<% End If %>
	</td>
	<td align="center"><%= oInterpark.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oInterpark.FItemList(i).FItemID%>','0');"><%= oInterpark.FItemList(i).FoptionCnt %>:<%= oInterpark.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oInterpark.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
'		If IsNULL(oInterpark.FItemList(i).FCatekey) then
'			<font color="darkred">매칭안됨</font><br>
'	 	Else
'			<a href="javascript:popItem2CategoryRedirect('oInterpark.FItemList(i).FItemID');">oInterpark.FItemList(i).Finterparkdispcategory</a>
'		End If

		If oInterpark.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If

		If (oInterpark.FItemList(i).FaccFailCNT > 0) Then
	    	response.write "<br><font color='red' title='" & oInterpark.FItemList(i).FlastErrStr & "'>ERR: " & oInterpark.FItemList(i).FaccFailCNT & "</font>"
		End If
	%>
	</td>
    <td align="center"><%= oInterpark.FItemList(i).FinfoDiv %></td>
</tr>
<% InterparkGoodNoArray = InterparkGoodNoArray & oInterpark.FItemList(i).FInterparkPrdno & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= InterparkGoodNoArray %></textarea>
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
        <% if oInterpark.HasPreScroll then %>
		<a href="javascript:goPage('<%= oInterpark.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oInterpark.StartScrollPage to oInterpark.FScrollCount + oInterpark.StartScrollPage - 1 %>
    		<% if i>oInterpark.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oInterpark.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
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
	<input type="hidden" name="interparkPrdno" value= <%= interparkPrdno %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="interparkYes10x10No" value= <%= interparkYes10x10No %>>
	<input type="hidden" name="interparkNo10x10Yes" value= <%= interparkNo10x10Yes %>>
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
<% SET oInterpark = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
