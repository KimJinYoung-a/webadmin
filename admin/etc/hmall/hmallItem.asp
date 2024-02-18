<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/hmall/hmallCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, hmallGoodNo, hmallGoodNo2, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans, MatchIMG
Dim expensive10x10, diffPrc, hmallYes10x10No, hmallNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, setMargin, isextusing, scheduleNotInItemid
Dim page, i, research, cisextusing, rctsellcnt
Dim oHmall, xl
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
hmallGoodNo				= request("hmallGoodNo")
hmallGoodNo2			= request("hmallGoodNo2")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchIMG				= request("MatchIMG")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
hmallYes10x10No			= request("hmallYes10x10No")
hmallNo10x10Yes			= request("hmallNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
setMargin				= request("setMargin")
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
	MatchIMG = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchIMG = ""
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
'hmall 상품코드 엔터키로 검색되게
If hmallGoodNo2 <> "" then
	Dim iA2, arrTemp2, arrhmallGoodNo2
	hmallGoodNo2 = replace(hmallGoodNo2,",",chr(10))
	hmallGoodNo2 = replace(hmallGoodNo2,chr(13),"")
	arrTemp2 = Split(hmallGoodNo2,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrhmallGoodNo2 = arrhmallGoodNo2 & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	hmallGoodNo2 = left(arrhmallGoodNo2,len(arrhmallGoodNo2)-1)
End If

Set oHmall = new CHmall
	oHmall.FCurrPage					= page
	oHmall.FPageSize					= 100
	oHmall.FRectCDL						= request("cdl")
	oHmall.FRectCDM						= request("cdm")
	oHmall.FRectCDS						= request("cds")
	oHmall.FRectItemID					= itemid
	oHmall.FRectItemName				= itemname
	oHmall.FRectSellYn					= sellyn
	oHmall.FRectLimitYn					= limityn
	oHmall.FRectSailYn					= sailyn
'	oHmall.FRectonlyValidMargin			= onlyValidMargin
	oHmall.FRectStartMargin				= startMargin
	oHmall.FRectEndMargin				= endMargin
	oHmall.FRectMakerid					= makerid
	oHmall.FRectHmallGoodNo				= hmallGoodNo2
	oHmall.FRectMatchCate				= MatchCate
	oHmall.FRectMatchIMG				= MatchIMG
	oHmall.FRectIsMadeHand				= isMadeHand
	oHmall.FRectIsOption				= isOption
	oHmall.FRectIsReged					= isReged
	oHmall.FRectNotinmakerid			= notinmakerid
	oHmall.FRectNotinitemid				= notinitemid
	oHmall.FRectExcTrans				= exctrans
	oHmall.FRectPriceOption				= priceOption
	oHmall.FRectIsSpecialPrice			= isSpecialPrice
	oHmall.FRectDeliverytype			= deliverytype
	oHmall.FRectMwdiv					= mwdiv
	oHmall.FRectSetMargin				= setMargin
	oHmall.FRectScheduleNotInItemid		= scheduleNotInItemid
	oHmall.FRectIsextusing				= isextusing
	oHmall.FRectCisextusing				= cisextusing
	oHmall.FRectRctsellcnt				= rctsellcnt

	oHmall.FRectExtNotReg				= ExtNotReg
	oHmall.FRectExpensive10x10			= expensive10x10
	oHmall.FRectdiffPrc					= diffPrc
	oHmall.FRectHmallYes10x10No			= hmallYes10x10No
	oHmall.FRectHmallNo10x10Yes			= hmallNo10x10Yes
	oHmall.FRectExtSellYn				= extsellyn
	oHmall.FRectInfoDiv					= infoDiv
	oHmall.FRectFailCntOverExcept		= ""
	oHmall.FRectFailCntExists			= failCntExists
	oHmall.FRectReqEdit					= reqEdit
	oHmall.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oHmall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHmall.FRectOrdType = "BM"
End If


If isReged = "R" Then					'품절처리요망 상품보기 리스트
	oHmall.getHmallreqExpireItemList
Else
	oHmall.getHmallRegedItemList		'그 외 리스트
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=hmall1010List"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=hmall1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=hmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 카테고리
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=hmall1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//마진 변경 카테고리
function popMarginCateList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginCateList.asp?mallid=hmall1010','popMarginCateList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//마진 변경 상품
function popMarginItemList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginItemList.asp?mallid=hmall1010','popMarginItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// 스케줄 제외 상품
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=hmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallCateList.asp","popCateHmallmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//전시매장
function pop_SectId() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallSectList.asp","popSectId","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//매장매칭건
function pop_SectId2() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallSectList2.asp","popSectId2","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
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
	if ((comp.name!="hmallYes10x10No")&&(frm.hmallYes10x10No.checked)){ frm.hmallYes10x10No.checked=false }
	if ((comp.name!="hmallNo10x10Yes")&&(frm.hmallNo10x10Yes.checked)){ frm.hmallNo10x10Yes.checked=false }
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

    if ((comp.name=="hmallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="hmallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.hmallYes10x10No.checked){
            comp.form.hmallYes10x10No.checked = false;
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
	if ((comp.name!="hmallYes10x10No")&&(frm.hmallYes10x10No.checked)){ frm.hmallYes10x10No.checked=false }
	if ((comp.name!="hmallNo10x10Yes")&&(frm.hmallNo10x10Yes.checked)){ frm.hmallNo10x10Yes.checked=false }
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

//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_hmall.asp?itemid="+iitemid+'&mallid=hmall1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 선택된 상품 일괄 등록
function hmallSelectRegProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 이미지 등록
function hmallSelectImagesProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 이미지를 등록 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "IMAGE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 등록	''2018-12-10 김진영 추가
function hmallSelectRegItemProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGAddItem";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 이미지 등록	''2018-12-10 김진영 추가
function hmallSelectImageProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 이미지를 등록 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGImage";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 등록	''2018-12-10 김진영 추가
function hmallSelectImageConfirmProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 이미지를 확인 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGImageConfirm";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 상태 변경
function hmallSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//수정
function hmallSelectEditProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 수정 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//수정 ''2018-12-10 김진영 추가
function hmallSelectEditItemProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 수정 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITItem";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//가격 수정
function hmallSelectPriceEditProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 단품 가격을 수정 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//옵션 수정
function hmallSelectOptionEditProcess() {
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

    if (confirm('HMall에 선택하신 ' + chkSel + '개 상품의 옵션을 수정 하시겠습니까?\n\n※HMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSelOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 상세 조회
function hmallSelectViewProcess() {
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

	if (confirm('HMall에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 상세 조회
function hmallSelectViewProcess2() {
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

	if (confirm('HMall에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//상품 재고 조회
function hmallSelectOptionViewProcess() {
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

	if (confirm('HMall에 선택하신 ' + chkSel + '개 재고를 조회 하시겠습니까?')){
        //document.getElementById("btnOptViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
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
		https://partner.hmall.com/
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ hs0027011 | cube101010 ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		hmall 상품코드 : <textarea rows="2" cols="20" name="hmallGoodNo2" id="itemid"><%=replace(hmallGoodNo2,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Hmall 등록실패
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >Hmall 등록예정
			<option value="E" <%= CHkIIF(ExtNotReg="E","selected","") %> >Hmall 취소
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >Hmall 반려
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >Hmall 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Hmall 등록완료(전시)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Hmall 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="hmallYes10x10No" <%= ChkIIF(hmallYes10x10No="on","checked","") %> ><font color=red>Hmall판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="hmallNo10x10Yes" <%= ChkIIF(hmallNo10x10Yes="on","checked","") %> ><font color=red>Hmall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 할인이 아닐 때 판매가 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 주문제작문구 상품, 옵션추가금액 있는 상품<br />

<p />
<form name="frmReg" method="post" action="hmallItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="매입마진변경(카테고리)" onclick="popMarginCateList();">&nbsp;
				<input class="button" type="button" value="매입마진변경(상품)" onclick="popMarginItemList();">&nbsp;
				<input class="button" type="button" value="스케쥴 제외 상품" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('hmall1010');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="카테고리" onclick="pop_CateManager();">

				<input class="button" type="button" value="전시매장" onclick="pop_SectId();">
				<input class="button" type="button" value="매장매칭건" onclick="pop_SectId2();">
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="hmallSelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnImages" value="이미지" onClick="hmallSelectImagesProcess();">
				<% If (session("ssBctID")="kjy8517") Then %>
			<!--
					&nbsp;&nbsp;<input class="button" type="button" id="btnRegItem" value="상품" onClick="hmallSelectRegItemProcess();">
					&nbsp;&nbsp;<input class="button" type="button" id="btnImage" value="이미지등록" onClick="hmallSelectImageProcess();">
					&nbsp;&nbsp;<input class="button" type="button" id="btnImageConfirm" value="이미지확인" onClick="hmallSelectImageConfirmProcess();">
			-->
				<% End If %>
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSel" value="수정" onClick="hmallSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="가격" onClick="hmallSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelOption" value="옵션" onClick="hmallSelectOptionEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="재고조회" onClick="hmallSelectOptionViewProcess();">&nbsp;&nbsp;
				<% If (session("ssBctID")="kjy8517") Then %>
			<!--
					<input class="button" type="button" id="btnEditItem" value="상품수정" onClick="hmallSelectEditItemProcess();">
			-->
				<% End If %>
				<br><br>
				승인예정 상품 :
				<!-- 
				<input class="button" type="button" id="btnViewSel" value="상세조회" onClick="hmallSelectViewProcess();">&nbsp;&nbsp;
				--> 
				<input class="button" type="button" id="btnViewSel" value="상세조회2" onClick="hmallSelectViewProcess2();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="hmallSellYnProcess(frmReg.chgSellYn.value);">
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= FormatNumber(oHmall.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHmall.FTotalPage,0) %></b>
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
	<td width="140">Hmall등록일<br>Hmall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Hmall<br>가격및판매</td>
	<td width="70">Hmall<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="50">적용마진</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
	<td width="100">이미지<br />등록 | 확인</td>
</tr>
<% For i=0 to oHmall.FResultCount - 1 %>
<tr align="center" <% If oHmall.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHmall.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oHmall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHmall.FItemList(i).FItemID %>','hmall1010','<%=oHmall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oHmall.FItemList(i).FItemID%>" target="_blank"><%= oHmall.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oHmall.FItemList(i).FHmallStatcd <> 7 Then
	%>
		<br><%= oHmall.FItemList(i).getHmallStatName %>
	<%
			End If
			response.write oHmall.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oHmall.FItemList(i).FMakerid %> <%= oHmall.FItemList(i).getDeliverytypeName %><br><%= oHmall.FItemList(i).FItemName %></td>
	<td align="center"><%= oHmall.FItemList(i).FRegdate %><br><%= oHmall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHmall.FItemList(i).FHmallRegdate %><br><%= oHmall.FItemList(i).FHmallLastUpdate %></td>
	<td align="right">
		<% If oHmall.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oHmall.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oHmall.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oHmall.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oHmall.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).FHmallPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oHmall.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oHmall.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oHmall.FItemList(i).FOrgSuplycash/oHmall.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oHmall.FItemList(i).IsSoldOut Then
			If oHmall.FItemList(i).FSellyn = "N" Then
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
		If oHmall.FItemList(i).FItemdiv = "06" OR oHmall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oHmall.FItemList(i).FHmallStatCd > 0) Then
			If Not IsNULL(oHmall.FItemList(i).FHmallPrice) Then
				If (oHmall.FItemList(i).Mustprice <> oHmall.FItemList(i).FHmallPrice) Then
	%>
					<strong><%= formatNumber(oHmall.FItemList(i).FHmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oHmall.FItemList(i).FHmallPrice,0)&"<br>"
				End If

				If Not IsNULL(oHmall.FItemList(i).FSpecialPrice) Then
					If (now() >= oHmall.FItemList(i).FStartDate) And (now() <= oHmall.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oHmall.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oHmall.FItemList(i).FSellyn="Y" and oHmall.FItemList(i).FHmallSellYn<>"Y") or (oHmall.FItemList(i).FSellyn<>"Y" and oHmall.FItemList(i).FHmallSellYn="Y") Then
	%>
					<strong><%= oHmall.FItemList(i).FHmallSellYn %></strong>
	<%
				Else
					response.write oHmall.FItemList(i).FHmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oHmall.FItemList(i).FHmallGoodNo <> "" Then %>
			<a target="_blank" href="https://www.hmall.com/pd/pda/itemPtc?slitmCd=<%=oHmall.FItemList(i).FHmallGoodNo%>"><%=oHmall.FItemList(i).FHmallGoodNo2%></a>
		<% End If %>
	</td>
	<td align="center"><%= oHmall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHmall.FItemList(i).FItemID%>','0');"><%= oHmall.FItemList(i).FoptionCnt %>:<%= oHmall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHmall.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= oHmall.FItemList(i).FSetMargin %></td>
	<td align="center">
	<%
		If oHmall.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oHmall.FItemList(i).FinfoDiv %>
		<%
		If (oHmall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oHmall.FItemList(i).FlastErrStr) &"'>ERR:"& oHmall.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oHmall.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oHmall.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>"&oHmall.FItemList(i).FAPIaddImg&"</font>") %>&nbsp;|
		<%= Chkiif(oHmall.FItemList(i).FAPIconfirmImg="Y","<font color='BLUE'>"&oHmall.FItemList(i).FAPIconfirmImg&"</font>", "<font color='RED'>"&oHmall.FItemList(i).FAPIconfirmImg&"</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oHmall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHmall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHmall.StartScrollPage to oHmall.FScrollCount + oHmall.StartScrollPage - 1 %>
    		<% if i>oHmall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHmall.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
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
	<input type="hidden" name="hmallGoodNo" value= <%= hmallGoodNo %>>
	<input type="hidden" name="hmallGoodNo2" value= <%= hmallGoodNo2 %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchIMG" value= <%= MatchIMG %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="hmallYes10x10No" value= <%= hmallYes10x10No %>>
	<input type="hidden" name="hmallNo10x10Yes" value= <%= hmallNo10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="setMargin" value= <%= setMargin %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% Set oHmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->