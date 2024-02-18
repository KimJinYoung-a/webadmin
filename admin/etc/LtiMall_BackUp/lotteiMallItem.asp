<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim itemid, itemname, eventid, mode
Dim page, makerid, LotteNotReg, MatchCate, sellyn, limityn, sailyn
Dim delitemid, lotteGoodNo, showminusmagin, expensive10x10, LotteYes10x10No, LotteNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research, lotteTmpGoodNo
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists,optnotExists, regedOptNull,failCntExists, optAddPrcRegTypeNone, isMadeHand
Dim bestOrd, bestOrdMall, extsellyn, infoDiv
Dim i
page    = request("page")
itemid  = request("itemid")

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

eventid					= request("eventid")
itemname				= html2db(request("itemname"))
mode					= request("mode")
makerid					= request("makerid")
LotteNotReg				= request("LotteNotReg")
MatchCate				= request("MatchCate")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
delitemid				= request("delitemid")
lotteGoodNo				= request("lotteGoodNo")
lotteTmpGoodNo			= request("lotteTmpGoodNo")
showminusmagin			= request("showminusmagin")
expensive10x10			= request("expensive10x10")
LotteYes10x10No			= request("LotteYes10x10No")
LotteNo10x10Yes			= request("LotteNo10x10Yes")
onreginotmapping		= request("onreginotmapping")
diffPrc					= request("diffPrc")
onlyValidMargin			= request("onlyValidMargin")
research				= request("research")
reqExpire				= request("reqExpire")
reqEdit					= request("reqEdit")
optAddprcExists			= request("optAddprcExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
optAddprcExistsExcept	= request("optAddprcExistsExcept")
optExists				= request("optExists")
optnotExists   			= request("optnotExists")
regedOptNull			= request("regedOptNull")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
failCntExists			= request("failCntExists")
extsellyn				= request("extsellyn")
infoDiv					= request("infoDiv")
isMadeHand				= request("isMadeHand")

If page = "" Then page = 1
If sellyn = "" Then sellyn = "Y"

''기본조건 등록예정이상
If (research="") Then
    LotteNotReg = "J"
    MatchCate = "" ''Y
    onlyValidMargin="on"    ''2013/05/23수정
    ''bestOrd="on"
    sellyn="Y"              ''2013/05/23수정
End If

Dim oiMall
Set oiMall = new CLotteiMall
If (LotteNotReg="F") then                       '''승인대기
	oiMall.FPageSize 					= 50
Else
	If (session("ssBctID")="kjy8517") Then
	oiMall.FPageSize 					= 30
	Else
	oiMall.FPageSize 					= 20
	End If
End If
	oiMall.FCurrPage					= page
	oiMall.FRectItemID					= itemid
	oiMall.FRectEventid					= eventid
	oiMall.FRectItemName				= itemname
	oiMall.FRectMakerid					= makerid
	oiMall.FRectCDL						= request("cdl")
	oiMall.FRectCDM						= request("cdm")
	oiMall.FRectCDS						= request("cds")
	oiMall.FRectLotteNotReg				= LotteNotReg
	oiMall.FRectMatchCate				= MatchCate
	oiMall.FRectSellYn					= sellyn
	oiMall.FRectLimitYn					= limityn
	oiMall.FRectSailYn					= sailyn
	oiMall.FRectLTiMallGoodNo			= lotteGoodNo
	oiMall.FRectLTiMallTmpGoodNo		= lotteTmpGoodNo
	oiMall.FRectMinusMigin				= showminusmagin
	oiMall.FRectExpensive10x10			= expensive10x10
	oiMall.FRectLotteYes10x10No			= LotteYes10x10No
	oiMall.FRectLotteNo10x10Yes			= LotteNo10x10Yes
	oiMall.FRectOnreginotmapping		= onreginotmapping
	oiMall.FRectdiffPrc					= diffPrc
	oiMall.FRectonlyValidMargin			= onlyValidMargin
	oiMall.FRectoptAddprcExists			= optAddprcExists
	oiMall.FRectoptAddPrcRegTypeNone	= optAddPrcRegTypeNone                         ''옵션추가금액상품 미설정 상품.
	oiMall.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oiMall.FRectoptExists				= optExists
	oiMall.FRectoptnotExists			= optnotExists
	oiMall.FRectregedOptNull			= regedOptNull
	oiMall.FRectFailCntExists			= failCntExists
	oiMall.FRectFailCntOverExcept		= ""
	oiMall.FRectExtSellYn				= extsellyn
	oiMall.FRectInfoDiv					= infoDiv
	oiMall.FRectisMadeHand				= isMadeHand
If (bestOrd = "on") Then
    oiMall.FRectOrdType					 = "B"
ElseIf (bestOrdMall = "on") Then
    oiMall.FRectOrdType					= "BM"
End If

If reqExpire <> "" Then
    oiMall.getLtiMallreqExpireItemList
Else
    oiMall.getLTiMallRegedItemList
End If
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteimall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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

// 선택된 상품 일괄 등록
function LotteSelectRegProcess(isreal) {
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
            document.frmSvArr.cmdparam.value = "RegSelect";
            document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
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

// 선택된 상품 일괄 등록
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
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?\n\n※30분단위로 배치 등록됩니다.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 삭제 하시겠습니까?')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "DelSelectWait";
            document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }
}

// 선택된 상품 일괄 수정
function LotteSelectEditProcess(v) {
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
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect" + v;
        document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 단품 수정
function LotteSelectSaleStatEditProcess() {
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
        document.getElementById("btnEditDanpum").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EdSaleDTSel";
        document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteiMall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteiMall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 임시상품 일괄 갱신
function LotteRealItemMapping() {
	if(confirm("임시등록 상품이 전시상품으로 등록되었는지 일괄 확인하시겠습니까?\n\n※ 통신상태에따라 다소 시간이 걸릴 수 있습니다.")) {
		document.getElementById("btnChkReal").disabled=true;
		xLink.location.href="actLotteiMallCheckRDItem.asp";
	}
}

function LotteGoodnoMapping() {
	if(confirm("상품코드 얻기 확인하시겠습니까?2013-08-29 진영 생성한 것..")) {
		xLink.location.href="actLotteiMallDetailItem.asp";
	}
}

// 선택된 상품 일괄 갱신
function LotteRealItemMappingChecked(chkYn) {
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

    if (confirm('롯데iMall에 선택하신 ' + chkSel + '개 상품 등록확인 하시겠습니까?\n\n※롯데iMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnChkPrdReal").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "getconfirmList";
        document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
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

        //document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.LotteNotReg.value="M";
            frm.MatchCate.value="Y";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="mapRealPrdCode"){
            frm.LotteNotReg.value="F";
            frm.MatchCate.value="Y";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }else if (comp.name=="reqEdit"){
            frm.LotteNotReg.value="R";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.LotteNotReg.value="D";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="LotteNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="LotteYes10x10No")&&(frm.LotteYes10x10No.checked)){ frm.LotteYes10x10No.checked=false }
        if ((comp.name!="LotteNo10x10Yes")&&(frm.LotteNo10x10Yes.checked)){ frm.LotteNo10x10Yes.checked=false }
        if ((comp.name!="reqREG")&&(frm.reqREG.checked)){ frm.reqREG.checked=false }
        if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
        if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
    }
}

function checkComp(comp){
    if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
        if ((comp.name=="bestOrd")&&(comp.checked)){
            comp.form.bestOrdMall.checked=false;
        }

        if ((comp.name=="bestOrdMall")&&(comp.checked)){
            comp.form.bestOrd.checked=false;
        }
    }else if ((comp.name=="optAddprcExists")||(comp.name=="optAddprcExistsExcept")){
        if ((comp.name=="optAddprcExists")&&(comp.checked)){
            comp.form.optAddprcExistsExcept.checked=false;
        }

        if ((comp.name=="optAddprcExistsExcept")&&(comp.checked)){
            comp.form.optAddprcExists.checked=false;
        }
    }
}

function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//상품판매상태,가격Check
function batchStatCheck(){
    xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckItemStatAuto";
}

function checkNdel(iitemid){
    xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckNDel&cksel="+iitemid;
}

function checkNdelReged(iitemid){
    if (confirm('삭제 하시겠습니까?')){
        xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckNDelReged&delitemid="+iitemid;
    }
}
function popApiTest(){
    var popwin = window.open('/admin/etc/ltiMall/lotteiMallApiTest.asp','lotteApiTest','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 선택 상품 판매상태 확인 //20130305
function LotteSelectStatCheck(){
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

    if (confirm('롯데아이몰에 선택하신 ' + chkSel + '개 상품을 판매 상태를 확인하시겠습니까?')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CheckItemStat";
        document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
//상품명 다른내역 수정.
function batchItemNmCheck(){
    xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckItemNmAuto";
}

// 선택된 상품명 수정요청
function LotteSelectItemNmEditProcess() {
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

    if (confirm('롯데아이몰에 선택하신 ' + chkSel + '개 상품 명을 일괄 수정 요청 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnUpNm").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditItemNm";
        document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 가격 수정
function LotteSelectPriceEditProcess() {
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
        document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택상품 재고조회
function LotteSelectcheckStock() {
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
        document.frmSvArr.cmdparam.value = "ChkStockSelect";
        document.frmSvArr.action = "actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
function LotteDateCheck(chkdate){
	chkdate = document.getElementById('startdate').value;
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "ChkDate";
    document.frmSvArr.action = "actLotteiMallReq.asp?chkdate="+chkdate
    document.frmSvArr.submit();
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		브 랜 드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		롯데iMall상품번호: <input type="text" name="lotteGoodNo" value="<%= lotteGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;&nbsp;
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		롯데iMall임시상품번호: <input type="text" name="lotteTmpGoodNo" value="<%= lotteTmpGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;
		<a href="https://partner.lotteimall.com/main/Login.lotte" target="_blank">롯데아이몰Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 011799LT | cube101010 ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품번호: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		이벤트번호: <input type="text" name="eventid" value="<%= eventid %>" size="8" maxlength="6" class="text">
		&nbsp;
		주문제작여부 :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		등록여부 :
		<select name="LotteNotReg" class="select">
		<option value="">전체
		<option value="M" <%= CHkIIF(LotteNotReg="M","selected","") %> >롯데iMall 미등록(등록가능)
		<option value="Q" <%= CHkIIF(LotteNotReg="Q","selected","") %> >롯데iMall 등록실패
		<option value="J" <%= CHkIIF(LotteNotReg="J","selected","") %> >롯데iMall 등록예정이상
		<option value="V" <%= CHkIIF(LotteNotReg="V","selected","") %> >롯데iMall 등록예정/등록가능
		<option value="A" <%= CHkIIF(LotteNotReg="A","selected","") %> >롯데iMall 전송시도중오류
		<option value="F" <%= CHkIIF(LotteNotReg="F","selected","") %> >롯데iMall 등록후 승인대기(임시)
		<option value="D" <%= CHkIIF(LotteNotReg="D","selected","") %> >롯데iMall 등록완료(전시)
		<option value="R" <%= CHkIIF(LotteNotReg="R","selected","") %> >롯데iMall 수정요망
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴)</b>
		&nbsp;
		카테매칭 :
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>
		&nbsp;
		판매여부 :
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >전체
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
		</select>
		&nbsp;
		한정여부 :
		<select name="limityn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >한정
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >일반
		</select>
		&nbsp;
		세일여부 :
		<select name="sailyn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >세일Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >세일N
		</select>

		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >마진 <%= CMAXMARGIN %>%이상 상품만 보기
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품
		&nbsp;
		<input type="checkbox" name="optAddPrcRegTypeNone" <%= ChkIIF(optAddPrcRegTypeNone="on","checked","") %> onClick="checkComp(this)">옵션추가판매미설정상품
		&nbsp;
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품제외
		&nbsp;

		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >옵션존재상품
		&nbsp;
		<input type="checkbox" name="optnotExists" <%= ChkIIF(optnotExists="on","checked","") %> >단품상품(옵션=0)
		&nbsp;
		<input type="checkbox" name="regedOptNull" <%= ChkIIF(regedOptNull="on","checked","") %> >단품목록 미수신
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >등록수정오류상품
		<br><br>
		-- Quick 검색 / 등록 / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >등록가능 상품
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="mapRealPrdCode"  >임시상품 일괄확인
		<br><br>
		-- Quick 검색 / 수정 / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>역마진</font>상품보기 (MaxMagin : <%= CMAXMARGIN %>%) (롯데 판매중)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>롯데iMall 가격<텐바이텐 판매가</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteYes10x10No" <%= ChkIIF(LotteYes10x10No="on","checked","") %> ><font color=red>롯데iMall판매중&텐바이텐품절</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteNo10x10Yes" <%= ChkIIF(LotteNo10x10Yes="on","checked","") %> ><font color=red>롯데iMall품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>품절처리요망</font>상품보기 (제휴몰 사용안함등)
		&nbsp;&nbsp;제휴판매상태 :
		<select name="extsellyn" class="select">
		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >종료
		</select>
		&nbsp;&nbsp;품목정보 :
		<% CALL DrawItemInfoDiv("infoDiv", infoDiv, true, "") %>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<br>
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
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" id="btnChkPrdReal" value="선택상품 등록확인" onClick="LotteRealItemMappingChecked();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal" value="임시상품 일괄확인 [100건씩]" onClick="LotteRealItemMapping();">
				&nbsp;&nbsp;
				<font color="RED">우측 2개 선작업 필요! :</font>
				<input class="button" type="button" value="LotteiMall담당MD" onclick="pop_MDList();"> &nbsp;
				&nbsp;&nbsp;
				<input class="button" type="button" value="LotteiMall카테고리매칭" onclick="pop_CateManager();">
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
	    		실제상품 가공 :
	    		<input class="button" type="button" id="btnRegSel" value="선택상품 실제 등록" onClick="LotteSelectRegProcess(true);">
	    		&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditSel" value="선택상품정보/가격 수정" onClick="LotteSelectEditProcess('');">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditPrice" value="선택상품가격 수정" onClick="LotteSelectPriceEditProcess();">
   			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnSelsStock" value="선택상품 재고조회" onClick="LotteSelectcheckStock();" >
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnUpNm" value="선택상품명 수정요청" onClick="LotteSelectItemNmEditProcess();" >
			    <br><br>
				예정여부 가공 :
			    <input class="button" type="button" id="btnRegSel" value="선택상품 예정 등록" onClick="LotteregIMSI(true);">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="선택상품 예정 삭제" onClick="LotteregIMSI(false);" >
			    &nbsp;&nbsp;
			    <br><br>
			    <input class="button" type="button" id="btnEditSel" value="선택상품정보 수정(등록대기)" onClick="LotteSelectEditProcess('2');">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnStatConfirm" value="선택상품 판매상태확인" onClick="LotteSelectStatCheck();">
				<% If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then %>
			    &nbsp;&nbsp;
			    <br>
			    <input type="text" name="startdate" id="startdate" value="2012-06-12">
			    <input class="button" type="button" id="btnDateConfirm" value="조회 후 등록" onClick="LotteDateCheck();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnDetailConfirm" value="상품코드얻기" onClick="LotteGoodnoMapping();">
				<% End If %>
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
				<% if (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then %>
				<br><input type="button" value="API_TEST(관리자)" class="button" onClick="popApiTest();">
				<% end if %>
				<br><br>
				<input type="button" value="판매상태Check(관리자)" class="button" onClick="batchStatCheck();">
				&nbsp;&nbsp;<input type="button" value="상품명수정(관리자)" class="button" onClick="batchItemNmCheck();">
			</td>
		</tr>
		</table>
    </td>
</tr>
<tr>
    <td>
    예외처리상품(강제 재등록) : 210499,724724,692489
    </td>
</tr>
</table>
</form>
<br>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oiMall.FTotalPage,0) %> 총건수: <%= FormatNumber(oiMall.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td >브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">롯데iMall등록일<br>롯데iMall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">롯데iMall<br>가격및판매</td>
	<td width="70">롯데iMall<br>상품번호<br>(임시번호)</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="80">품목</td>
</tr>
<% for i=0 to oiMall.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oiMall.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oiMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oiMall.FItemList(i).FItemID %>','lotteimall','')" style="cursor:pointer"></td>
    <td align="center"><%= oiMall.FItemList(i).FItemID %>
    <% if oiMall.FItemList(i).FLimitYn="Y" then %><br><%= oiMall.FItemList(i).getLimitHtmlStr %></font><% end if %>
    </td>
    <td><%= oiMall.FItemList(i).FMakerid %> <%= oiMall.FItemList(i).getDeliverytypeName %><br><%= oiMall.FItemList(i).FItemName %></td>
    <td align="center"><%= oiMall.FItemList(i).FRegdate %><br><%= oiMall.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oiMall.FItemList(i).FLTiMallRegdate %><br><%= oiMall.FItemList(i).FLTiMallLastUpdate %></td>
    <td align="right">
        <% if oiMall.FItemList(i).FSaleYn="Y" then %>
        <strike><%= FormatNumber(oiMall.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %></font>
        <% else %>
        <%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %>
        <% end if %>
    </td>
    <td align="center">
        <% if oiMall.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oiMall.FItemList(i).IsSoldOut then %>
            <% if oiMall.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td align="center">
	<%
		If oiMall.FItemList(i).FItemdiv = "06" OR oiMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
    <td align="center">
    <% if (oiMall.FItemList(i).FLTiMallStatCd>0) then %>
    <% if Not IsNULL(oiMall.FItemList(i).FLTiMallPrice) then %>
        <% if (oiMall.FItemList(i).Fsellcash<>oiMall.FItemList(i).FLTiMallPrice) then %>
        <strong><%= formatNumber(oiMall.FItemList(i).FLTiMallPrice,0) %></strong>
        <% else %>
        <%= formatNumber(oiMall.FItemList(i).FLTiMallPrice,0) %>
        <% end if %>
        <br>
        <% if (oiMall.FItemList(i).FLTiMallSellYn="X" or oiMall.FItemList(i).FLTiMallSellYn="N") then %><a href="javascript:checkNdelReged('<%=oiMall.FItemList(i).FItemID%>');"><% end if %>
        <% if (oiMall.FItemList(i).FSellyn<>oiMall.FItemList(i).FLTiMallSellYn) then %>
        <strong><%= oiMall.FItemList(i).FLTiMallSellYn %></strong>
        <% else %>
        <%= oiMall.FItemList(i).FLTiMallSellYn %>
        <% end if %>
        <% if (oiMall.FItemList(i).FLTiMallSellYn="X" or oiMall.FItemList(i).FLTiMallSellYn="N") then %></a><% end if %>
    <% end if %>
    <% end if %>
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
					Response.Write oiMall.FItemList(i).getLotteItemStatCd & "<br>(" & oiMall.FItemList(i).FLtiMallTmpGoodNo & ")"
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
    <% if oiMall.FItemList(i).FCateMapCnt>0 then %>
	    매칭됨
    <% else %>
    	<font color="darkred">매칭안됨</font>
    <% end if %>

    <% if (oiMall.FItemList(i).FaccFailCNT>0) then %>
        <br><font color="red" title="<%= oiMall.FItemList(i).FlastErrStr %>">ERR:<%= oiMall.FItemList(i).FaccFailCNT %></font>
    <% end if %>
    </td>
    <td align="center"><%= oiMall.FItemList(i).FinfoDiv %>
    <% if (oiMall.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oiMall.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oiMall.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">옵션금액</font>
	    <% If oiMall.FItemList(i).FoptAddPrcRegType<>0 Then %>
	    (<%=oiMall.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% end if %>
    </td>

</tr>
<% next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
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
<% if C_ADMIN_AUTH then %>
<tr>
    <td colspan="16" align="center" bgcolor="#FFFFFF">
    <%= ltiMallAuthNo %>
    </td>
</tr>
<% end if %>
</table>
</form>
<form name="frmDel" method="post" action="Lotteitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="600"></iframe>
<% set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->