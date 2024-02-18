<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
dim itemid, itemname, eventid, mode
dim page, makerid, LotteNotReg, MatchCate, sellyn, limityn, sailyn
dim delitemid, lotteGoodNo, showminusmagin, expensive10x10, LotteYes10x10No, LotteNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research, lotteTmpGoodNo
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists, regedOptNull,failCntExists, optAddPrcRegTypeNone, optnotExists, isMadeHand
dim bestOrd,bestOrdMall, ckLimitOver, extsellyn, infoDivYn

dim i
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

eventid  = request("eventid")
itemname= html2db(request("itemname"))
mode    = request("mode")
makerid= request("makerid")
LotteNotReg = request("LotteNotReg")
MatchCate = request("MatchCate")
sellyn = request("sellyn")
limityn = request("limityn")
sailyn = request("sailyn")
delitemid = requestCheckvar(request("delitemid"),9)
lotteGoodNo = requestCheckvar(request("lotteGoodNo"),9)
showminusmagin = request("showminusmagin")
expensive10x10 = request("expensive10x10")
LotteYes10x10No = request("LotteYes10x10No")
LotteNo10x10Yes = request("LotteNo10x10Yes")
onreginotmapping = request("onreginotmapping")
diffPrc = request("diffPrc")
onlyValidMargin = request("onlyValidMargin")
research = request("research")
lotteTmpGoodNo	= request("lotteTmpGoodNo")
reqExpire = request("reqExpire")
reqEdit   = request("reqEdit")
optAddprcExists = request("optAddprcExists")
optAddPrcRegTypeNone  = request("optAddPrcRegTypeNone")
optAddprcExistsExcept = request("optAddprcExistsExcept")
optExists   = request("optExists")
regedOptNull= request("regedOptNull")
bestOrd  = request("bestOrd")
bestOrdMall= request("bestOrdMall")
failCntExists= request("failCntExists")
ckLimitOver= request("ckLimitOver")
extsellyn   = request("extsellyn")
infoDivYn   = request("infoDivYn")
optnotExists   			= request("optnotExists")
isMadeHand					= request("isMadeHand")
if page="" then page=1



if sellyn="" then sellyn="Y"
if research="" then
    onlyValidMargin="on"
    LotteNotReg = "D"
end if


dim oLotteitem
set oLotteitem = new CLotte
oLotteitem.FPageSize       = 30
oLotteitem.FCurrPage       = page
oLotteitem.FRectItemID     = itemid
oLotteitem.FRectEventid    = eventid
oLotteitem.FRectItemName   = itemname
oLotteitem.FRectMakerid    = makerid
oLotteitem.FRectCDL = request("cdl")
oLotteitem.FRectCDM = request("cdm")
oLotteitem.FRectCDS = request("cds")
oLotteitem.FRectLotteNotReg  = LotteNotReg
oLotteitem.FRectMatchCate  = MatchCate
oLotteitem.FRectSellYn  = sellyn
oLotteitem.FRectLimitYn  = limityn
oLotteitem.FRectSailYn  = sailyn
oLotteitem.FRectLotteGoodNo  = lotteGoodNo
oLotteitem.FRectMinusMigin = showminusmagin
oLotteitem.FRectExpensive10x10 = expensive10x10
oLotteitem.FRectLotteYes10x10No = LotteYes10x10No
oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
oLotteitem.FRectOnreginotmapping = onreginotmapping
oLotteitem.FRectdiffPrc = diffPrc
oLotteitem.FRectonlyValidMargin = onlyValidMargin
oLotteitem.FRectoptAddprcExists= optAddprcExists
oLotteitem.FRectLotteTmpGoodNo		= lotteTmpGoodNo
oLotteitem.FRectoptAddPrcRegTypeNone = optAddPrcRegTypeNone                         ''옵션추가금액상품 미설정 상품.
oLotteitem.FRectoptAddprcExistsExcept= optAddprcExistsExcept
oLotteitem.FRectoptExists= optExists
oLotteitem.FRectregedOptNull= regedOptNull
oLotteitem.FRectFailCntExists = failCntExists
oLotteitem.FRectFailCntOverExcept = ""
oLotteitem.FRectExtSellYn  = extsellyn
oLotteitem.FRectInfoDivYn = infoDivYn
oLotteitem.FRectoptnotExists			= optnotExists
oLotteitem.FRectisMadeHand				= isMadeHand
if (ckLimitOver="on") then
    oLotteitem.FRectLimitOver=CStr(CMAXLIMITSELL)
end if

IF (bestOrd="on") then
    oLotteitem.FRectOrdType = "B"
ELSEIF (bestOrdMall="on") then
    oLotteitem.FRectOrdType = "BM"
end if

IF reqExpire<>"" then
    oLotteitem.getLottereqExpireItemList
ELSE
    oLotteitem.GetLotteRegedItemList
ENd IF

Dim outMallItemArr
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteCom&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

// 롯데닷컴 담당MD 목록
function pop_MDList() {
	var pMD = window.open("./lotte/popLotteMDList.asp","popMDList","width=600,height=300,scrollbars=yes,resizable=yes");
	pMD.focus();
}

// 롯데닷컴 카테고리 관리
function pop_CateManager() {
	var pCM = window.open("./lotte/popLotteCateList.asp","popCateMan","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}

// 롯데닷컴 브랜드 관리
function pop_BrandList() {
	alert("브랜드는 텐바이텐(155112)으로 고정됩니다.");
	//var pBM = window.open("./lotte/popLotteBrandMap.asp","popBrandMan","width=500,height=500,scrollbars=yes,resizable=yes");
	//pBM.focus();
}

// 미등록 상품 일괄등록
function LotteRegProcess(){
    if (confirm('롯데닷컴에 미등록된 상품을 일괄 등록 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnRegAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegAll";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 등록
function LotteSelectRegProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 재고 조회 하시겠습니까?')){
       // document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "ChkStockSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 등록예정 등록
function LotteregIMSI(isreg){
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
        if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?\n\n※30분단위로 배치 등록됩니다.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.mode.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 예정 등록 삭제 하시겠습니까?')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.mode.value = "DelSelectWait";
            document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
            document.frmSvArr.submit();
        }
    }

}

// 수정된 상품 일괄수정
function LotteEditProcess(){
    if (confirm('수정된 상품을 일괄 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 판매 상태를 확인하시겠습니까?')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "CheckItemStat";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 선택 상품 전시상품상세조회 //20130612
function LotteSelectStatWithOption(){
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 상세조회를 하시겠습니까?')){
        //document.getElementById("btnStatWithOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "StatWithOption";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditSelect" + v;
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
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
        document.getElementById("btnEditPSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditPriceSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 갱신
function LotteRealItemMappingSel() {
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
    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품 등록확인 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnChkReal2").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "getconfirmList";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}


// 선택된 상품 가격 수정
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품 명을 일괄 수정 요청 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditItemNm";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 품목 수정
function LotteSelectPoomOkEditProcess() {
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

    if (confirm('롯데닷컴에 선택하신 ' + chkSel + '개 상품 품목을 일괄 수정 요청 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnPOEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditItemPO";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// 제휴몰 아닌것 삭제
function LotteDelJaeHyuProcess(){
    //return;
    if (confirm('제휴몰판매가 아닌것을 일괄 삭제 하시겠습니까?')){
        document.getElementById("btnDelJehyu").disabled=true;

        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelJaeHyu";

        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }

}

// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteCom","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 등록제외 상품
function NotInItemid()
{
	var popwin = window.open('JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteCom','notinItem','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 등록제외 브랜드(업체 수정 불가)
function LotteNotInMakerid()
{
	var popwin = window.open('onlyLotte_Not_In_Makerid.asp','lottenotin','width=900,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 반품 주소지 검토
function ReturnCodeMappid()
{
	var popwin = window.open('JaehyuMall_ReturnCode_Mappid.asp?mallgubun=lotteCom&lotteSellyn=Y','notin','width=900,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 임시상품 일괄 갱신
function LotteRealItemMapping() {
	if(confirm("임시등록 상품이 전시상품으로 등록되었는지 일괄 확인하시겠습니까?\n\n※ 통신상태에따라 다소 시간이 걸릴 수 있습니다.")) {
		document.getElementById("btnChkReal").disabled=true;
		xLink.location.href="./lotte/actLotteCheckRDItem.asp";
	}
}

//상품판매상태,가격Check
function batchStatCheck(){
    xLink.location.href="./lotte/actRegLotteItem.asp?mode=CheckItemStatAuto";
}

//상품명 다른내역 수정.
function batchItemNmCheck(){
    xLink.location.href="./lotte/actRegLotteItem.asp?mode=CheckItemNmAuto";
}

// 선택된 상품 판매여부 변경
function LotteSellYnProcess(chkYn) {
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

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※롯데닷컴과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 롯데닷컴에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }

        document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
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
            frm.MatchCate.value="Y";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="expensive10x10"){
            frm.sellyn.value="Y";                   //판매중인거만수정
        }else{
            frm.LotteNotReg.value="D";
            frm.MatchCate.value="Y";
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

        if ((comp.name=="LotteYes10x10No")||(comp.name=="expensive10x10")){
            frm.MatchCate.value="";
        }

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

function popApiTest(){
    var popwin = window.open('/admin/etc/lotte/lotteApiTest.asp','lotteApiTest','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
/*
function popPoomOk(mallid,infodiv,itemid){
    var popwin = window.open('/admin/etc/lotte/popPoomOk.asp?mallid='+mallid+'&infoDiv='+infodiv+'&itemid='+itemid+'','popPoomOk','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
*/
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//삭제 순서 : 판매종료, 조회후종료확인후 삭제.
function delexpiredItem(iitemid){
    if (confirm('판매 종료된 상품만 삭제 하셔야 합니다.\n\n계속하시겠습니까?')){
        document.frmDel.target = "xLink";
        document.frmDel.mode.value = "DelSelectExpireItem";
        document.frmDel.delitemid.value =iitemid;
        document.frmDel.action = "/admin/etc/lotte/actRegLotteItem.asp"
        document.frmDel.submit();
    }
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		브 랜 드 :
		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		롯데닷컴상품번호:
		<input type="text" name="lotteGoodNo" value="<%= lotteGoodNo %>" size="9" maxlength="9" class="text"> &nbsp;&nbsp;
		상품명:
		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">&nbsp;&nbsp;
		롯데닷컴임시상품번호: <input type="text" name="lotteTmpGoodNo" value="<%= lotteTmpGoodNo %>" size="15" maxlength="15" class="text">
		&nbsp;
		<a href="https://partner.lotte.com/main/Login.lotte" target="_blank">롯데닷컴Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 124072 | cube101010 ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품번호: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		이벤트번호:
		<input type="text" name="eventid" value="<%= eventid %>" size="8" maxlength="6" class="text">
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
		<option value="M" <%= CHkIIF(LotteNotReg="M","selected","") %> >롯데닷컴 미등록(등록가능)
		<option value="J" <%= CHkIIF(LotteNotReg="J","selected","") %> >롯데닷컴 반려
		<option value="W" <%= CHkIIF(LotteNotReg="W","selected","") %> >롯데닷컴 등록예정
		<option value="F" <%= CHkIIF(LotteNotReg="F","selected","") %> >롯데닷컴 등록완료(임시)
		<option value="D" <%= CHkIIF(LotteNotReg="D","selected","") %> >롯데닷컴 등록완료(전시)
		<option value="R" <%= CHkIIF(LotteNotReg="R","selected","") %> >롯데닷컴 수정요망
		</select>
		&nbsp;
		    <input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>
		&nbsp;
		    <input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴)</b>
		&nbsp;
		카테고리매칭여부 :
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
		&nbsp;
		<input type="checkbox" name="ckLimitOver" <%= ChkIIF(ckLimitOver="on","checked","") %> >한정 <%=CMAXLIMITSELL%>개이상


		<br><br>
		옵션 추가 금액 존재 상품 등록 불가(2012-07-23)
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
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>롯데닷컴 가격<텐바이텐 판매가</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteYes10x10No" <%= ChkIIF(LotteYes10x10No="on","checked","") %> ><font color=red>롯데닷컴판매중&텐바이텐품절</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteNo10x10Yes" <%= ChkIIF(LotteNo10x10Yes="on","checked","") %> ><font color=red>롯데닷컴품절&텐바이텐판매가능</font>(판매중,한정>=<%=CMAXLIMITSELL%>) 상품보기
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
		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >종료제외
		</select>
		&nbsp;&nbsp;품목정보입력여부 :
		<select name="infoDivYn" class="select">
		<option value="" <%= CHkIIF(infoDivYn="","selected","") %> >전체
		<option value="Y" <%= CHkIIF(infoDivYn="Y","selected","") %> >입력
		<option value="N" <%= CHkIIF(infoDivYn="N","selected","") %> >미입력
		<option value="15" <%= CHkIIF(infoDivYn="15","selected","") %> >15
		<option value="21" <%= CHkIIF(infoDivYn="21","selected","") %> >21
		<option value="22" <%= CHkIIF(infoDivYn="22","selected","") %> >22
		<option value="23" <%= CHkIIF(infoDivYn="23","selected","") %> >23
		<option value="35" <%= CHkIIF(infoDivYn="35","selected","") %> >35
		</select>

		<!--
		<input type="checkbox" name="onreginotmapping" <%= ChkIIF(onreginotmapping="on","checked","") %> ><font color=red>롯데닷컴 등록&미패칭 카테고리</font>상품보기
		&nbsp;
		-->

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
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 브랜드(업체수정불가)" onclick="LotteNotInMakerid();"> &nbsp;
				<input class="button" type="button" id="button" value="반품 주소지 검토" onClick="ReturnCodeMappid();" >
			</td>
			<td align="right">

				<input class="button" type="button" value="Lotte담당MD" onclick="pop_MDList();"> &nbsp;
				<!--
				<input class="button" type="button" value="Lotte브랜드매칭" onclick="pop_BrandList();"> &nbsp;
			-->
				<input class="button" type="button" value="Lotte카테고리매칭" onclick="pop_CateManager();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal2" value="선택상품 등록확인" onClick="LotteRealItemMappingSel();">

				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal" value="임시상품 일괄확인 [200건씩]" onClick="LotteRealItemMapping();">

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
			    <input class="button" type="button" id="btnRegSel" value="상품 등록" onClick="LotteSelectRegProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditSel" value="상품정보/가격 수정" onClick="LotteSelectEditProcess('');">
			    &nbsp;&nbsp;
	    	    <!--
			    <input class="button" type="button" id="btnEditAll" value="수정된 상품 롯데닷컴으로 일괄수정 [10건씩]" onClick="LotteEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegAll" value="미등록 상품 롯데닷컴으로 일괄등록 [10건씩]" onClick="LotteRegProcess();">
			    -->
			    <input class="button" type="button" id="btnEditPSel" value="가격 수정" onClick="LotteSelectPriceEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="상품 재고조회" onClick="LotteSelectcheckStock();" >
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="상품명 수정요청" onClick="LotteSelectItemNmEditProcess();" >
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnPOEditSel" value="상품 품목 수정" onClick="LotteSelectPoomOkEditProcess();" >
			    <br><br>
				예정여부 가공 :
			    <!--
			    <input class="button" type="button" id="btnDelJehyu" value="제휴몰아닌것 일괄삭제 [20건씩]" onClick="LotteDelJaeHyuProcess();">
			    -->
			    <input class="button" type="button" id="btnRegSel" value="상품 등록" onClick="LotteregIMSI(true);">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="상품 삭제" onClick="LotteregIMSI(false);" >
			    &nbsp;&nbsp;
			    <br><br>
			    <input class="button" type="button" id="btnEditSel" value="선택상품정보 수정(등록대기)" onClick="LotteSelectEditProcess('2');">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnStatConfirm" value="선택상품 판매상태확인" onClick="LotteSelectStatCheck();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnStatWithOption" value="선택상품 전시상품상세조회" onClick="LotteSelectStatWithOption();">
			</td>
			<td align="right" valign="top">

				선택상품을
				<Select name="chgSellYn" class="select">
				<option value="N"  >품절</option>
				<option value="Y"  >판매중</option>
				<% if (True) or (reqExpire="on") then %>
				<option value="X" >판매종료(삭제)</option><!-- 삭제하면 이후 수정 할 수 없음 -->
				<% end if %>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="LotteSellYnProcess(frmReg.chgSellYn.value);">

				<% if (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then %>
				<br><input type="button" value="API_TEST(관리자)" onClick="popApiTest();">
				<% end if %>
				<br><br><input type="button" value="판매상태Check(관리자)" onClick="batchStatCheck();">
				&nbsp;&nbsp;<input type="button" value="상품명수정(관리자)" onClick="batchItemNmCheck();">

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
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oLotteitem.FTotalPage,0) %> 총건수: <%= FormatNumber(oLotteitem.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td >브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">롯데닷컴등록일<br>롯데닷컴최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">롯데닷컴<br>가격및판매</td>
	<td width="70">롯데닷컴<br>상품번호<br>(임시번호)</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="60">품목</td>
</tr>
<% for i=0 to oLotteitem.FResultCount - 1 %>
<%
    outMallItemArr = outMallItemArr & null2blank(oLotteitem.FItemList(i).FLotteGoodNo) & ","
    outMallItemArr = replace(outMallItemArr,",,",",")
%>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oLotteitem.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oLotteitem.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oLotteitem.FItemList(i).FItemID %>','lotteCom','<%=oLotteitem.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
    <td align="center"><%= oLotteitem.FItemList(i).FItemID %>
    <% if oLotteitem.FItemList(i).FLimitYn="Y" then %><br><font color=blue>한정:<%= oLotteitem.FItemList(i).getLimitEa %></font><% end if %>
    </td>
    <td><%= oLotteitem.FItemList(i).FMakerid %> <%= oLotteitem.FItemList(i).getDeliverytypeName %><br><%= oLotteitem.FItemList(i).FItemName %></td>
    <td align="center"><%= oLotteitem.FItemList(i).FRegdate %><br><%= oLotteitem.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oLotteitem.FItemList(i).FLotteRegdate %><br><%= oLotteitem.FItemList(i).FLotteLastUpdate %></td>
    <td align="right">
        <% if oLotteitem.FItemList(i).FSaleYn="Y" then %>
        <strike><%= FormatNumber(oLotteitem.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %></font>
        <% else %>
        <%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %>
        <% end if %>
    </td>
    <td align="center">
        <% if oLotteitem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oLotteitem.FItemList(i).IsSoldOut then %>
            <% if oLotteitem.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td align="center">
	<%
		If oLotteitem.FItemList(i).FItemdiv = "06" OR oLotteitem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
    <td align="center">
    <% if Not IsNULL(oLotteitem.FItemList(i).FLottePrice) then %>
        <% if (oLotteitem.FItemList(i).Fsellcash<>oLotteitem.FItemList(i).FLottePrice) then %>
        <strong><%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %></strong>
        <% else %>
        <%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %>
        <% end if %>
        <br>
        <% if (oLotteitem.FItemList(i).FLotteSellYn="X") then %>
        <a href="javascript:delexpiredItem('<%= oLotteitem.FItemList(i).FItemID %>');">
        <% end if %>
            <% if (oLotteitem.FItemList(i).FSellyn<>oLotteitem.FItemList(i).FLotteSellYn) then %>
            <strong><%= oLotteitem.FItemList(i).FLotteSellYn %></strong>
            <% else %>
            <%= oLotteitem.FItemList(i).FLotteSellYn %>
            <% end if %>
        <% if (oLotteitem.FItemList(i).FLotteSellYn="X") then %>
        </a>
        <% end if %>
    <% end if %>
    </td>
    <td align="center">
    <%
    	'#실상품번호
    	if Not(IsNULL(oLotteitem.FItemList(i).FLotteGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotte.com/goods/viewGoodsDetail.lotte?goods_no="&oLotteitem.FItemList(i).FLotteGoodNo&"'>"&oLotteitem.FItemList(i).FLotteGoodNo&"</a>"
		else
			'#임시상품번호
			if Not(IsNULL(oLotteitem.FItemList(i).FLotteTmpGoodNo)) then
				if oLotteitem.FItemList(i).FLotteStatCd<>"30" then
					Response.Write oLotteitem.FItemList(i).getLotteItemStatCd & "<br>(" & oLotteitem.FItemList(i).FLotteTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if
	%>
    </td>
    <td align="center"><%= oLotteitem.FItemList(i).Freguserid %></td>
    <td align="center"><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','0');"><%= oLotteitem.FItemList(i).FoptionCnt %>:<%= oLotteitem.FItemList(i).FregedOptCnt %></a></td>
    <td align="center"><%= oLotteitem.FItemList(i).FrctSellCNT %></td>
    <td align="center">
    <% if oLotteitem.FItemList(i).FCateMapCnt>0 then %>
	    매칭됨
    <% else %>
    	<font color="darkred">매칭안됨</font>
    <% end if %>

    <% if (oLotteitem.FItemList(i).FaccFailCNT>0) then %>
        <br><font color="red" title="<%= oLotteitem.FItemList(i).FlastErrStr %>">ERR:<%= oLotteitem.FItemList(i).FaccFailCNT %></font>
    <% end if %>
    </td>
    <td align="center"><%=oLotteitem.FItemList(i).FinfoDiv%>
    <% if (oLotteitem.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','1');">
    <font color="<%=CHKIIF(oLotteitem.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">옵션금액</font>
    <% if oLotteitem.FItemList(i).FoptAddPrcRegType<>0 then %>
    (<%=oLotteitem.FItemList(i).FoptAddPrcRegType%>)
    <% end if %>
    </a>
    <% end if %>
    </td>
</tr>
<% next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oLotteitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotteitem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLotteitem.StartScrollPage to oLotteitem.FScrollCount + oLotteitem.StartScrollPage - 1 %>
    		<% if i>oLotteitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLotteitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<tr height="20">
    <td colspan="16" align="left" bgcolor="#FFFFFF">
    <% if Right(outMallItemArr,1)="," then outMallItemArr=Left(outMallItemArr,Len(outMallItemArr)-1) %>
    <%= outMallItemArr %>

    <% if C_ADMIN_AUTH then %>
    <br><%=lotteAuthNo%>
    <% end if %>
    </td>
</tr>

</table>
</form>
<form name="frmDel" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="delitemid" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
<% set oLotteitem = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
