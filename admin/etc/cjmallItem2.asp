<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/cjmall2/cjmallitemcls.asp"-->
<!-- #include virtual="/admin/etc/cjmall2/incCJMallFunction.asp"-->
<%
Dim makerid, cjmallitemid, itemname, itemid, eventid, ExtNotReg, bestOrd, bestOrdMall, MatchCate, sellyn, limityn, sailyn, onlyValidMargin, showminusmargin, MatchPrddiv
Dim optAddprcExists, optAddprcExistsExcept
Dim cjSell10x10Soldout, expensive10x10, cjshowminusmagin
Dim failCntExists, optExists, diffPrc, optnotExists, isMadeHand
Dim page, i, research, infodiv, reqExpire, extsellyn
page    				= request("page")
itemid  				= request("itemid")

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

makerid					= request("makerid")
eventid					= request("eventid")
itemname				= request("itemname")
cjmallitemid			= request("cjmallitemid")
ExtNotReg				= request("ExtNotReg")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
showminusmargin			= request("showminusmargin")
research				= request("research")
failCntExists			= request("failCntExists")
infodiv					= request("infodiv")
optAddprcExists 		= request("optAddprcExists")
optAddprcExistsExcept	= request("optAddprcExistsExcept")
optExists   			= request("optExists")
optnotExists   			= request("optnotExists")
isMadeHand				= request("isMadeHand")
cjSell10x10Soldout      = request("cjSell10x10Soldout")
cjshowminusmagin		= request("cjshowminusmagin")
expensive10x10          = request("expensive10x10")
reqExpire				= request("reqExpire")
extsellyn				= request("extsellyn")
diffPrc					= request("diffPrc")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"

''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
	onlyValidMargin = "on"  ''on
	bestOrd = "on"
	sellyn = "Y"
End If

Dim cjMall
Set cjMall = new CCjmall
	cjMall.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	cjMall.FPageSize					= 50
Else
	cjMall.FPageSize					= 20
End If
	cjMall.FRectMakerid					= makerid
	cjMall.FRectItemName				= itemname
	cjMall.FRectCJMallPrdNo				= cjmallitemid
	cjMall.FRectCDL 					= request("cdl")
	cjMall.FRectCDM 					= request("cdm")
	cjMall.FRectCDS 					= request("cds")
	cjMall.FRectItemID					= itemid
	cjMall.FRectEventid					= eventid
	cjMall.FRectExtNotReg				= ExtNotReg
	cjMall.FRectMatchCate				= MatchCate
	cjMall.FRectPrdDivMatch				= MatchPrddiv
	cjMall.FRectSellYn					= sellyn
	cjMall.FRectLimitYn					= limityn
	cjMall.FRectSailYn					= sailyn
	cjMall.FRectonlyValidMargin 		= onlyValidMargin
	cjMall.FRectMinusMargin 			= showminusmargin
	cjMall.FRectFailCntExists			= failCntExists
	cjMall.Finfodiv						= infodiv
	cjMall.FRectdiffPrc 				= diffPrc
	cjMall.FRectoptAddprcExists			= optAddprcExists
	cjMall.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	cjMall.FRectoptExists				= optExists
	cjMall.FRectoptnotExists			= optnotExists
	cjMall.FRectisMadeHand				= isMadeHand
	cjMall.FRectCjSell10x10Soldout      = CjSell10x10Soldout
	cjMall.FRectCjshowminusmagin		= cjshowminusmagin
	cjMall.FRectExpensive10x10          = expensive10x10
	cjMall.FRectExtSellYn  				= extsellyn
	If (bestOrd="on") Then
	    cjMall.FRectOrdType = "B"
	ElseIf (bestOrdMall="on") Then
	    cjMall.FRectOrdType = "BM"
	End If

	If reqExpire <> "" Then
	    cjMall.getCjmallreqExpireItemList
	Else
	    cjMall.GetCjmallRegedItemList
	End If

If (session("ssBctID")="kjy8517") Then
	Dim JYSQL, qq, ww
	ww = ""
	JYSQL = ""
	JYSQL = JYSQL & " select TOP 100 itemid from "
	JYSQL = JYSQL & " db_outmall.dbo.tbl_item "
	JYSQL = JYSQL & " where itemid in ( "
	JYSQL = JYSQL & " 	select itemid from [db_outmall].[dbo].tbl_OutMall_regedoption where mallid = 'cjmall' and itemoption = '0000' and outmallsellyn = 'N' and lastupdate < '2013-10-07' group by itemid "
	JYSQL = JYSQL & " ) "
	JYSQL = JYSQL & " and isusing = 'Y' "
	JYSQL = JYSQL & " and sellyn = 'Y' "
	JYSQL = JYSQL & " and itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='cjmall') "
	JYSQL = JYSQL & " and makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='cjmall') "
	JYSQL = JYSQL & " and itemid not in (147391,147393,230552,359218,359220,368619,394415,399605,483146,483153,496532,552817,598861,598862,598867,598880,598883,598889,625614,642212,643690,662015,662022,662027,672830,673715,691890,695865,704132,742944,743918,745879,755202,755203,755204,755205,755206,755207,755227,755230,755232,755247,763029,763495,771158,771160,795016,795021,795022,795025,795026,795468,800146,802496,802497,802499,805812,818832,818833,818834,818835,819415,819434,819436,819446,819452,819457,819459,819462,819465,819469,821876,826952,827518,827528,846214,848259,848260,848794,848887,850412,850413,850592,850593,850594,858009,858395,858654,858658,858661,858667,858684,858689,858702,859978,859979,859980,859982,864586,864609,864612,864614,864619,864620,864621,864623,864628,864629,864670,867525,877189,877342,877343,877345,880644,882650,882652,883089,884608,884609,888093,892290,892992,893000,897417,898864,899538,899539,899540,899541,899542,899543,899544,899545,899546,899547,899548,899549,899550,899551,899553,899555,899556,899557,899558,899561,899562,899563,899605,900836,910593,913629,916273,916274,916324,916330,918488,918491,918492,918493,918495,918496,918497,918498) "			'조건배송이면서 10000원 미만상품
	JYSQL = JYSQL & " group by itemid "
	JYSQL = JYSQL & " ORDER by itemid ASC "
'	rsCTget.Open JYSQL, dbCTget
'	If Not(rsCTget.EOF or rsCTget.BOF) Then
'		For qq = 1 to rsCTget.RecordCount
'			if qq = rsCTget.RecordCount Then
'				ww = ww & trim(rsCTget("itemid"))
'			Else
'				ww = ww & trim(rsCTget("itemid"))&","
'			End If
'			rsCTget.MoveNext
'		Next
'	End If
'	rsCTget.Close
End If

%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}


function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=cjmall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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

    if ((comp.name=="cjSell10x10Soldout")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }

        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "A";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.cjSell10x10Soldout.checked){
            comp.form.cjSell10x10Soldout.checked = false;
        }

        comp.form.ExtNotReg.value = "D";
        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "Y";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
        comp.form.onlyValidMargin.checked=false;
    }

    if ((comp.name=="cjshowminusmagin")&&(comp.checked)){

        comp.form.ExtNotReg.value = "D";
        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "Y";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
        comp.form.onlyValidMargin.checked=false;
    }

	if ((comp.name=="reqExpire")&&(comp.checked)){
	    comp.form.ExtNotReg.value="D";
	    comp.form.MatchCate.value="Y";
	    comp.form.sellyn.value="A";
	    comp.form.limityn.value="";
	    comp.form.onlyValidMargin.checked=false;
	}
	if ((comp.name!="cjshowminusmagin")&&(frm.cjshowminusmagin.checked)){ frm.cjshowminusmagin.checked=false }
	if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
	if ((comp.name=="diffPrc")){frm.onlyValidMargin.checked=true;}
}

// 등록제외 브랜드
function NotInMakerid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=cjmall','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=cjmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM = window.open("/admin/etc/cjmall2/popcjmallCateList.asp","popCateMancjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//상품분류 관리
function pop_prdDivManager() {
	var pCM2 = window.open("/admin/etc/cjmall2/popcjmallprdDivList.asp","popprdDivcjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// 선택된 상품 일괄 등록
function CjregIMSI(isreg) {
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
		if (confirm('CjMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 하시겠습니까?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}else{
		if (confirm('CjMall에 선택하신 ' + chkSel + '개 상품을 예정 등록 삭제 하시겠습니까?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}
}

// 선택된 상품 일괄 등록
function CjSelectRegProcess(isreal) {
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
		if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelect";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
        }
	}
}

// 선택된 상품 일괄 수정
function CjSelectEditProcess() {
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

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 정보+단품 일괄 수정
function CjSelectEdit2Process() {
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

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect2";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}


function CjSelectPriceEditProcess() {
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

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품 가격을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
       // document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

function CjSelectPriceEditProcess2() {
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

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품 가격을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        //document.getElementById("btnPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect2";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

//선택상품 수량변경
function CjSelectQTYEditProcess(){
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

    if (confirm('CJMall에 선택하신 ' + chkSel + '개 상품 가격을 수정 하시겠습니까?\n\n※CJMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditqty").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditQty";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}


// 선택된 상품 판매여부 변경
function CjmallSellYnProcess(chkYn) {
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
		case "N": strSell="일시중단";break;
		case "X": strSell="판매종료(삭제)";break;
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※cjmall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '로 변경하면 cjmall에서 수정 불가/등록목록에서 삭제되며 재판매시  차후 새로 등록하셔야 합니다. 계속 하시겠습니까?')) return;
        }

        //document.getElementById("btnSellYn").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 단품, 판매상태 수정
function CjSelectSaleStatEditProcess() {
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

	if (confirm('CjMall에 선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※CjMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditDanpum").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdSaleDTSel";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}

//선택상품 예약수정
function CjSelectDateEditProcess() {
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

	if (confirm('CjMall에 선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※CjMall과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditDate").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdDateSel";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}

//선택상품 승인확인 및 판매상태 check - batch
function batchStatCheck(){
    document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItemAuto";
	document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}


//선택상품 승인확인 및 판매상태 check
function checkCjItemConfirm(comp) {
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

	//if (confirm('선택 상품 승인여부 및 판매상태 조회 하시겠습니까?')){
		//comp.disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "confirmItem";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	//}
}

function CjSellynSubmit(yn){

	if(yn == true){
	    if (document.getElementById('theday').value.length!=10) {
	        alert('날짜 형식으로 입력해 주세요.yyyy-mm-dd');
	        return;
	     }
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST&sday="+document.getElementById('theday').value+"";
	}else{
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday="+document.getElementById('sday').value+"";
	}
	document.getElementById("btnSell1").disabled=true;
	document.getElementById("btnSell2").disabled=true;
}


function NumObj(obj){
	if (event.keyCode >= 48 && event.keyCode <= 57) { //숫자키만 입력
		return true;
	} else {
		event.returnValue = false;
	}
}
function popCjCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=cjmallCommonCode&CommCD="+ccd+"";
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td class="a">
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		Cjmall상품번호: <input type="text" name="cjmallitemid" value="<%= cjmallitemid %>" size="9" maxlength="9" class="input">
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
		&nbsp;
		<a href="http://partner.cjmall.com/login.jsp" target="_blank">CJ몰Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 411378 | store10x10 | 1010cube* ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품번호: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
   		이벤트번호: <input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
    	<br>
    	등록여부 :
		<select name="ExtNotReg" class="select">
			<option value="">전체
			<option value="M" <%= CHkIIF(ExtNotReg="M","selected","") %> >CJmall 미등록(등록가능)
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >CJmall 등록실패
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >CJmall 등록예정이상
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >CJmall 등록예정
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >CJmall 전송시도중오류
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >CJmall 등록후 승인대기(임시)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >CJmall 등록완료(전시)
			<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >CJmall 수정요망
		</select>&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴)</b>&nbsp;
		카테매칭 :
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
		상품분류매칭 :
		<select name="MatchPrddiv" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >미매칭
		</select>&nbsp;
		판매여부 :
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >전체
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
		</select>&nbsp;
		한정여부 :
		<select name="limityn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >한정
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >일반
		</select>&nbsp;
		세일여부 :
		<select name="sailyn" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >세일Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >세일N
		</select>&nbsp;
		품목정보 :
		<select name="infodiv" class="select">
			<option value="" <%= CHkIIF(infoDiv="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >입력
			<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >미입력
			<option value="01" <%= CHkIIF(infodiv="01","selected","") %> >01
			<option value="02" <%= CHkIIF(infodiv="02","selected","") %> >02
			<option value="03" <%= CHkIIF(infodiv="03","selected","") %> >03
			<option value="04" <%= CHkIIF(infodiv="04","selected","") %> >04
			<option value="05" <%= CHkIIF(infodiv="05","selected","") %> >05
			<option value="06" <%= CHkIIF(infodiv="06","selected","") %> >06
			<option value="07" <%= CHkIIF(infodiv="07","selected","") %> >07
			<option value="08" <%= CHkIIF(infodiv="08","selected","") %> >08
			<option value="09" <%= CHkIIF(infodiv="09","selected","") %> >09
			<option value="10" <%= CHkIIF(infodiv="10","selected","") %> >10
			<option value="11" <%= CHkIIF(infodiv="11","selected","") %> >11
			<option value="35" <%= CHkIIF(infodiv="12","selected","") %> >35
		</select>&nbsp;
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >마진 <%= CMAXMARGIN %>%이상 상품만 보기
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >등록수정오류상품
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품
		&nbsp;
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품제외
		&nbsp;
		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >옵션존재상품
		&nbsp;
		<input type="checkbox" name="optnotExists" <%= ChkIIF(optnotExists="on","checked","") %> >단품상품(옵션=0)
		&nbsp;
		주문제작여부 :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		<input type="checkbox" name="cjshowminusmagin"  <%= ChkIIF(cjshowminusmagin="on","checked","") %> onClick="checkComp(this)"  ><font color=red>역마진</font>상품보기 (MaxMagin : <%= CMAXMARGIN %>%) (CJ 판매중)
		&nbsp;
		<input type="checkbox" name="cjSell10x10Soldout" <%= ChkIIF(cjSell10x10Soldout="on","checked","") %> onClick="checkComp(this)"><font color=red>CJ판매중&텐바이텐품절</font>상품보기
		&nbsp;
		<input type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> onClick="checkComp(this)"><font color=red>CJ 가격<텐바이텐 판매가</font>상품보기
		&nbsp;
		<input onClick="checkComp(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
		<br>
		<input onClick="checkComp(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>품절처리요망</font>상품보기 (제휴몰 사용안함등)
		&nbsp;&nbsp;제휴판매상태 :
		<select name="extsellyn" class="select">
		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >종료
		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >종료제외
		</select>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>

<br>
<form name="frmReg" method="post" action="cjmallItem.asp" style="margin:0px;">
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
				<font color="RED">우측 2개 선작업 필요! :</font>
				<input class="button" type="button" value="cjMall상품분류매칭" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="cjMall카테고리매칭" onclick="pop_CateManager();">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="선택상품 실제 등록" onClick="CjSelectRegProcess(true);">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="선택상품 정보 수정" onClick="CjSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="선택상품 단품가격 기준 수정" onClick="CjSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditqty" value="선택상품 수량 수정" onClick="CjSelectQTYEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDanpum" value="선택 단품 수정" onClick="CjSelectSaleStatEditProcess();">
				<input class="button" type="button" id="btnEditDate" value="선택상품 정보+단품 수정" onClick="CjSelectEdit2Process();">
				<% If C_ADMIN_AUTH or (session("ssBctID")="kjy8517") or (session("ssBctID")="cogusdk") or (session("ssBctID")="areum531") or (session("ssBctID")="therthis") or (session("ssBctID")="joohyun49") then %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input class="button" type="button" id="btnPrice" value="선택상품 판매가격 기준 수정" onClick="CjSelectPriceEditProcess2();">
				<% End If %>
				<!--
				<input class="button" type="button" id="btnEditDate" value="선택상품 예약수정" onClick="CjSelectDateEditProcess();">
				-->
				<br><br>
				예정여부 가공 :
				<input class="button" type="button" id="btnRegSel" value="선택상품 예정 등록" onClick="CjregIMSI(true);">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegSel" value="선택상품 예정 삭제" onClick="CjregIMSI(false);" >
				<br><br>
				승인여부 검색 :
				<!--
				<input type="text" name="theday" value="" size="10" maxlength="10">
				<input class="button" type="button" id="btnSell1" value="특정날짜 승인여부 확인" onClick="CjSellynSubmit(true);">&nbsp;&nbsp;
				<select name="sday" class="select" id="sday">
				<% For i = 0 to 9 %>
					<option value="<%=i%>"><%=i%>
				<% Next %>
				</select>일전&nbsp;
				<input class="button" type="button" id="btnSell2" value="일정 기간 승인여부 확인" onClick="CjSellynSubmit(false);" >
				-->
				<input class="button" type="button"  value="선택상품 승인(판매상태) 확인" onClick="checkCjItemConfirm(this);" >
				<br><br>
				공통코드 검색 :
				<select name="CommCD" class="select" id="CommCD">
					<option value="L126">택배사코드
					<option value="6009">리드타임
					<option value="8047">가등록채널구분
				</select>
				<input class="button" type="button" id="btnCommcd" value="공통코드확인" onClick="popCjCommCDSubmit();" >
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">일시중단</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="CjmallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input type="button" value="판매상태Check(관리자)" onClick="batchStatCheck();">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(cjMall.FTotalPage,0) %> 총건수: <%= FormatNumber(cjMall.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">CJMall등록일<br>CJMall최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">CJMall<br>가격및판매</td>
	<td width="70">CJMall<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="60">상품분류<br>매칭여부</td>
</tr>
<% For i = 0 To cjMall.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= cjMall.FItemList(i).FItemID %>"></td>
	<td><img src="<%= cjMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= cjMall.FItemList(i).FItemID %>','cjMall','<%=cjMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center"><a href="<%=wwwURL%>/<%=cjMall.FItemList(i).FItemID%>" target="_blank"><%= cjMall.FItemList(i).FItemID %></a><br><%= cjMall.FItemList(i).getcjmallStatName %></td>
	<td><%= cjMall.FItemList(i).FMakerid %><%= cjMall.FItemList(i).getDeliverytypeName %><br><%= cjMall.FItemList(i).FItemName %></td>
	<td align="center"><%= cjMall.FItemList(i).FRegdate %><br><%= cjMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= cjMall.FItemList(i).FcjmallRegdate %><br><%= cjMall.FItemList(i).FcjmallLastUpdate %></td>
	<td align="right">
	<% If cjMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(cjMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(cjMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(cjMall.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-cjMall.FItemList(i).Fbuycash/cjMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).IsSoldOut Then
			If cjMall.FItemList(i).FSellyn = "N" Then
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
		If cjMall.FItemList(i).FItemdiv = "06" OR cjMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (cjMall.FItemList(i).FcjmallStatCd > 0) Then
			If Not IsNULL(cjMall.FItemList(i).FcjmallPrice) Then
				If (cjMall.FItemList(i).Fsellcash <> cjMall.FItemList(i).FcjmallPrice) Then
	%>
					<strong><%= formatNumber(cjMall.FItemList(i).FcjmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(cjMall.FItemList(i).FcjmallPrice,0)&"<br>"
				End If

				If (cjMall.FItemList(i).FSellyn="Y" and cjMall.FItemList(i).FcjmallSellYn<>"Y") or (cjMall.FItemList(i).FSellyn<>"Y" and cjMall.FItemList(i).FcjmallSellYn="Y") Then
	%>
					<strong><%= cjMall.FItemList(i).FcjmallSellYn %></strong>
	<%
				Else
					response.write cjMall.FItemList(i).FcjmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(cjMall.FItemList(i).FcjmallPrdNo)) Then
			Response.Write "<a target='_blank' href='http://www.cjmall.com/prd/detail_cate.jsp?item_cd="&cjMall.FItemList(i).FcjmallPrdNo&"'>"&cjMall.FItemList(i).FcjmallPrdNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= cjMall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=cjMall.FItemList(i).FItemID%>','0');"><%= cjMall.FItemList(i).FoptionCnt %>:<%= cjMall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= cjMall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If cjMall.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).Fcddkey <> "" Then
			response.write "매칭됨("&cjMall.FItemList(i).Finfodiv&")"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If

		If (cjMall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& cjMall.FItemList(i).FlastErrStr &"'>ERR:"& cjMall.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="20">
	<td colspan="17" align="center" bgcolor="#FFFFFF">
	<% If cjMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= cjMall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + cjMall.StartScrollPage To cjMall.FScrollCount + cjMall.StartScrollPage - 1 %>
		<% If i>cjMall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If cjMall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set cjMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->