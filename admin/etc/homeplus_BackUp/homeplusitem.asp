<%@ language=vbscript %>
<%' option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim oHomeplus, research, itemname, showminusmagin, expensive10x10, diffPrc, HomeplusYes10x10No, HomeplusNo10x10Yes, extsellyn, infoDiv
Dim i, page, itemid, sellyn, makerid, HomeplusGoodNo, limityn, sailyn, optAddprcExists, optAddprcExistsExcept, optAddPrcRegTypeNone, failCntExists
Dim HomeplusNotReg, MatchCate, dftMatchCate, onlyValidMargin, optExists, reqEdit, reqExpire, regedOptNull, optnotExists, isMadeHand
Dim bestOrd, bestOrdMall
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

makerid						= request("makerid")
itemname					= html2db(request("itemname"))
HomeplusNotReg				= request("HomeplusNotReg")
MatchCate					= request("MatchCate")
dftMatchCate				= request("dftMatchCate")
onlyValidMargin				= request("onlyValidMargin")
sellyn						= request("sellyn")
limityn						= request("limityn")
sailyn						= request("sailyn")
HomeplusGoodNo				= request("HomeplusGoodNo")
research					= request("research")
optExists					= request("optExists")
optnotExists				= request("optnotExists")
bestOrd						= request("bestOrd")
bestOrdMall					= request("bestOrdMall")
optAddprcExists				= request("optAddprcExists")
optAddPrcRegTypeNone		= request("optAddPrcRegTypeNone")
regedOptNull				= request("regedOptNull")
failCntExists				= request("failCntExists")

showminusmagin				= request("showminusmagin")
expensive10x10				= request("expensive10x10")
diffPrc						= request("diffPrc")
HomeplusYes10x10No			= request("HomeplusYes10x10No")
HomeplusNo10x10Yes			= request("HomeplusNo10x10Yes")
reqEdit						= request("reqEdit")
reqExpire					= request("reqExpire")
extsellyn					= request("extsellyn")
infoDiv						= request("infoDiv")
optAddprcExistsExcept		= request("optAddprcExistsExcept")
isMadeHand					= request("isMadeHand")

If page = "" Then page = 1
If sellyn = "" Then sellyn = "Y"

''기본조건 등록예정이상
If (research="") Then
	HomeplusNotReg = "J"
	MatchCate = ""
	dftMatchCate = ""
	onlyValidMargin="on"
	sellyn="Y"
'	optAddprcExistsExcept = "on"
'	limityn="N"
End If

Set oHomeplus = new CHomeplus
	oHomeplus.FPageSize 					= 20
	oHomeplus.FCurrPage						= page
	oHomeplus.FRectCDL						= request("cdl")
	oHomeplus.FRectCDM						= request("cdm")
	oHomeplus.FRectCDS						= request("cds")
	oHomeplus.FRectItemID					= itemid
	oHomeplus.FRectItemName					= itemname
	oHomeplus.FRectSellYn					= sellyn
	oHomeplus.FRectLimitYn					= limityn
	oHomeplus.FRectSailYn					= sailyn
	oHomeplus.FRectonlyValidMargin			= onlyValidMargin
	oHomeplus.FRectMakerid					= makerid
	oHomeplus.FRectHomeplusGoodNo			= HomeplusGoodNo
	oHomeplus.FRectMatchCate				= MatchCate
	oHomeplus.FRectdftMatchCate				= dftMatchCate

	oHomeplus.FRectoptExists				= optExists
	oHomeplus.FRectoptnotExists				= optnotExists
	oHomeplus.FRectHomeplusNotReg			= HomeplusNotReg
	oHomeplus.FRectMinusMigin				= showminusmagin
	oHomeplus.FRectExpensive10x10			= expensive10x10
	oHomeplus.FRectdiffPrc					= diffPrc
	oHomeplus.FRectHomeplusYes10x10No		= HomeplusYes10x10No
	oHomeplus.FRectHomeplusNo10x10Yes		= HomeplusNo10x10Yes
	oHomeplus.FRectExtSellYn				= extsellyn
	oHomeplus.FRectInfoDiv					= infoDiv
	oHomeplus.FRectFailCntOverExcept		= ""
	oHomeplus.FRectoptAddprcExists			= optAddprcExists
	oHomeplus.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oHomeplus.FRectoptAddPrcRegTypeNone		= optAddPrcRegTypeNone                         ''옵션추가금액상품 미설정 상품.
	oHomeplus.FRectregedOptNull				= regedOptNull
	oHomeplus.FRectFailCntExists			= failCntExists
	oHomeplus.FRectisMadeHand				= isMadeHand
If (bestOrd = "on") Then
    oHomeplus.FRectOrdType					 = "B"
ElseIf (bestOrdMall = "on") Then
    oHomeplus.FRectOrdType					= "BM"
End If

If reqExpire <> "" Then
	oHomeplus.getHomeplusreqExpireItemList
Else
	oHomeplus.getHomeplusRegedItemList
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=homeplus","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin2=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=homeplus','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//카테고리 관리
function pop_prdDivManager() {
	var pCM1 = window.open("/admin/etc/homeplus/pophomeplusprdDivList.asp","popCatehomeplus","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM1.focus();
}
//카테고리 관리
function pop_cateManager() {
	var pCM2 = window.open("/admin/etc/homeplus/pophomepluscateList.asp","popCatehomeplusmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// 선택된 상품 일괄 등록
function HomeplusSelectRegProcess() {
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

    if (confirm('Homeplus에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※Homeplus와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
//      document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "RegSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품정보 수정
function HomeplusSelectEditProcess() {
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

    if (confirm('Homeplus에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※상품명, 카테고리, 정보고시, 이미지, 상품설명 등이 수정됩니다.\n\n※상품가격, 판매상태는 수정되지 않습니다.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품정보 수정
function HomeplusSelectEditItemProcess() {
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

    if (confirm('Homeplus에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\n※해당 아이템 추가/수정/판매중/판매중지/가격 정보가 변경됩니다.\n\n※상품명, 카테고리, 정보고시, 이미지, 상품설명은 수정되지 않습니다.')){
        document.getElementById("btnEditOptSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditItemSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 수정
function HomeplusSelectImgEditProcess() {
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

    if (confirm('Homeplus에 선택하신 ' + chkSel + '개 상품의 이미지를 수정 하시겠습니까?\n\n※Homeplus와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnEditImgSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditImgSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 이미지 수정
function HomeplusSelectViewProcess() {
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

    if (confirm('Homeplus에 선택하신 ' + chkSel + '개 상품정보를 조회 하시겠습니까?\n\n※Homeplus와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ViewSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 판매여부 변경
function HomeplusSellYnProcess(chkYn) {
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
	}

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Homeplus과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.HomeplusNotReg.value="M";
            frm.MatchCate.value="Y";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="reqEdit"){
            frm.HomeplusNotReg.value="R";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.HomeplusNotReg.value="D";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="HomeplusNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="HomeplusYes10x10No")&&(frm.HomeplusYes10x10No.checked)){ frm.HomeplusYes10x10No.checked=false }
        if ((comp.name!="HomeplusNo10x10Yes")&&(frm.HomeplusNo10x10Yes.checked)){ frm.HomeplusNo10x10Yes.checked=false }
        if ((comp.name!="reqREG")&&(frm.reqREG.checked)){ frm.reqREG.checked=false }
        if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
        if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
    }
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_homeplus.asp?itemid="+iitemid+'&mallid=homeplus&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
}

function homeplusCateAPI() {
    if (confirm('홈플러스 카테고리API를 실행하시겠습니까?\n기존에 등록된 카테고리가 삭제될 수 있습니다.')){
    	document.getElementById("btncate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CategoryView";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		Homeplus상품번호: <input type="text" name="HomeplusGoodNo" value="<%= HomeplusGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;&nbsp;
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		<a href="https://bos.homeplus.co.kr:446/LoginForm.jsp" target="_blank">Homeplus Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  292811 | cube1010!! ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품번호: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		주문제작여부 :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		등록여부 :
		<select name="HomeplusNotReg" class="select">
			<option value="">전체
			<option value="M" <%= CHkIIF(HomeplusNotReg="M","selected","") %> >Homeplus 미등록(등록가능)
			<option value="Q" <%= CHkIIF(HomeplusNotReg="Q","selected","") %> >Homeplus 등록실패
			<option value="J" <%= CHkIIF(HomeplusNotReg="J","selected","") %> >Homeplus 등록예정이상
			<option value="A" <%= CHkIIF(HomeplusNotReg="A","selected","") %> >Homeplus 전송시도중오류
			<option value="D" <%= CHkIIF(HomeplusNotReg="D","selected","") %> >Homeplus 등록완료(전시)
			<option value="R" <%= CHkIIF(HomeplusNotReg="R","selected","") %> >Homeplus 수정요망
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴)</b>
		&nbsp;
		전시카테매칭 :
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>
		&nbsp;
		기준카테매칭 :
		<select name="dftMatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(dftMatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(dftMatchCate="N","selected","") %> >미매칭
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
		<br><br>
		-- Quick 검색 / 수정 / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>역마진</font>상품보기 (MaxMagin : <%= CMAXMARGIN %>%) (Homeplus 판매중)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus 가격<텐바이텐 판매가</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusYes10x10No" <%= ChkIIF(HomeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus판매중&텐바이텐품절</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusNo10x10Yes" <%= ChkIIF(HomeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplus품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>품절처리요망</font>상품보기 (제휴몰 사용안함등)
		&nbsp;&nbsp;제휴판매상태 :
		<select name="extsellyn" class="select">
			<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
			<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
			<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
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
<form name="frmReg" method="post" action="homeplusItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
			<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" id="btncate" value="카테고리API" onclick="homeplusCateAPI();"> &nbsp;
			<% End If %>
				<input class="button" type="button" value="등록 제외 브랜드" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="등록 제외 상품" onclick="NotInItemid();">
			</td>
			<td align="right">
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="기준 및 전시카테고리" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="전문 카테고리" onclick="pop_cateManager();">
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
	    		<input class="button" type="button" id="btnRegSel" value="상품 등록" onClick="HomeplusSelectRegProcess();">
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSel" value="정보 수정" onClick="HomeplusSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditOptSel" value="정보외 수정" onClick="HomeplusSelectEditItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditImgSel" value="이미지 수정" onClick="HomeplusSelectImgEditProcess();">&nbsp;&nbsp;
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="cogusdk") Then %>
			    <br><br>
				실제상품 조회 :
				<input class="button" type="button" id="btnViewSel" value="정보 조회" onClick="HomeplusSelectViewProcess();">&nbsp;&nbsp;
				<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="HomeplusSellYnProcess(frmReg.chgSellYn.value);">
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
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="18" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oHomeplus.FTotalPage,0) %> 총건수: <%= FormatNumber(oHomeplus.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="80">텐바이텐<br>상품번호</td>
	<td >브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">Homeplus등록일<br>Homeplus최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Homeplus<br>가격및판매</td>
	<td width="70">Homeplus<br>상품번호</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="100">Homeplus<br>전시카테고리</td>
	<td width="100">Homeplus<br>기준카테고리</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oHomeplus.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHomeplus.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHomeplus.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHomeplus.FItemList(i).FItemID %>','Homeplus','')" style="cursor:pointer"></td>
	<td align="center"><%= oHomeplus.FItemList(i).FItemID %><br>
	<% If oHomeplus.FItemList(i).FLimitYn= "Y" Then %><%= oHomeplus.FItemList(i).getLimitHtmlStr %></font><% End If %>
	</td>
	<td><%= oHomeplus.FItemList(i).FMakerid %> <%= oHomeplus.FItemList(i).getDeliverytypeName %><br><%= oHomeplus.FItemList(i).FItemName %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FRegdate %><br><%= oHomeplus.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FHomeplusRegdate %><br><%= oHomeplus.FItemList(i).FHomeplusLastUpdate %></td>
	<td align="right">
	    <% If oHomeplus.FItemList(i).FSaleYn = "Y" Then %>
	    <strike><%= FormatNumber(oHomeplus.FItemList(i).FOrgPrice,0) %></strike><br>
	    <font color="#CC3333"><%= FormatNumber(oHomeplus.FItemList(i).FSellcash,0) %></font>
	    <% Else %>
	    <%= FormatNumber(oHomeplus.FItemList(i).FSellcash,0) %>
	    <% End If %>
	</td>
	<td align="center">
	    <% If oHomeplus.FItemList(i).Fsellcash <> 0 Then %>
	    <%= CLng(10000-oHomeplus.FItemList(i).Fbuycash/oHomeplus.FItemList(i).Fsellcash*100*100)/100 %> %
	    <% End If %>
	</td>
	<td align="center">
	    <% If oHomeplus.FItemList(i).IsSoldOut Then %>
	        <% If oHomeplus.FItemList(i).FSellyn = "N" Then %>
	        <font color="red">품절</font>
	        <% Else %>
	        <font color="red">일시<br>품절</font>
	        <% End If %>
	    <% End If %>
	</td>
	<td align="center">
	<%
		If oHomeplus.FItemList(i).FItemdiv = "06" OR oHomeplus.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<% If (oHomeplus.FItemList(i).FHomeplusStatCd > 0) Then %>
	<% If Not IsNULL(oHomeplus.FItemList(i).FHomeplusPrice) Then %>
	    <% If (oHomeplus.FItemList(i).Fsellcash<>oHomeplus.FItemList(i).FHomeplusPrice) Then %>
	    <strong><%= formatNumber(oHomeplus.FItemList(i).FHomeplusPrice,0) %></strong>
	    <% Else %>
	    <%= formatNumber(oHomeplus.FItemList(i).FHomeplusPrice,0) %>
	    <% End If %>
	    <br>
	    <% If (oHomeplus.FItemList(i).FSellyn<>oHomeplus.FItemList(i).FHomeplusSellYn) Then %>
	    <strong><%= oHomeplus.FItemList(i).FHomeplusSellYn %></strong>
	    <% Else %>
	    <%= oHomeplus.FItemList(i).FHomeplusSellYn %>
	    <% End If %>
	<% End If %>
	<% End If %>
	</td>
	<td align="center">
	<%
		'#실상품번호
		If Not(IsNULL(oHomeplus.FItemList(i).FHomeplusGoodNo)) Then
	    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://direct.homeplus.co.kr/app.product.Product.ghs?comm=usr.product.detail&i_style="&oHomeplus.FItemList(i).FHomeplusGoodNo&"')>"&oHomeplus.FItemList(i).FHomeplusGoodNo&"</span>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oHomeplus.FItemList(i).FHomeplusStatCd="0","(등록예정)","")
		End If
	%>
	</td>
	<td align="center"><%= oHomeplus.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHomeplus.FItemList(i).FItemID%>','0');"><%= oHomeplus.FItemList(i).FoptionCnt %>:<%= oHomeplus.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHomeplus.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oHomeplus.FItemList(i).FCateMapCnt > 0 Then %>
		<font color='BLUE'>매칭됨</font>
	<% Else %>
		<font color="darkred">매칭안됨</font>
	<% End If %>

	<% If (oHomeplus.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oHomeplus.FItemList(i).FlastErrStr %>">ERR:<%= oHomeplus.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oHomeplus.FItemList(i).FhDIVISION = "" Then
			response.write "<font color='darkred'>매칭안됨</font>"
		Else
			response.write "<font color='BLUE'>매칭됨</font>"
		End If
	%>
	</td>
	<td align="center"><%= oHomeplus.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oHomeplus.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHomeplus.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHomeplus.StartScrollPage to oHomeplus.FScrollCount + oHomeplus.StartScrollPage - 1 %>
    		<% if i>oHomeplus.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHomeplus.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="600"></iframe>
<% Set oHomeplus = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->