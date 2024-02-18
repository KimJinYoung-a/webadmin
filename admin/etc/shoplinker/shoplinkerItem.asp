<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/incShoplinkerFunction.asp"-->
<%
Dim itemid, itemname,  mode
Dim page, makerid, ShoplinkerNotReg, sellyn, limityn
Dim delitemid, ShoplinkerGoodNo, showminusmagin, expensive10x10, ShoplinkerYes10x10No, ShoplinkerNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists, failCntExists, optAddPrcRegTypeNone
Dim bestOrd, bestOrdMall, extsellyn, infoDiv
Dim i
page    = request("page")
itemid  = request("itemid")
If itemid <> "" Then
	If Right(itemid,1) = "," OR Right(itemid,1) = " " Then
		Response.Write "<script>alert('상품코드가 잘못 입력되었습니다.');history.back();</script>"
		Response.End
	End IF
End IF

itemname				= html2db(request("itemname"))
mode					= request("mode")
makerid					= request("makerid")
ShoplinkerNotReg		= request("ShoplinkerNotReg")
sellyn					= request("sellyn")
limityn					= request("limityn")
delitemid				= request("delitemid")
ShoplinkerGoodNo		= request("ShoplinkerGoodNo")
showminusmagin			= request("showminusmagin")
expensive10x10			= request("expensive10x10")
ShoplinkerYes10x10No	= request("ShoplinkerYes10x10No")
ShoplinkerNo10x10Yes	= request("ShoplinkerNo10x10Yes")
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
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
failCntExists			= request("failCntExists")
extsellyn				= request("extsellyn")
infoDiv					= request("infoDiv")
mall_name				= request("mall_name")

If page = "" Then page = 1
If sellyn = "" Then sellyn = "Y"

''기본조건 등록예정이상
If (research="") Then
    ShoplinkerNotReg = "M"		'J
    onlyValidMargin = "on"
    bestOrd="on"
    sellyn = "Y"
End If

Dim oshoplinker
SET oshoplinker = new CShoplinker
If (ShoplinkerNotReg="F") then                       '''승인대기
	oshoplinker.FPageSize 					= 20
Else
	oshoplinker.FPageSize 					= 20
End If
	oshoplinker.FCurrPage					= page
	oshoplinker.FRectItemID					= itemid
	oshoplinker.FRectItemName				= itemname
	oshoplinker.FRectMakerid				= makerid
	oshoplinker.FRectCDL					= request("cdl")
	oshoplinker.FRectCDM					= request("cdm")
	oshoplinker.FRectCDS					= request("cds")
	oshoplinker.FRectShoplinkerNotReg		= ShoplinkerNotReg
	oshoplinker.FRectSellYn					= sellyn
	oshoplinker.FRectLimitYn				= limityn
	oshoplinker.FRectShoplinkerGoodNo		= ShoplinkerGoodNo
	oshoplinker.FRectMinusMigin				= showminusmagin
	oshoplinker.FRectExpensive10x10			= expensive10x10
	oshoplinker.FRectShoplinkerYes10x10No	= ShoplinkerYes10x10No
	oshoplinker.FRectShoplinkerNo10x10Yes	= ShoplinkerNo10x10Yes
	oshoplinker.FRectOnreginotmapping		= onreginotmapping
	oshoplinker.FRectdiffPrc				= diffPrc
	oshoplinker.FRectonlyValidMargin		= onlyValidMargin
	oshoplinker.FRectoptAddprcExists		= optAddprcExists
	oshoplinker.FRectoptAddPrcRegTypeNone	= optAddPrcRegTypeNone                         ''옵션추가금액상품 미설정 상품.
	oshoplinker.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oshoplinker.FRectoptExists				= optExists
	oshoplinker.FRectFailCntExists			= failCntExists
	oshoplinker.FRectFailCntOverExcept		= ""
	oshoplinker.FRectExtSellYn				= extsellyn
	oshoplinker.FRectInfoDiv				= infoDiv
	oshoplinker.FRectMall_name				= mall_name

If (bestOrd = "on") Then
    oshoplinker.FRectOrdType				 = "B"
ElseIf (bestOrdMall = "on") Then
    oshoplinker.FRectOrdType				= "BM"
End If

If reqExpire <> "" Then
	oshoplinker.getShoplinkerreqExpireItemList
Else
	oshoplinker.getShoplinkerRegedItemList
End If
%>
<script language="javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.shoplinkerNotReg.value="M";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="reqEdit"){
            frm.shoplinkerNotReg.value="R";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.shoplinkerNotReg.value="D";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="shoplinkerNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="shoplinkerYes10x10No")&&(frm.shoplinkerYes10x10No.checked)){ frm.shoplinkerYes10x10No.checked=false }
        if ((comp.name!="shoplinkerNo10x10Yes")&&(frm.shoplinkerNo10x10Yes.checked)){ frm.shoplinkerNo10x10Yes.checked=false }
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

// 선택된 상품 일괄 등록
function ShoplinkerSelectRegProcess(isreal) {

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

	if(document.getElementById('categbn').value == ""){
		alert('대카테고리구분을 선택하세요');
		document.getElementById('categbn').focus();
		return;
	}

    if (isreal){
		if (confirm('Shoplinker에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?\n\Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.getElementById("btnRegSelR").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelect";
			document.frmSvArr.subcmd.value = document.getElementById('categbn').value;
			document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
			document.frmSvArr.submit();
		}
	}else{
		if (confirm('Shoplinker에 선택하신 ' + chkSel + '개 상품을 수정 하시겠습니까?\n\Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
			document.getElementById("btnRegSelR2").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EDTSelect";
			document.frmSvArr.subcmd.value = document.getElementById('categbn').value;
			document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
			document.frmSvArr.submit();
		}
	}
}

function ShoplinkerSelectRegPoomOKProcess() {
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

	if (confirm('Shoplinker에 선택하신 ' + chkSel + '개 품목정보를 일괄 등록 하시겠습니까?\n\Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSelP").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "RegPoom";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}

function OutmallSelectEditProcess() {
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

	if (confirm('Shoplinker에 연동된 선택하신 ' + chkSel + '개 외부몰 정보를 수정 하시겠습니까?\n\Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditOut").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditOutMall";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}

function SelectItemCDSearch(){
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

	if (confirm('Shoplinker에 선택하신 ' + chkSel + '개 상품코드 조회 하시겠습니까?\n\Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSelS").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SearchITEM";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}
function ShoplinkerSellYnProcess(chkYn){
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

    if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?\n\n※Shoplinker와의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
    }
}
function divch(divid, itemid){
	document.frmSvArr.cmdparam.value = divid;
	document.frmSvArr.subcmd.value = itemid;
	document.frmSvArr.target="xLink";
	document.frmSvArr.action='shoplinker_Outmallsearch.asp';
	document.frmSvArr.submit();
}
function OutmallSetting(){
	var popwin=window.open('/admin/etc/shoplinker/popOutmallsetting.asp','notin','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 브랜드
function NotInMakerid(){
	var popwin=window.open('/admin/etc/shoplinker/JaehyuMall_Not_In_Makerid.asp?mallgubun=shoplinker','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// 등록제외 상품
/*
function NotInItemid(){
	var popwin=window.open('/admin/etc/shoplinker/JaehyuMall_Not_In_Itemid.asp?mallgubun=shoplinker','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
*/
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		브 랜 드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		샵링커상품번호: <input type="text" name="shoplinkerGoodNo" value="<%= shoplinkerGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;&nbsp;
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		&nbsp;
		<a href="http://ad2.shoplinker.co.kr/" target="_blank">샵링커Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 10x10 | cube1010 ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		상품번호: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		<br>
		등록여부 :
		<select name="shoplinkerNotReg" class="select">
		<option value="">전체
		<option value="M" <%= CHkIIF(shoplinkerNotReg="M","selected","") %> >샵링커 미등록(등록가능)
		<option value="Q" <%= CHkIIF(shoplinkerNotReg="Q","selected","") %> >샵링커 등록실패
		<option value="J" <%= CHkIIF(shoplinkerNotReg="J","selected","") %> >샵링커 등록시도+등록완료
		<option value="A" <%= CHkIIF(shoplinkerNotReg="A","selected","") %> >샵링커 등록시도중오류
		<option value="F" <%= CHkIIF(shoplinkerNotReg="F","selected","") %> >샵링커 등록완료(외부몰 미연결)
		<option value="D" <%= CHkIIF(shoplinkerNotReg="D","selected","") %> >샵링커 등록완료(외부몰 연결)
		<option value="R" <%= CHkIIF(shoplinkerNotReg="R","selected","") %> >샵링커 수정요망
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순(제휴)</b>
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

		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >마진 <%= CMAXMARGIN %>%이상 상품만 보기
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품
		&nbsp;
<!--	<input type="checkbox" name="optAddPrcRegTypeNone" <%= ChkIIF(optAddPrcRegTypeNone="on","checked","") %> onClick="checkComp(this)">옵션추가판매미설정상품 &nbsp; -->
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">옵션추가금액존재상품제외
		&nbsp;

		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >옵션존재상품
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >등록수정오류상품
		<br><br>
		-- Quick 검색 / 등록 / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >등록가능 상품
		<br><br>
		-- Quick 검색 / 수정 / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>역마진</font>상품보기 (MaxMagin : <%= CMAXMARGIN %>%) (샵링커 판매중)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>샵링커 가격<텐바이텐 판매가</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="shoplinkerYes10x10No" <%= ChkIIF(shoplinkerYes10x10No="on","checked","") %> ><font color=red>샵링커판매중&텐바이텐품절</font>상품보기
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="shoplinkerNo10x10Yes" <%= ChkIIF(shoplinkerNo10x10Yes="on","checked","") %> ><font color=red>샵링커품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기
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
		&nbsp;&nbsp;외부몰 :
		<% CALL DrawShoplinkerOutmall("mall_name", mall_name, true, "") %>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a">
   		대카테고리 구분 : 
   		<select name="categbn" id="categbn" class="select">
   			<option value="">--CHOICE--</option>
   			<option value="10x10">10x10(10x10)</option>
   			<option value="Sourcing">Sourcing(10x10)</option>
   			<option value="ithinkso">ithinkso(ithinkso)</option>
   		</select>&nbsp;
   		한정 5개미만이나 등록 : 
   		<input type="button" class="button" value="예외브랜드" onclick="NotInMakerid();">&nbsp;
   		<!--
   		한정 5개미만이나 등록 : 
   		<input type="button" class="button" value="예외상품" onclick="NotInItemid();">
   		-->
   		<br><br>
   		샵링커 상품 가공 :
   		<input class="button" type="button" id="btnRegSelR" value="상품등록" onClick="ShoplinkerSelectRegProcess(true);">&nbsp
   		<input class="button" type="button" id="btnRegSelR2" value="상품수정" onClick="ShoplinkerSelectRegProcess(false);">&nbsp
   		<input class="button" type="button" id="btnRegSelP" value="품목등록/수정" onClick="ShoplinkerSelectRegPoomOKProcess();">&nbsp
   		<input class="button" type="button" id="btnRegSelS" value="쇼핑몰 상품코드 조회" onClick="SelectItemCDSearch();">
   		<br><br>
   		외부몰 상품 가공 :
   		<input class="button" type="button" id="btnEditOut" value="상품수정" onClick="OutmallSelectEditProcess();">&nbsp
	</td>
	<td align="right" valign="top" class="a">
		<font Color= "BLUE">※ 우측 버튼을 클릭하여 미리 세팅해주셔야 합니다.</font>
		<input class="button" type="button" id="btnOutmallSet" value="아웃몰관리" onClick="OutmallSetting()">
		<br><br>
		선택상품을
		<Select name="chgSellYn" class="select">
			<option value="N">품절</option>
			<option value="Y">판매중</option>
		</Select>(으)로
		<input class="button" type="button" id="btnSellYn" value="변경" onClick="ShoplinkerSellYnProcess(frm.chgSellYn.value);">
		<br><br>
		<a href='/admin/etc/shoplinker/201305.xls' onfocus='this.blur()'><font color="RED">*샵링커제휴몰정리_엑셀다운</font>
    </td>
</tr>
</table>
</form>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oshoplinker.FTotalPage,0) %> 총건수: <%= FormatNumber(oshoplinker.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td >브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">샵링커 등록일<br>샵링커 최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">샵링커<br>가격및판매</td>
	<td width="120">샵링커 상품번호<br>(임시번호)</td>
	<td width="80">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">품목</td>
	<td width="80">샵링커<br>품목등록YN</td>
	<td width="80">외부몰연결YN</td>
</tr>
<%
If oshoplinker.FResultCount > 0 Then
	For i=0 to oshoplinker.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oshoplinker.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oshoplinker.FItemList(i).Fsmallimage %>" width="50"</td>
    <td align="center"><%= oshoplinker.FItemList(i).FItemID %><br><%= oshoplinker.FItemList(i).getShoplinkerItemStatCd %>
    <% If oshoplinker.FItemList(i).FLimitYn = "Y" Then %><br><%= oshoplinker.FItemList(i).getLimitHtmlStr %></font><% end if %>
    </td>
    <td><%= oshoplinker.FItemList(i).FMakerid %> <%= oshoplinker.FItemList(i).getDeliverytypeName %><br><%= oshoplinker.FItemList(i).FItemName %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FRegdate %><br><%= oshoplinker.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FShoplinkerRegdate %><br><%= oshoplinker.FItemList(i).FShoplinkerLastUpdate %></td>
    <td align="right">
        <% If oshoplinker.FItemList(i).FSaleYn = "Y" Then %>
        <strike><%= FormatNumber(oshoplinker.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oshoplinker.FItemList(i).FSellcash,0) %></font>
        <% Else %>
        <%= FormatNumber(oshoplinker.FItemList(i).FSellcash,0) %>
        <% End If %>
    </td>
    <td align="center">
        <% If oshoplinker.FItemList(i).Fsellcash <> 0 Then %>
        <%= CLng(10000-oshoplinker.FItemList(i).Fbuycash/oshoplinker.FItemList(i).Fsellcash*100*100)/100 %> %
        <% End If %>
    </td>
    <td align="center">
        <% If oshoplinker.FItemList(i).IsSoldOut Then %>
            <% If oshoplinker.FItemList(i).FSellyn = "N" Then %>
            <font color="red">품절</font>
            <% Else %>
            <font color="red">일시<br>품절</font>
            <% End if %>
        <% End If %>
    </td>
    <td align="center">
    <% If (oshoplinker.FItemList(i).FshoplinkerStatCd > 0) then %>
    <% If Not IsNULL(oshoplinker.FItemList(i).FshoplinkerPrice) Then %>
        <% If (oshoplinker.FItemList(i).Fsellcash <> oshoplinker.FItemList(i).FshoplinkerPrice) Then %>
        <strong><%= formatNumber(oshoplinker.FItemList(i).FshoplinkerPrice,0) %></strong>
        <% Else %>
        <%= formatNumber(oshoplinker.FItemList(i).FshoplinkerPrice,0) %>
        <% End If %>
        <br>
        <% If (oshoplinker.FItemList(i).FshoplinkerSellYn = "X" or oshoplinker.FItemList(i).FshoplinkerSellYn = "N") Then %><a href="javascript:checkNdelReged('<%=oshoplinker.FItemList(i).FItemID%>');"><% End If %>
        <% If (oshoplinker.FItemList(i).FSellyn<>oshoplinker.FItemList(i).FshoplinkerSellYn) Then %>
        <strong><%= oshoplinker.FItemList(i).FshoplinkerSellYn %></strong>
        <% Else %>
        <%= oshoplinker.FItemList(i).FshoplinkerSellYn %>
        <% End If %>
        <% If (oshoplinker.FItemList(i).FshoplinkerSellYn = "X" or oshoplinker.FItemList(i).FshoplinkerSellYn="N") Then %></a><% End If %>
    <% End if %>
    <% End if %>
    </td>
    <td align="center">
    <%
    	If Not(IsNULL(oshoplinker.FItemList(i).FshoplinkerGoodNo)) then
        	Response.Write oshoplinker.FItemList(i).FshoplinkerGoodNo
		End If
	%>
    </td>
    <td align="center"><%= oshoplinker.FItemList(i).FReguserid %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FoptionCnt %></td>
    <td align="center">
    	<%= oshoplinker.FItemList(i).FrctSellCNT %>
	    <% if (oshoplinker.FItemList(i).FaccFailCNT>0) then %>
	        <br><font color="red" title="<%= oshoplinker.FItemList(i).FlastErrStr %>">ERR:<%= oshoplinker.FItemList(i).FaccFailCNT %></font>
	    <% end if %>
   	</td>
    <td align="center"><%= oshoplinker.FItemList(i).FinfoDiv %>
    <% If (oshoplinker.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oshoplinker.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oshoplinker.FItemList(i).FoptAddPrcRegType <> 0,"gray","red")%>">옵션금액</font>
	    <% If oshoplinker.FItemList(i).FoptAddPrcRegType <> 0 Then %>
	    (<%=oshoplinker.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% End If %>
    </td>
    <td align="center"><%= oshoplinker.FItemList(i).FInsert_infoCD %></td>
    <td align="center">
		<% if oshoplinker.FItemList(i).FShoplinkerOutMallConnect = "Y" then %>
			<div name="div<%=i%>" id="div<%=i%>">
				<img src="/images/icon_search.jpg" onclick="javascript:divch('div<%=i%>','<%=oshoplinker.FItemList(i).FItemID%>');" style="cursor:pointer;">
			</div>
		<%
		   Else
				response.write oshoplinker.FItemList(i).FShoplinkerOutMallConnect
    	   End If
    	%>
    </td>
</tr>
<%  Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% If oshoplinker.HasPreScroll then %>
		<a href="javascript:goPage('<%= oshoplinker.StartScrollPage-1 %>');">[pre]</a>
    	<% Else %>
    		[pre]
    	<% End If %>

    	<% For i = 0 + oshoplinker.StartScrollPage to oshoplinker.FScrollCount + oshoplinker.StartScrollPage - 1 %>
    		<% If i>oshoplinker.FTotalpage Then Exit For %>
    		<% If CStr(page) = CStr(i) Then %>
    		<font color="red">[<%= i %>]</font>
    		<% Else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% End If %>
    	<% Next %>

    	<% If oshoplinker.HasNextScroll Then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% Else %>
    		[next]
    	<% End If %>
    </td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF" height="50" align="center">
    <td colspan="17">상품이 없습니다.</td>
</tr>
<% End If %>
</form>
</table>
<% Set oshoplinker = nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="600"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->