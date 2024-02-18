<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, coupangGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, morningJY, deliverytype, mwdiv, GosiEqual, MatchShipping
Dim expensive10x10, diffPrc, coupangYes10x10No, coupangNo10x10Yes, reqEdit, reqExpire, failCntExists, notinmakerid, priceOption
Dim page, i, research, j
Dim oCoupang, splitMetaname, changMetaname, splitCoupangGosi, changeCoupangInfoDiv
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
isMadeHand				= request("isMadeHand")
isOption				= request("isOption")

infoDiv					= request("infoDiv")
morningJY				= request("morningJY")
extsellyn				= request("extsellyn")
coupangGoodNo			= request("coupangGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchShipping			= request("MatchShipping")
GosiEqual				= request("GosiEqual")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
coupangYes10x10No		= request("coupangYes10x10No")
coupangNo10x10Yes		= request("coupangNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''기본조건 등록예정이상
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchShipping = ""
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

'쿠팡 상품코드 엔터키로 검색되게
If coupangGoodNo <> "" then
	Dim iA2, arrTemp2, arrcoupangGoodNo
	coupangGoodNo = replace(coupangGoodNo,",",chr(10))
	coupangGoodNo = replace(coupangGoodNo,chr(13),"")
	arrTemp2 = Split(coupangGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2) 
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrcoupangGoodNo = arrcoupangGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	coupangGoodNo = left(arrcoupangGoodNo,len(arrcoupangGoodNo)-1)
End If

Set oCoupang = new CCoupang
	oCoupang.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oCoupang.FPageSize					= 50
Else
	oCoupang.FPageSize					= 20
End If
	oCoupang.FRectCDL					= request("cdl")
	oCoupang.FRectCDM					= request("cdm")
	oCoupang.FRectCDS					= request("cds")
	oCoupang.FRectItemID				= itemid
	oCoupang.FRectItemName				= itemname
	oCoupang.FRectSellYn				= sellyn
	oCoupang.FRectLimitYn				= limityn
	oCoupang.FRectSailYn				= sailyn
	oCoupang.FRectonlyValidMargin		= onlyValidMargin
	oCoupang.FRectMakerid				= makerid
	oCoupang.FRectCoupangGoodNo			= coupangGoodNo
	oCoupang.FRectMatchCate				= MatchCate
	oCoupang.FRectMatchShipping			= MatchShipping
	oCoupang.FRectGosiEqual				= GosiEqual
	oCoupang.FRectIsMadeHand			= isMadeHand
	oCoupang.FRectIsOption				= isOption
	oCoupang.FRectIsReged				= isReged
	oCoupang.FRectNotinmakerid			= notinmakerid
	oCoupang.FRectPriceOption			= priceOption
	oCoupang.FRectDeliverytype			= deliverytype
	oCoupang.FRectMwdiv					= mwdiv

	oCoupang.FRectExtNotReg				= ExtNotReg
	oCoupang.FRectExpensive10x10		= expensive10x10
	oCoupang.FRectdiffPrc				= diffPrc
	oCoupang.FRectCoupangYes10x10No		= coupangYes10x10No
	oCoupang.FRectCoupangNo10x10Yes		= coupangNo10x10Yes
	oCoupang.FRectExtSellYn				= extsellyn
	oCoupang.FRectInfoDiv				= infoDiv
	oCoupang.FRectFailCntOverExcept		= ""
	oCoupang.FRectFailCntExists			= failCntExists
	oCoupang.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oCoupang.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oCoupang.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oCoupang.getCoupangreqExpireItemList
Else
	oCoupang.getCoupangRegedItemList		'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=coupang","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=coupang','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="coupangYes10x10No")&&(frm.coupangYes10x10No.checked)){ frm.coupangYes10x10No.checked=false }
	if ((comp.name!="coupangNo10x10Yes")&&(frm.coupangNo10x10Yes.checked)){ frm.coupangNo10x10Yes.checked=false }
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

    if ((comp.name=="coupangYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="coupangNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.coupangYes10x10No.checked){
            comp.form.coupangYes10x10No.checked = false;
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
	if ((comp.name!="coupangYes10x10No")&&(frm.coupangYes10x10No.checked)){ frm.coupangYes10x10No.checked=false }
	if ((comp.name!="coupangNo10x10Yes")&&(frm.coupangNo10x10Yes.checked)){ frm.coupangNo10x10Yes.checked=false }
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
//브랜드코드 택배사 / 반품지코드 관리
function pop_brandDeliver(){
	var pCM4 = window.open("/admin/etc/coupang/popCoupangBrandDeliveryList.asp","popbrandDelivergsshop","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM4.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/coupang/popCoupangCateList.asp","popCateCoupangmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_coupang.asp?itemid="+iitemid+'&mallid=coupang&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
// 선택된 상품 등록
function CoupangSelectRegProcess() {
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

    if (confirm('Coupang에 선택하신 ' + chkSel + '개 상품을 등록 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 조회
function CoupangSelectViewProcess() {
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

    if (confirm('Coupang에 선택하신 ' + chkSel + '개 상품조회 하시겠습니까?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 상태 수정
function coupangSellYnProcess(chkYn) {
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
		case "X": strSell="삭제";break;
	}

	if (confirm('선택하신 ' + chkSel + '개 상품의 판매여부를 "' + strSell + '"(으)로 수정 하시겠습니까?')){
        if (chkYn=="X"){
			document.frmSvArr.cmdparam.value = "DELETE";
        }else{
        	document.frmSvArr.cmdparam.value = "EditSellYn";	
        }
		document.frmSvArr.target = "xLink";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 수정
function CoupangSelectEditProcess() {
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

	if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?')){		
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 가격 수정
function CoupangSelectPriceProcess() {
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

	if (confirm('선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?')){		
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 재고 수정
function CoupangSelectQuantityProcess() {
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

	if (confirm('선택하신 ' + chkSel + '개 상품 재고를 일괄 수정 하시겠습니까?')){		
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "QTY";
		document.frmSvArr.action = "/admin/etc/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
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
		<a href="https://wing.coupang.com" target="_blank">쿠팡Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 10x10 | store10x10 ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		쿠팡 상품코드 : <textarea rows="2" cols="20" name="coupangGoodNo" id="itemid"><%=replace(coupangGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 : 
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >쿠팡 등록성공_승인대기
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >쿠팡 전송시도 중 오류
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >쿠팡 반려
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >쿠팡 등록예정
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >쿠팡 등록완료(전시)
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
		카테고리
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
		고시일치
		<select name="GosiEqual" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(GosiEqual="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(GosiEqual="N","selected","") %> >미매칭
		</select>&nbsp;
		출고지
		<select name="MatchShipping" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchShipping="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchShipping="N","selected","") %> >미매칭
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>쿠팡 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="coupangYes10x10No" <%= ChkIIF(coupangYes10x10No="on","checked","") %> ><font color=red>쿠팡판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="coupangNo10x10Yes" <%= ChkIIF(coupangNo10x10Yes="on","checked","") %> ><font color=red>쿠팡품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="coupangItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('coupang');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="출고지" onclick="pop_brandDeliver();">&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="CoupangSelectRegProcess();">&nbsp;&nbsp;
				<br><br>
				실제상품 수정 :
				<input class="button" type="button" id="btnEditSel" value="수정" onClick="CoupangSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnPriceSel" value="가격" onClick="CoupangSelectPriceProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnQtySel" value="재고" onClick="CoupangSelectQuantityProcess();">&nbsp;&nbsp;
				<br><br>
				승인여부 조회 :
				<input class="button" type="button" id="btnViewSel" value="조회" onClick="CoupangSelectViewProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
					<option value="X">삭제</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="coupangSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- 리스트 시작 -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= FormatNumber(oCoupang.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCoupang.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">Coupang등록일<br>Coupang최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">Coupang<br>가격및판매</td>
	<td width="70">Coupang<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="60">카테고리<br>매칭여부</td>
	<td width="60">출고지<br>매칭여부</td>
	<td width="50">품목</td>
	<td width="60">쿠팡고시</td>
	<td width="150">Meta정보</td>
</tr>
<% For i=0 to oCoupang.FResultCount - 1 %>
<tr align="center" bgcolor='#FFFFFF'">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oCoupang.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oCoupang.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oCoupang.FItemList(i).FItemID %>','coupang','<%=oCoupang.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oCoupang.FItemList(i).FItemID%>" target="_blank"><%= oCoupang.FItemList(i).FItemID %></a>
		<% If oCoupang.FItemList(i).FCoupangStatcd <> 7 Then %>
		<br><%= oCoupang.FItemList(i).getCoupangStatName %>
		<% End If %>
		<%= oCoupang.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oCoupang.FItemList(i).FMakerid %> <%= oCoupang.FItemList(i).getDeliverytypeName %><br><%= oCoupang.FItemList(i).FItemName %></td>
	<td align="center"><%= oCoupang.FItemList(i).FRegdate %><br><%= oCoupang.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oCoupang.FItemList(i).FCoupangRegdate %><br><%= oCoupang.FItemList(i).FCoupangLastUpdate %></td>
	<td align="right">
		<% If oCoupang.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oCoupang.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oCoupang.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oCoupang.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-oCoupang.FItemList(i).Fbuycash/oCoupang.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).IsSoldOut Then
			If oCoupang.FItemList(i).FSellyn = "N" Then
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
		If oCoupang.FItemList(i).FItemdiv = "06" OR oCoupang.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oCoupang.FItemList(i).FCoupangStatCd > 0) Then
			If Not IsNULL(oCoupang.FItemList(i).FCoupangPrice) Then
				If (oCoupang.FItemList(i).Fsellcash <> oCoupang.FItemList(i).FCoupangPrice) Then
	%>
					<strong><%= formatNumber(oCoupang.FItemList(i).FCoupangPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oCoupang.FItemList(i).FCoupangPrice,0)&"<br>"
				End If

				If (oCoupang.FItemList(i).FSellyn="Y" and oCoupang.FItemList(i).FCoupangSellYn<>"Y") or (oCoupang.FItemList(i).FSellyn<>"Y" and oCoupang.FItemList(i).FCoupangSellYn="Y") Then
	%>
					<strong><%= oCoupang.FItemList(i).FCoupangSellYn %></strong>
	<%
				Else
					response.write oCoupang.FItemList(i).FCoupangSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oCoupang.FItemList(i).FCoupangGoodNo)) Then
			Response.Write oCoupang.FItemList(i).FCoupangGoodNo & "<br />"
		End If
		
		If Not(IsNULL(oCoupang.FItemList(i).FProductId)) Then
			Response.Write "<a target='_blank' href='http://www.coupang.com/vp/products/"&oCoupang.FItemList(i).FProductId&"'><font color='blue'>"&oCoupang.FItemList(i).FProductId&"</font></a>"
		End If
	%>
	</td>
	<td align="center"><%= oCoupang.FItemList(i).Freguserid %></td>
	<td align="center">
		<a href="javascript:popManageOptAddPrc('<%=oCoupang.FItemList(i).FItemID%>','0');"><%= oCoupang.FItemList(i).FoptionCnt %>:<%= oCoupang.FItemList(i).FregedOptCnt %></a>
	</td>
	<td align="center"><%= oCoupang.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).FOutboundShippingPlaceCode > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oCoupang.FItemList(i).FinfoDiv %>
		<%
		If (oCoupang.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oCoupang.FItemList(i).FlastErrStr) &"'>ERR:"& oCoupang.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
	<%
		changeCoupangInfoDiv = ""
		splitCoupangGosi = Split(oCoupang.FItemList(i).FMallinfoDiv, ",")
		For j = 0 to Ubound(splitCoupangGosi)
			rw oCoupang.FItemList(i).getCoupangInfoDiv(Trim(splitCoupangGosi(j)))
		Next
	%>
	</td>
	<td align="center">
	<%
		changMetaname = ""
		splitMetaname = Split(oCoupang.FItemList(i).FMetaOption, ",")
		For j = 0 to Ubound(splitMetaname)
			If instr(splitMetaname(j), "***") > 0 Then
				changMetaname = changMetaname & "<font color='red'>" & Replace(splitMetaname(j), "***", "") & "</font>,"
			Else
				changMetaname = changMetaname & splitMetaname(j) & ","
			End If
		Next
		If Right(changMetaname,1) = "," Then
			changMetaname = Left(changMetaname, Len(changMetaname) - 1)
		End If
		response.write changMetaname
	%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oCoupang.HasPreScroll then %>
		<a href="javascript:goPage('<%= oCoupang.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oCoupang.StartScrollPage to oCoupang.FScrollCount + oCoupang.StartScrollPage - 1 %>
    		<% if i>oCoupang.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oCoupang.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oCoupang = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->