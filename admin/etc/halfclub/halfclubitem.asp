<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/halfclub/halfclubcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, halfclubGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, halfclubYes10x10No, halfclubNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice
Dim page, i, research
Dim oHalfclub
Dim startMargin, endMargin, isextusing
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
halfclubGoodNo			= request("halfclubGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
halfclubYes10x10No		= request("halfclubYes10x10No")
halfclubNo10x10Yes		= request("halfclubNo10x10Yes")
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
'	itemid = "1291678"
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

'하프클럽 상품코드 엔터키로 검색되게
If halfclubGoodNo <> "" then
	Dim iA2, arrTemp2, arrhalfclubGoodNo
	halfclubGoodNo = replace(halfclubGoodNo,",",chr(10))
	halfclubGoodNo = replace(halfclubGoodNo,chr(13),"")
	arrTemp2 = Split(halfclubGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrhalfclubGoodNo = arrhalfclubGoodNo & trim("'"&arrTemp2(iA2)&"'") & ","
		End If
		iA2 = iA2 + 1
	Loop
	halfclubGoodNo = left(arrhalfclubGoodNo,len(arrhalfclubGoodNo)-1)
End If

Set oHalfclub = new CHalfclub
	oHalfclub.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oHalfclub.FPageSize					= 100
Else
	oHalfclub.FPageSize					= 50
End If
	oHalfclub.FRectCDL					= request("cdl")
	oHalfclub.FRectCDM					= request("cdm")
	oHalfclub.FRectCDS					= request("cds")
	oHalfclub.FRectItemID				= itemid
	oHalfclub.FRectItemName				= itemname
	oHalfclub.FRectSellYn				= sellyn
	oHalfclub.FRectLimitYn				= limityn
	oHalfclub.FRectSailYn				= sailyn
	oHalfclub.FRectonlyValidMargin		= onlyValidMargin
	oHalfclub.FRectMakerid				= makerid
	oHalfclub.FRectHalfclubGoodNo		= halfclubGoodNo
	oHalfclub.FRectMatchCate			= MatchCate
	oHalfclub.FRectIsMadeHand			= isMadeHand
	oHalfclub.FRectIsOption				= isOption
	oHalfclub.FRectIsReged				= isReged
	oHalfclub.FRectNotinmakerid			= notinmakerid
	oHalfclub.FRectNotinitemid			= notinitemid
	oHalfclub.FRectExcTrans				= exctrans
	oHalfclub.FRectPriceOption			= priceOption
	oHalfclub.FRectIsSpecialPrice       = isSpecialPrice
	oHalfclub.FRectDeliverytype			= deliverytype
	oHalfclub.FRectMwdiv				= mwdiv

	oHalfclub.FRectExtNotReg			= ExtNotReg
	oHalfclub.FRectExpensive10x10		= expensive10x10
	oHalfclub.FRectdiffPrc				= diffPrc
	oHalfclub.FRectHalfClubYes10x10No	= halfclubYes10x10No
	oHalfclub.FRectHalfClubNo10x10Yes	= halfclubNo10x10Yes
	oHalfclub.FRectExtSellYn			= extsellyn
	oHalfclub.FRectInfoDiv				= infoDiv
	oHalfclub.FRectFailCntOverExcept	= ""
	oHalfclub.FRectFailCntExists		= failCntExists
	oHalfclub.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oHalfclub.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHalfclub.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oHalfclub.getHalfclubreqExpireItemList
Else
	oHalfclub.getHalfclubRegedItemList		'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=halfclub","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=halfclub','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="halfclubYes10x10No")&&(frm.halfclubYes10x10No.checked)){ frm.halfclubYes10x10No.checked=false }
	if ((comp.name!="halfclubNo10x10Yes")&&(frm.halfclubNo10x10Yes.checked)){ frm.halfclubNo10x10Yes.checked=false }
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

    if ((comp.name=="halfclubYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="halfclubNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.halfclubYes10x10No.checked){
            comp.form.halfclubYes10x10No.checked = false;
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
	if ((comp.name!="halfclubYes10x10No")&&(frm.halfclubYes10x10No.checked)){ frm.halfclubYes10x10No.checked=false }
	if ((comp.name!="halfclubNo10x10Yes")&&(frm.halfclubNo10x10Yes.checked)){ frm.halfclubNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/halfclub/pophalfclubcateList.asp","popCatehalfclubmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que 로그 리스트 팝업
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// 선택된 상품 일괄 등록
function halfclubSelectRegProcess() {
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

    if (confirm('하프클럽에 선택하신 ' + chkSel + '개 상품을 일괄 등록 하시겠습니까?\n\n※하프클럽과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}
// 선택된 상품 상태 변경
function halfclubSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
        document.frmSvArr.submit();
    }
}
// 선택된 상품 가격 수정
function halfclubriceEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품 가격을 일괄 수정 하시겠습니까?\n\n※하프클럽과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function halfclubEditProcess() {
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

    if (confirm('선택하신 ' + chkSel + '개 상품을 일괄 수정 하시겠습니까?\n\n※하프클럽과의 통신상태에 따라 시간이 다소 걸릴 수 있습니다.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품 일괄 수정
function halfclubSelectDeliveryProcess() {

    if (confirm('택배사 코드를 조회 하시겠습니까?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "Delivery";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=halfclub&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		<a href="http://scm.halfclub.com/Home/Default.aspx" target="_blank">하프클럽Admin바로가기</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then
				response.write "<font color='GREEN'>[ 10x10 | store10x10** ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		하프클럽 상품코드 : <textarea rows="2" cols="20" name="halfclubGoodNo" id="itemid"><%=replace(halfclubGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >하프클럽 등록시도
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >하프클럽 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >하프클럽 등록완료(전시)
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>하프클럽 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="halfclubYes10x10No" <%= ChkIIF(halfclubYes10x10No="on","checked","") %> ><font color=red>하프클럽판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="halfclubNo10x10Yes" <%= ChkIIF(halfclubNo10x10Yes="on","checked","") %> ><font color=red>하프클럽품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>

<p />

* 기준마진 : 제휴판매가 대비 매입가, 마진은 반올림함<br />
* 제휴판매가 : 할인가(기준마진 미만인 경우 정상가), 원단위 올림처리(하프클럽은 원단위 안씀)<br />
* 전송제외상품1 : 등록제외브랜드, 등록제외상품, 제휴몰사용안함, 업체착불, 딜상품, 꽃배달, 화물배달, 티켓(강좌) 상품, 판매가(할인가) 1만원 미만, 한정재고5개 이하, 옵션별한정재고 전부 5개 이하<br />
* 전송제외상품2 : 주문제작문구 상품<br />

<p />

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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('halfclub');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
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
				<input class="button" type="button" id="btnRegSel" value="등록" onClick="halfclubSelectRegProcess();">
				<br><br>
				실제상품 수정 :
				<!--
				<input class="button" type="button" id="btnEditPrice" value="가격" onClick="halfclubriceEditProcess();">&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnStock" value="수정" onClick="halfclubEditProcess();">&nbsp;&nbsp;
				<br><br>
				공통코드 조회 :
				<input class="button" type="button" id="btnStock" value="택배사" onClick="halfclubSelectDeliveryProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				선택상품을
				<Select name="chgSellYn" class="select">
					<option value="N">품절</option>
					<option value="Y">판매중</option>
				</Select>(으)로
				<input class="button" type="button" id="btnSellYn" value="변경" onClick="halfclubSellYnProcess(frmReg.chgSellYn.value);">
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oHalfclub.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHalfclub.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td width="140">상품등록일<br>상품최종수정일</td>
	<td width="140">하프클럽등록일<br>하프클럽최종수정일</td>
	<td width="70">텐바이텐<br>판매가</td>
	<td width="70">텐바이텐<br>마진</td>
	<td width="70">품절여부</td>
	<td width="70">주문제작<br>여부</td>
	<td width="70">하프클럽<br>가격및판매</td>
	<td width="70">하프클럽<br>상품번호</td>
	<td width="50">등록자ID</td>
	<td width="50">옵션수</td>
	<td width="50">3개월<br>판매량</td>
	<td width="80">매칭여부</td>
	<td width="80">품목</td>
</tr>
<% For i=0 to oHalfclub.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHalfclub.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHalfclub.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHalfclub.FItemList(i).FItemID %>','halfclub','<%=oHalfclub.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oHalfclub.FItemList(i).FItemID%>" target="_blank"><%= oHalfclub.FItemList(i).FItemID %></a>
		<% If oHalfclub.FItemList(i).FHalfclubStatcd <> 7 Then %>
		<br><%= oHalfclub.FItemList(i).getHalfclubStatName %>
		<% End If %>
		<%= oHalfclub.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oHalfclub.FItemList(i).FMakerid %> <%= oHalfclub.FItemList(i).getDeliverytypeName %><br><%= oHalfclub.FItemList(i).FItemName %></td>
	<td align="center"><%= oHalfclub.FItemList(i).FRegdate %><br><%= oHalfclub.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHalfclub.FItemList(i).FHalfClubRegdate %><br><%= oHalfclub.FItemList(i).FHalfClubLastUpdate %></td>
	<td align="right">
		<% If oHalfclub.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oHalfclub.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oHalfclub.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oHalfclub.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oHalfclub.FItemList(i).Fsellcash = 0 Then
			'//
		elseIf (oHalfclub.FItemList(i).FHalfclubStatCd > 0) and Not IsNULL(oHalfclub.FItemList(i).FHalfClubPrice) Then
			If (oHalfclub.FItemList(i).FSaleYn = "Y") and (oHalfclub.FItemList(i).FSellcash < oHalfclub.FItemList(i).FHalfClubPrice) Then
				'// 제휴몰 정상가 판매중
		%>
		<strike><%= CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		<font color="#CC3333"><%= CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).FHalfClubPrice*100*100)/100 & "%" %></font>
		<%
			else
				response.write CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%"
			end if
		else
			response.write CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oHalfclub.FItemList(i).IsSoldOut Then
			If oHalfclub.FItemList(i).FSellyn = "N" Then
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
		If oHalfclub.FItemList(i).FItemdiv = "06" OR oHalfclub.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>주문제작</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oHalfclub.FItemList(i).FHalfclubStatCd > 0) Then
			If Not IsNULL(oHalfclub.FItemList(i).FHalfClubPrice) Then
				If (oHalfclub.FItemList(i).Mustprice <> oHalfclub.FItemList(i).FHalfClubPrice) Then
	%>
					<strong><%= formatNumber(oHalfclub.FItemList(i).FHalfClubPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oHalfclub.FItemList(i).FHalfClubPrice,0)&"<br>"
				End If

				If Not IsNULL(oHalfclub.FItemList(i).FSpecialPrice) Then
					If (now() >= oHalfclub.FItemList(i).FStartDate) And (now() <= oHalfclub.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(특)" & formatNumber(oHalfclub.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oHalfclub.FItemList(i).FSellyn="Y" and oHalfclub.FItemList(i).FHalfClubSellYn<>"Y") or (oHalfclub.FItemList(i).FSellyn<>"Y" and oHalfclub.FItemList(i).FHalfClubSellYn="Y") Then
	%>
					<strong><%= oHalfclub.FItemList(i).FHalfClubSellYn %></strong>
	<%
				Else
					response.write oHalfclub.FItemList(i).FHalfClubSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oHalfclub.FItemList(i).FHalfclubGoodNo <> "" Then %>
			<a target="_blank" href="http://www.halfclub.com/Detail?PrStCd=<%=oHalfclub.FItemList(i).FHalfclubGoodNo%>&ColorCd=ZZ9"><%=oHalfclub.FItemList(i).FHalfclubGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oHalfclub.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHalfclub.FItemList(i).FItemID%>','0');"><%= oHalfclub.FItemList(i).FoptionCnt %>:<%= oHalfclub.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHalfclub.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oHalfclub.FItemList(i).FCateMapCnt > 0 Then
			response.write "매칭됨"
		Else
			response.write "<font color='darkred'>매칭안됨</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oHalfclub.FItemList(i).FinfoDiv %>
		<%
		If (oHalfclub.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oHalfclub.FItemList(i).FlastErrStr) &"'>ERR:"& oHalfclub.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oHalfclub.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHalfclub.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHalfclub.StartScrollPage to oHalfclub.FScrollCount + oHalfclub.StartScrollPage - 1 %>
    		<% if i>oHalfclub.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHalfclub.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>

<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
