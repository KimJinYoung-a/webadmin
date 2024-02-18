<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, homeplusGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, dftMatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid
Dim expensive10x10, diffPrc, homeplusYes10x10No, homeplusNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oHomeplus

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
homeplusGoodNo			= request("homeplusGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
dftMatchCate			= request("dftMatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
homeplusYes10x10No		= request("homeplusYes10x10No")
homeplusNo10x10Yes		= request("homeplusNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")

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
'Homeplus 상품코드 엔터키로 검색되게
If homeplusGoodNo <> "" then
	Dim iA2, arrTemp2, arrhomeplusGoodNo
	homeplusGoodNo = replace(homeplusGoodNo,",",chr(10))
	homeplusGoodNo = replace(homeplusGoodNo,chr(13),"")
	arrTemp2 = Split(homeplusGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrhomeplusGoodNo = arrhomeplusGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	homeplusGoodNo = left(arrhomeplusGoodNo,len(arrhomeplusGoodNo)-1)
End If

SET oHomeplus = new CHomeplus
	oHomeplus.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oHomeplus.FPageSize					= 50
Else
	oHomeplus.FPageSize					= 20
End If
	oHomeplus.FRectCDL					= request("cdl")
	oHomeplus.FRectCDM					= request("cdm")
	oHomeplus.FRectCDS					= request("cds")
	oHomeplus.FRectItemID				= itemid
	oHomeplus.FRectItemName				= itemname
	oHomeplus.FRectSellYn				= sellyn
	oHomeplus.FRectLimitYn				= limityn
	oHomeplus.FRectSailYn				= sailyn
	oHomeplus.FRectonlyValidMargin		= onlyValidMargin
	oHomeplus.FRectMakerid				= makerid
	oHomeplus.FRectHomeplusGoodNo		= homeplusGoodNo
	oHomeplus.FRectMatchCate			= MatchCate
	oHomeplus.FRectDftMatchCate			= dftMatchCate
	oHomeplus.FRectIsMadeHand			= isMadeHand
	oHomeplus.FRectIsOption				= isOption
	oHomeplus.FRectIsReged				= isReged
	oHomeplus.FRectNotinmakerid			= notinmakerid

	oHomeplus.FRectExtNotReg			= ExtNotReg
	oHomeplus.FRectExpensive10x10		= expensive10x10
	oHomeplus.FRectdiffPrc				= diffPrc
	oHomeplus.FRectHomeplusYes10x10No	= homeplusYes10x10No
	oHomeplus.FRectHomeplusNo10x10Yes	= homeplusNo10x10Yes
	oHomeplus.FRectExtSellYn			= extsellyn
	oHomeplus.FRectInfoDiv				= infoDiv
	oHomeplus.FRectFailCntOverExcept	= ""
	oHomeplus.FRectFailCntExists		= failCntExists
	oHomeplus.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oHomeplus.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHomeplus.FRectOrdType = "BM"
End If

If isReged = "R" Then						'품절처리요망 상품보기 리스트
	oHomeplus.getHomeplusreqExpireItemList
Else
	oHomeplus.getHomeplusRegedItemList			'그 외 리스트
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// 등록제외 브랜드
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=homeplus","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// 등록제외 상품
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=homeplus','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//카테고리 관리
function pop_prdDivManager() {
	var pCM1 = window.open("/admin/etc/homeplus/pophomeplusprdDivList.asp","popCatehomeplus","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM1.focus();
}
//카테고리 관리
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/homeplus/pophomepluscateList.asp","popCatehomeplusmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
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
	if ((comp.name!="homeplusYes10x10No")&&(frm.homeplusYes10x10No.checked)){ frm.homeplusYes10x10No.checked=false }
	if ((comp.name!="homeplusNo10x10Yes")&&(frm.homeplusNo10x10Yes.checked)){ frm.homeplusNo10x10Yes.checked=false }
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

    if ((comp.name=="homeplusYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="homeplusNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.homeplusYes10x10No.checked){
            comp.form.homeplusYes10x10No.checked = false;
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
	if ((comp.name!="homeplusYes10x10No")&&(frm.homeplusYes10x10No.checked)){ frm.homeplusYes10x10No.checked=false }
	if ((comp.name!="homeplusNo10x10Yes")&&(frm.homeplusNo10x10Yes.checked)){ frm.homeplusNo10x10Yes.checked=false }
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
function homeplusCateAPI() {
    if (confirm('홈플러스 카테고리API를 실행하시겠습니까?\n기존에 등록된 카테고리가 삭제될 수 있습니다.')){
    	document.getElementById("btncate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CategoryView";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
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
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
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
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
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
		document.frmSvArr.cmdparam.value = "ITEMNAME";
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
		document.frmSvArr.submit();
    }
}

// 선택된 상품정보외 수정
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
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
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
        document.frmSvArr.cmdparam.value = "EditImg";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// 선택된 상품 정보 조회
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
        document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
//옵션 수 팝업
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=homeplus&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
		<a href="https://bos.homeplus.co.kr:446/LoginForm.jsp" target="_blank">Homeplus Admin바로가기</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  292811 | tenbyten10*$ ]</font>"
			End If
		%>
		<br>
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		Homeplus 상품코드 : <textarea rows="2" cols="20" name="homeplusGoodNo" id="itemid"><%=replace(homeplusGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		등록여부 :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Homeplus 등록실패
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >Homeplus 전송시도중 오류
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >Homeplus 등록예정이상
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Homeplus 등록완료(전시)
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
		전시카테고리
		<select name="MatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >미매칭
		</select>&nbsp;

		기준카테고리
		<select name="dftMatchCate" class="select">
			<option value="">전체
			<option value="Y" <%= CHkIIF(dftMatchCate="Y","selected","") %> >매칭
			<option value="N" <%= CHkIIF(dftMatchCate="N","selected","") %> >미매칭
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus 가격<텐바이텐 판매가</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >오전관리</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="homeplusYes10x10No" <%= ChkIIF(homeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus판매중&텐바이텐품절</font>상품보기</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="homeplusNo10x10Yes" <%= ChkIIF(homeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplus품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="homeplusitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('homeplus');">&nbsp;&nbsp;
				<font color="RED">우측 선작업 필요! :</font>
				<input class="button" type="button" value="기준 및 전시카테고리" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="전문 카테고리" onclick="pop_CateManager();">
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
				실제상품 등록 : <input class="button" type="button" id="btnRegSel" value="등록" onClick="HomeplusSelectRegProcess();">
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
<!-- 액션 끝 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= FormatNumber(oHomeplus.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHomeplus.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">이미지</td>
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
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
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHomeplus.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHomeplus.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHomeplus.FItemList(i).FItemID %>','homeplus','<%=oHomeplus.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oHomeplus.FItemList(i).FItemID%>" target="_blank"><%= oHomeplus.FItemList(i).FItemID %></a>
		<% If oHomeplus.FItemList(i).FHomeplusStatCd <> 7 Then %>
		<br><%= oHomeplus.FItemList(i).getHomeplusItemStatCd %>
		<% End If %>
		<% If oHomeplus.FItemList(i).FLimitYn= "Y" Then %><br><%= oHomeplus.FItemList(i).getLimitHtmlStr %></font><% End If %>
	</td>
	<td align="left"><%= oHomeplus.FItemList(i).FMakerid %><%= oHomeplus.FItemList(i).getDeliverytypeName %><br><%= oHomeplus.FItemList(i).FItemName %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FRegdate %><br><%= oHomeplus.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FHomeplusRegdate %><br><%= oHomeplus.FItemList(i).FHomeplusLastUpdate %></td>
	<td align="right">
	<% If oHomeplus.FItemList(i).FSaleYn="Y" Then %>
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
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
