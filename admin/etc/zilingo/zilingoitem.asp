<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/zilingo/zilingocls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, zilingoGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, deliverytype, mwdiv
Dim expensive10x10, diffPrc, zilingoYes10x10No, zilingoNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research, exctrans
Dim oZilingo

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
zilingoGoodNo			= request("zilingoGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
zilingoYes10x10No		= request("zilingoYes10x10No")
zilingoNo10x10Yes		= request("zilingoNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
exctrans				= requestCheckVar(request("exctrans"), 1)

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
'	ExtNotReg = "D"
	MatchCate = ""
	bestOrd = "on"
	sellyn = "Y"
	isReged = "A"
End If

'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
'zilingo ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If zilingoGoodNo <> "" then
	Dim iA2, arrTemp2, arrzilingoGoodNo
	zilingoGoodNo = replace(zilingoGoodNo,",",chr(10))
	zilingoGoodNo = replace(zilingoGoodNo,chr(13),"")
	arrTemp2 = Split(zilingoGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrzilingoGoodNo = arrzilingoGoodNo & trim("'"&arrTemp2(iA2)&"'") & ","
		End If
		iA2 = iA2 + 1
	Loop
	zilingoGoodNo = left(arrzilingoGoodNo,len(arrzilingoGoodNo)-1)
End If

Set oZilingo = new CZilingo
	oZilingo.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oZilingo.FPageSize					= 50
Else
	oZilingo.FPageSize					= 20
End If
	oZilingo.FRectCDL					= request("cdl")
	oZilingo.FRectCDM					= request("cdm")
	oZilingo.FRectCDS					= request("cds")
	oZilingo.FRectItemID				= itemid
	oZilingo.FRectItemName				= itemname
	oZilingo.FRectSellYn				= sellyn
	oZilingo.FRectLimitYn				= limityn
	oZilingo.FRectSailYn				= sailyn
	oZilingo.FRectonlyValidMargin		= onlyValidMargin
	oZilingo.FRectMakerid				= makerid
	oZilingo.FRectzilingoGoodNo			= zilingoGoodNo
	oZilingo.FRectMatchCate				= MatchCate
	oZilingo.FRectIsMadeHand			= isMadeHand
	oZilingo.FRectIsOption				= isOption
	oZilingo.FRectIsReged				= isReged
	oZilingo.FRectDeliverytype			= deliverytype
	oZilingo.FRectMwdiv					= mwdiv

	oZilingo.FRectExtNotReg				= ExtNotReg
	oZilingo.FRectExpensive10x10		= expensive10x10
	oZilingo.FRectdiffPrc				= diffPrc
	oZilingo.FRectZilingoYes10x10No		= zilingoYes10x10No
	oZilingo.FRectZilingoNo10x10Yes		= zilingoNo10x10Yes
	oZilingo.FRectExtSellYn				= extsellyn
	oZilingo.FRectFailCntOverExcept		= ""
	oZilingo.FRectFailCntExists			= failCntExists
	oZilingo.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oZilingo.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oZilingo.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oZilingo.getZilingoreqExpireItemList
Else
	oZilingo.getZilingoRegedItemList		'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
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
				comp.form.MatchCate.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="zilingoYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="zilingoNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.zilingoYes10x10No.checked){
            comp.form.zilingoYes10x10No.checked = false;
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
	if ((comp.name!="zilingoYes10x10No")&&(frm.zilingoYes10x10No.checked)){ frm.zilingoYes10x10No.checked=false }
	if ((comp.name!="zilingoNo10x10Yes")&&(frm.zilingoNo10x10Yes.checked)){ frm.zilingoNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//��Ͽ��� ���� Reset
function ckeckReset(){
	document.frm.ExtNotReg.disabled = false;
	document.frm.wReset.checked=false;
	document.getElementById("AR").checked=false;
	document.getElementById("NR").checked=false;
	document.getElementById("RR").checked=false;
	document.getElementById("QR").checked=false;
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
	if ((comp.name!="zilingoYes10x10No")&&(frm.zilingoYes10x10No.checked)){ frm.zilingoYes10x10No.checked=false }
	if ((comp.name!="zilingoNo10x10Yes")&&(frm.zilingoNo10x10Yes.checked)){ frm.zilingoNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogNewitemList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// TEST SubCategory
function zilingoCategoryProcess() {
    if (confirm('������ ī�װ� Ȯ��?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SubCategory";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function zilingoSelectRegProcess() {
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��zilingo���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ��ǰ ��ȸ
function checkzilingoItemConfirm() {
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ��� ��ȸ
function checkzilingoQuantityConfirm() {
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� �����ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKQUANTITY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function zilingoEditProcess(){
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n��zilingo���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
	}
}

// ���õ� ��ǰ ���� �ϰ� ����
function zilingoPriceEditProcess(){
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ������ ���� �Ͻðڽ��ϱ�?\n\n��zilingo���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ��� �ϰ� ����
function zilingoEditQtyProcess(){
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ��� ���� �Ͻðڽ��ϱ�?\n\n��zilingo���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "QTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ȸ �� ��� �ϰ� ����
function zilingoSelectEditQtyProcess(){
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

    if (confirm('zilingo�� �����Ͻ� ' + chkSel + '�� ��� ���� �Ͻðڽ��ϱ�?\n\n��zilingo���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDITQTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �Ǹſ��� ����
function zilingoSellYnProcess(chkYn) {
	var chkSel=0;
	var strSell;
	var strcmdparam;
	try {
		if(frmSvArr.cksel.length>1) {
			for(var i=0;i<frmSvArr.cksel.length;i++) {
				if(frmSvArr.cksel[i].checked) chkSel++;
			}
		} else {
			if(frmSvArr.cksel.checked) chkSel++;
		}
		if(chkSel<=0) {
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}
	if(chkYn == "Y"){
		strSell = "�Ǹ���";
		strcmdparam = "ONSALE";
	}else if(chkYn == "N"){
		strSell = "������";
		strcmdparam = "SOLDOUT";
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��ZILINGO���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/zilingo/actzilingoReq.asp"
		document.frmSvArr.submit();
    }
}
// zilingo ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/zilingo/popzilingoCateList.asp","popzilingo","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=ZILINGO&ml=EN','itemWeightEdit','width=1024,height=768,scrollbars=yes,resizable=yes')
	popwin.focus();
}
function PopAttributes(iitemid, iitemoption, icatekey){
	var popwin = window.open('/admin/etc/zilingo/popAttribute.asp?itemid=' + iitemid +'&itemoption='+iitemoption+'&catekey='+icatekey,'itemWeightEdit','width=1024,height=500,scrollbars=yes,resizable=yes')
	popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="https://sellers.zilingo.com" target="_blank">zilingo_Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then
				response.write "<font color='GREEN'>[ csglobal@10x10.co.kr | xpsqkdlxps1! ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		zilingo ��ǰ�ڵ� : <textarea rows="2" cols="20" name="zilingoGoodNo" id="itemid"><%=Replace(replace(zilingoGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >zilingo ��Ͻ���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >zilingo �ݷ�
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >zilingo ��Ͽ����̻�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >zilingo ���ο���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >zilingo ��ϿϷ�(����)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">��ü</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">�̵��</label>&nbsp;
		<label><input type="radio" id="RR" name="isReged" <%= ChkIIF(isReged="R","checked","") %> onClick="checkisReged(this)" value="R">ǰ��ó�����</label>
		<label><input type="radio" id="QR" name="isReged" <%= ChkIIF(isReged="Q","checked","") %> onClick="checkisReged(this)" value="Q">��ϻ�ǰ �ǸŰ���</label>
		<label><input type="radio" name="wReset" onclick="ckeckReset(this);">��Ͽ�������Reset</label>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!-- #include virtual="/admin/etc/incsearch1.asp"-->
		ī�װ�
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>zilingo ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="zilingoYes10x10No" <%= ChkIIF(zilingoYes10x10No="on","checked","") %> ><font color=red>zilingo�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="zilingoNo10x10Yes" <%= ChkIIF(zilingoNo10x10Yes="on","checked","") %> ><font color=red>zilingoǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="zilingoitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="SubCategory" onclick="zilingoCategoryProcess();">&nbsp;&nbsp;
				<% End If %>
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('zilingo');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ī�װ�" onclick="pop_CateManager();">
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
				������ǰ ��� :
				<input class="button" type="button" id="btnRegSel" value="���" onClick="zilingoSelectRegProcess();">
				<br /><br />
				������ǰ �˻� :
				<input class="button" type="button" id="btnSelectGoodNo" value="��ǰ" onClick="checkzilingoItemConfirm();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnSelectQty" value="���" onClick="checkzilingoQuantityConfirm();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="zilingoPriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnQty" value="���" onClick="zilingoSelectEditQtyProcess();">
<!--
				�Ʒ� ���� ����Ǵ� ����
				<input class="button" type="button" id="btnQty" value="���" onClick="zilingoEditQtyProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEdit" value="����" onClick="zilingoEditProcess();">
				&nbsp;&nbsp;
			</td>
-->
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">������</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="zilingoSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="20">
		�˻���� : <b><%= FormatNumber(oZilingo.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oZilingo.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">zilingo�����<br>zilingo����������</td>
	<td width="90">���ǸŰ�<br /><font color='BLUE'>�Ǹŵɰ���</font></td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">��ǰ<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">zilingo<br>����<!--���Ǹ� --></td>
	<td width="70">zilingo���</td>
	<td width="100">zilingo<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">����</td>
	<td width="70">Attribute</td>
</tr>
<% For i=0 to oZilingo.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oZilingo.FItemList(i).FItemID %>_<%= oZilingo.FItemList(i).FItemoption %>"></td>
	<td><img src="<%= oZilingo.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oZilingo.FItemList(i).FItemID%>" target="_blank"><%= oZilingo.FItemList(i).FItemID %></a>
		<br /><%= oZilingo.FItemList(i).FItemoption %>
		<% If oZilingo.FItemList(i).FzilingoStatcd <> 7 Then %>
		<br><%= oZilingo.FItemList(i).getzilingoStatName %>
		<% End If %>
		<%= oZilingo.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left">
		<%= oZilingo.FItemList(i).FMakerid %> <%= oZilingo.FItemList(i).getDeliverytypeName %>
		<br /><%= oZilingo.FItemList(i).FItemName %>
		<% If NOT isnull(oZilingo.FItemList(i).FOptionName) Then %>
		<br /><%= oZilingo.FItemList(i).FOptionName %>
		<% End If %>
		<br /><font color="BLUE"><%= oZilingo.FItemList(i).FChgitemname %>
		<br /><%= oZilingo.FItemList(i).FChgOptionname %></font>
	</td>
	<td align="center"><%= oZilingo.FItemList(i).FRegdate %><br><%= oZilingo.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oZilingo.FItemList(i).FzilingoRegdate %><br><%= oZilingo.FItemList(i).FzilingoLastUpdate %></td>
	<td align="right">
		<%= FormatNumber(oZilingo.FItemList(i).FOrgprice,0) %><br>
		<%
			If oZilingo.FItemList(i).FMaySellPrice <> "" Then
				response.write "<font color='BLUE'>"&formatNumber(oZilingo.FItemList(i).FMaySellPrice,0)&"</font>"
			End If
		%>
	</td>
	<td align="center">
	<%
		If oZilingo.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-oZilingo.FItemList(i).Fbuycash / oZilingo.FItemList(i).Fsellcash*100*100)/100 & "%" &" <br>"
		End If
	%>
	</td>
	<td align="center"><%= FormatNumber((oZilingo.FItemList(i).FitemWeight/1000),3) %>kg</td>
	<td align="center">
	<%
		If oZilingo.FItemList(i).IsOptionSoldOut Then
			If oZilingo.FItemList(i).FOptSellyn = "N" Then
	%>
		<font color="red">ǰ��</font>
	<%
			Else
	%>
		<font color="red">�Ͻ�<br>ǰ��</font>
	<%
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oZilingo.FItemList(i).FzilingoStatCd > 0) Then
			If Not IsNULL(oZilingo.FItemList(i).FzilingoPrice) Then
				If (oZilingo.FItemList(i).FOrgprice <> oZilingo.FItemList(i).FRegOrgprice) Then
	%>
					<strong><%= CDBL(formatNumber(oZilingo.FItemList(i).FzilingoPrice,2)) %></strong><br>
	<%
				Else
					response.write CDBL(formatNumber(oZilingo.FItemList(i).FzilingoPrice,2))&"<br>"
				End If

'				If (oZilingo.FItemList(i).FSellyn="Y" and oZilingo.FItemList(i).FzilingoSellYn<>"Y") or (oZilingo.FItemList(i).FSellyn<>"Y" and oZilingo.FItemList(i).FzilingoSellYn="Y") Then
'					response.write "<strong>" & oZilingo.FItemList(i).FzilingoSellYn & "</strong>"
'				Else
'					response.write oZilingo.FItemList(i).FzilingoSellYn
'				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oZilingo.FItemList(i).FQuantity)) Then
			response.write oZilingo.FItemList(i).FQuantity
		End If
	%>
	</td>

	<td align="center">
	<%
		If Not(IsNULL(oZilingo.FItemList(i).FzilingoGoodNo)) Then
			Response.Write "<a target='_blank' href='https://zilingo.com/en-sg/product/details/"&oZilingo.FItemList(i).FzilingoGoodNo&"'>"&oZilingo.FItemList(i).FzilingoGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oZilingo.FItemList(i).Freguserid %></td>
	<td align="center">
	<%
		If oZilingo.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
		If (oZilingo.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oZilingo.FItemList(i).FlastErrStr) &"'>ERR:"& oZilingo.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
	<td><input type="button" class="button" value="����" onclick="PopItemContent('<%=oZilingo.FItemList(i).FItemid%>')"></td>
	<td>
	<% If oZilingo.FItemList(i).FAttributes = "" Then %>
		<input type="button" class="button" value="�Է�" style=color:red;font-weight:bold onclick="PopAttributes('<%=oZilingo.FItemList(i).FItemid%>', '<%=oZilingo.FItemList(i).FItemoption%>', '<%= oZilingo.FItemList(i).FCateKey%>')">
	<% Else %>
		<input type="button" class="button" value="����" style=color:blue;font-weight:bold onclick="PopAttributes('<%=oZilingo.FItemList(i).FItemid%>', '<%=oZilingo.FItemList(i).FItemoption%>', '<%= oZilingo.FItemList(i).FCateKey%>')">
	<% End If %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oZilingo.HasPreScroll then %>
		<a href="javascript:goPage('<%= oZilingo.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oZilingo.StartScrollPage to oZilingo.FScrollCount + oZilingo.StartScrollPage - 1 %>
    		<% if i>oZilingo.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oZilingo.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->