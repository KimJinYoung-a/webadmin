<%@ language=vbscript %>
<% option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, EzwelGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchPrddiv, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, EzwelYes10x10No, EzwelNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isextusing, cisextusing, rctsellcnt
Dim page, i, research, getRegdate
Dim oEzwel
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
EzwelGoodNo				= request("EzwelGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
EzwelYes10x10No			= request("EzwelYes10x10No")
EzwelNo10x10Yes			= request("EzwelNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
getRegdate				= request("getRegdate")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
purchasetype			= request("purchasetype")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
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
'Ezwel ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If EzwelGoodNo<>"" then
	Dim iA2, arrTemp2, arrEzwelGoodNo
	EzwelGoodNo = replace(EzwelGoodNo,",",chr(10))
	EzwelGoodNo = replace(EzwelGoodNo,chr(13),"")
	arrTemp2 = Split(EzwelGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrEzwelGoodNo = arrEzwelGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	EzwelGoodNo = left(arrEzwelGoodNo,len(arrEzwelGoodNo)-1)
End If

Set oEzwel = new CEzwel
	oEzwel.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oEzwel.FPageSize					= 100
Else
	oEzwel.FPageSize					= 50
End If
	oEzwel.FRectCDL					= request("cdl")
	oEzwel.FRectCDM					= request("cdm")
	oEzwel.FRectCDS					= request("cds")
	oEzwel.FRectItemID				= itemid
	oEzwel.FRectItemName			= itemname
	oEzwel.FRectSellYn				= sellyn
	oEzwel.FRectLimitYn				= limityn
	oEzwel.FRectSailYn				= sailyn
'	oEzwel.FRectonlyValidMargin		= onlyValidMargin
	oEzwel.FRectStartMargin			= startMargin
	oEzwel.FRectEndMargin			= endMargin
	oEzwel.FRectMakerid				= makerid
	oEzwel.FRectEzwelGoodNo			= EzwelGoodNo
	oEzwel.FRectMatchCate			= MatchCate
	oEzwel.FRectIsMadeHand			= isMadeHand
	oEzwel.FRectIsOption			= isOption
	oEzwel.FRectIsReged				= isReged
	oEzwel.FRectNotinmakerid		= notinmakerid
	oEzwel.FRectNotinitemid			= notinitemid
	oEzwel.FRectExcTrans			= exctrans
	oEzwel.FRectPriceOption			= priceOption
	oEzwel.FRectIsSpecialPrice     	= isSpecialPrice
	oEzwel.FRectDeliverytype		= deliverytype
	oEzwel.FRectMwdiv				= mwdiv
	oEzwel.FRectGetRegdate			= getRegdate
	oEzwel.FRectIsextusing			= isextusing
	oEzwel.FRectCisextusing			= cisextusing
	oEzwel.FRectRctsellcnt			= rctsellcnt

	oEzwel.FRectExtNotReg			= ExtNotReg
	oEzwel.FRectExpensive10x10		= expensive10x10
	oEzwel.FRectdiffPrc				= diffPrc
	oEzwel.FRectEzwelYes10x10No		= EzwelYes10x10No
	oEzwel.FRectEzwelNo10x10Yes		= EzwelNo10x10Yes
	oEzwel.FRectExtSellYn			= extsellyn
	oEzwel.FRectInfoDiv				= infoDiv
	oEzwel.FRectFailCntOverExcept	= ""
	oEzwel.FRectFailCntExists		= failCntExists
	oEzwel.FRectReqEdit				= reqEdit
	oEzwel.FRectPurchasetype		= purchasetype
If (bestOrd = "on") Then
    oEzwel.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oEzwel.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oEzwel.getEzwelreqExpireItemList
Else
	oEzwel.getEzwelRegedItemList			'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
//ũ�� ������Ʈ�� alert ����..2021-07-26
function systemAlert(message){
	alert(message);
}
window.addEventListener("message", (event) => {
    var data = event.data;
    if (typeof(window[data.action]) == "function") {
        window[data.action].call(null, data.message);
    } },
false);
//ũ�� ������Ʈ�� alert ����..2021-07-26 ��

// ������� �귣��
function ezwelNotInMakerid(){
	var popwin=window.open('/admin/etc/ezwel/targetMall_Not_In_Makerid.asp?mallgubun=ezwel','notin','width=900,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ��ǰ
function ezwelNotInItemid(){
	var popwin2=window.open('/admin/etc/ezwel/targetMall_Not_In_Itemid.asp?mallgubun=ezwel','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=ezwel','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_cateManager() {
	var pCM2 = window.open("/admin/etc/ezwel/popezwelcateList.asp","popCateezwelmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Newī�װ� ����
function pop_newcateManager() {
	var pCMNew = window.open("/admin/etc/ezwel/popNewezwelcateList.asp","popCateezwelmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMNew.focus();
}
//2022-11-03 ������..����ī�װ��� ��Ī
function pop_dispcateManager() {
	var pCMDisp = window.open("/admin/etc/ezwel/popDispezwelcateList.asp","popdispcateManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMDisp.focus();
}
//���θ�� ����
function pop_statcdManager() {
	var pCMMng = window.open("/admin/etc/ezwel/popstatcdList.asp","popstatcdManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCMMng.focus();
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
	if ((comp.name!="EzwelYes10x10No")&&(frm.EzwelYes10x10No.checked)){ frm.EzwelYes10x10No.checked=false }
	if ((comp.name!="EzwelNo10x10Yes")&&(frm.EzwelNo10x10Yes.checked)){ frm.EzwelNo10x10Yes.checked=false }
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

    if ((comp.name=="EzwelYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="EzwelNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.EzwelYes10x10No.checked){
            comp.form.EzwelYes10x10No.checked = false;
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
	if ((comp.name!="EzwelYes10x10No")&&(frm.EzwelYes10x10No.checked)){ frm.EzwelYes10x10No.checked=false }
	if ((comp.name!="EzwelNo10x10Yes")&&(frm.EzwelNo10x10Yes.checked)){ frm.EzwelNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}

//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
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

// ���õ� ��ǰ �ϰ� ���
function EzwelSelectRegProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��Ezwel���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        //document.frmSvArr.action = "/admin/etc/ezwel/actezwelReq.asp"
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �Ǹſ��� ����
function EzwelSellYnProcess(chkYn) {
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="�Ǹ���";break;
		case "N": strSell="ǰ��";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Ezwel���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		if (chkYn == 'Y'){
			if (confirm('[�߿�]��������� ����ڿ� �ǸŻ��� ���� �Ͽ����ϱ�?')){
		        document.frmSvArr.target = "xLink";
		        document.frmSvArr.cmdparam.value = "EditSellYn";
		        document.frmSvArr.chgSellYn.value = chkYn;
		        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
		        document.frmSvArr.submit();
			}
		}else{
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
	        document.frmSvArr.submit();
	    }
    }
}
// ���õ� ��ǰ���� ����
function EzwelSelectEditProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �Ǹſ��� ����
function EzwelSellYnNewProcess(chkYn) {
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
			alert("������ ��ǰ�� �����ϴ�.");
			return;
		}
	}
	catch(e) {
		alert("��ǰ�� �����ϴ�.");
		return;
	}

	switch(chkYn) {
		case "Y": strSell="�Ǹ���";break;
		case "N": strSell="ǰ��";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Ezwel���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		if (chkYn == 'Y'){
			if (confirm('[�߿�]��������� ����ڿ� �ǸŻ��� ���� �Ͽ����ϱ�?')){
		        document.frmSvArr.target = "xLink";
		        document.frmSvArr.cmdparam.value = "EditSellYn";
		        document.frmSvArr.chgSellYn.value = chkYn;
		        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
		        document.frmSvArr.submit();
			}
		}else{
	        document.frmSvArr.target = "xLink";
	        document.frmSvArr.cmdparam.value = "EditSellYn";
	        document.frmSvArr.chgSellYn.value = chkYn;
	        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
	        document.frmSvArr.submit();
	    }
    }
}

// ���õ� ��ǰ �ϰ� ���
function EzwelSelectNewRegProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��Ezwel���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}

function EzwelSelectNewEditProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ��ȸ
function EzwelSelectNewViewProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "/admin/etc/ezwel/actEzwelNewReq.asp"
        document.frmSvArr.submit();
    }
}


// ���õ� ��ǰ ���� ��ȸ
function EzwelSelectViewProcess() {
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

    if (confirm('Ezwel�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function EzwelSelectStatCdProcess() {
	var conStr;
	if (document.getElementById("getRegdate").value != ""){
		conStr = "Ezwel�� "+document.getElementById("getRegdate").value+"�Ͽ� ����� ��ǰ�� ���� �޾ҽ��ϱ�?";
	}else{
		conStr = "Ezwel�� ��� ��ϵ� ��ǰ�� MD���� �޾ҽ��ϱ�?";
	}

    if (confirm(conStr)){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "STAT";
        document.frmSvArr.getRegdate.value = document.getElementById("getRegdate").value;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
        document.frmSvArr.submit();
    }
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=ezwel&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
<% if request("auto") = "Y" then %>
function EzwelSelectEditProcessAuto() {
	var cnt = <%= oEzwel.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/ezwel/actezwelReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oEzwel.FResultCount %>;
	if (cnt === 0) {
		// 45�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		EzwelSelectEditProcessAuto();
		// 5�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 5*60*1000);
	}
}

$(document).ready(function() {
    $('table').hide();
});
<% end if %>

</script>
<!-- �˻� ���� -->
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
		<a href="http://shopadmin.ezwel.com" target="_blank">��������� Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ CSP������ | 10x10 | Cube1010* ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		Ezwel ��ǰ�ڵ� : <textarea rows="2" cols="20" name="EzwelGoodNo" id="itemid"><%=replace(EzwelGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Ezwel ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >Ezwel ��Ͽ����̻�
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >Ezwel ���۽õ��߿���
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >Ezwel ���ο���
			<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >Ezwel ���Ǹſ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Ezwel ��ϿϷ�(����)
		</select>&nbsp;
		<label><input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">��ü</label>&nbsp;
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Ezwel ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="EzwelYes10x10No" <%= ChkIIF(EzwelYes10x10No="on","checked","") %> ><font color=red>Ezwel�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="EzwelNo10x10Yes" <%= ChkIIF(EzwelNo10x10Yes="on","checked","") %> ><font color=red>Ezwelǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
		&nbsp;
		����� : <input id="getRegdate" name="getRegdate" value="<%= getRegdate %>" class="text" size="10" maxlength="10" />
		<img src="http://scm.10x10.co.kr/images/calicon.gif" id="gDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		(Ezwel���ο��� �˻� �� ������ǰ ������ ���ο��� ����ϴ� ��¥ �Դϴ�)
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<% if request("auto") <> "Y" then %>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, <del>�ǸŰ�(���ΰ�) 1���� �̸�</del>, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ɼǼ�(����:����) �ٸ���ǰ, �ֹ����۹�����ǰ

<p />
<% end if %>
<!-- �׼� ���� -->
<form name="frmReg" method="post" action="ezwelItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="��� ���� �귣��" onclick="ezwelNotInMakerid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="ezwelNotInItemid();">&nbsp;
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('ezwel');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<!--
				<input class="button" type="button" value="ī�װ�" onclick="pop_cateManager();">
				-->
				<input class="button" type="button" value="ī�װ�" onclick="pop_newcateManager();">
				<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="ī�װ�" onclick="pop_dispcateManager();">
				<% End If %>
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
	    		<!-- <input class="button" type="button" id="btnRegSel" value="(��)��ǰ ���" onClick="EzwelSelectRegProcess();">&nbsp;&nbsp; -->
				<input class="button" type="button" id="btnViewSel" value="���" onClick="EzwelSelectNewRegProcess();">&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<!--
			    <input class="button" type="button" id="btnEditSel" value="(��)����" onClick="EzwelSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnViewSel" value="(��)��ȸ" onClick="EzwelSelectViewProcess();">&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnViewSel" value="����" onClick="EzwelSelectNewEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnViewSel" value="��ȸ" onClick="EzwelSelectNewViewProcess();">&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" value="���θ��" onClick="pop_statcdManager();">
			   <!-- <input class="button" type="button" id="btnStat" value="����" onClick="EzwelSelectStatCdProcess();"> -->
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<!-- <option value="Y">�Ǹ���</option> -->
				</Select>(��)��
				<!-- <input class="button" type="button" id="btnSellYn" value="(��)����" onClick="EzwelSellYnProcess(frmReg.chgSellYn.value);">&nbsp;&nbsp; -->
				<input class="button" type="button" id="btnSellYn" value="����" onClick="EzwelSellYnNewProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- ����Ʈ ���� -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="getRegdate">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oEzwel.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oEzwel.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">Ezwel�����<br>Ezwel����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Ezwel<br>���ݹ��Ǹ�</td>
	<td width="70">Ezwel<br>�ǰ���</td>
	<td width="70">Ezwel<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oEzwel.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oEzwel.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oEzwel.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oEzwel.FItemList(i).FItemID %>','ezwel','')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oEzwel.FItemList(i).FItemID%>" target="_blank"><%= oEzwel.FItemList(i).FItemID %></a>
		<% If oEzwel.FItemList(i).FLimitYn= "Y" Then %><br><%= oEzwel.FItemList(i).getLimitHtmlStr %></font><% End If %>
		<%
			If oEzwel.FItemList(i).FEzwelStatcd = "3" Then
				response.write "<br />���ο���"
			ElseIf oEzwel.FItemList(i).FEzwelStatcd = "4" Then
				response.write "<br />���Ǹſ���"
			End If
		%>
	</td>
	<td align="left"><%= oEzwel.FItemList(i).FMakerid %> <%= oEzwel.FItemList(i).getDeliverytypeName %><br><%= oEzwel.FItemList(i).FItemName %></td>
	<td align="center"><%= oEzwel.FItemList(i).FRegdate %><br><%= oEzwel.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oEzwel.FItemList(i).FEzwelRegdate %><br><%= oEzwel.FItemList(i).FEzwelLastUpdate %></td>
	<td align="right">
		<% If oEzwel.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oEzwel.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oEzwel.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).Fsellcash = 0 Then
		elseif (oEzwel.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oEzwel.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oEzwel.FItemList(i).FOrgSuplycash/oEzwel.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oEzwel.FItemList(i).Fbuycash/oEzwel.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oEzwel.FItemList(i).IsSoldOut Then
			If oEzwel.FItemList(i).FSellyn = "N" Then
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
		If oEzwel.FItemList(i).FItemdiv = "06" OR oEzwel.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oEzwel.FItemList(i).FEzwelStatCd > 0) Then
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				If (oEzwel.FItemList(i).Fsellcash<>oEzwel.FItemList(i).FEzwelPrice) Then
	%>
					<strong><%= formatNumber(oEzwel.FItemList(i).FEzwelPrice,0) %></strong>
	<%
				Else
					response.write formatNumber(oEzwel.FItemList(i).FEzwelPrice,0)
				End If
	%>
				<br>
	<%
				If Not IsNULL(oEzwel.FItemList(i).FSpecialPrice) Then
					If (now() >= oEzwel.FItemList(i).FStartDate) And (now() <= oEzwel.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oEzwel.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oEzwel.FItemList(i).FSellyn="Y" and oEzwel.FItemList(i).FEzwelSellYn<>"Y") or (oEzwel.FItemList(i).FSellyn<>"Y" and oEzwel.FItemList(i).FEzwelSellYn="Y") Then
	%>
					<strong><%= oEzwel.FItemList(i).FEzwelSellYn %></strong>
	<%
				Else
					response.write oEzwel.FItemList(i).FEzwelSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<%
			If Not IsNULL(oEzwel.FItemList(i).FEzwelPrice) Then
				response.write FormatNumber(Fix(oEzwel.FItemList(i).FEzwelPrice/100)*100,0)
		 	End If
		%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oEzwel.FItemList(i).FEzwelGoodNo)) Then
		    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://shop.ezwel.com/shopNew/goods/preview/goodsDetailView.ez?preview=yes&goodsBean.goodsCd="&oEzwel.FItemList(i).FEzwelGoodNo&"')>"&oEzwel.FItemList(i).FEzwelGoodNo&"</span><br>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oEzwel.FItemList(i).FEzwelStatCd="0","(��Ͽ���)","")
		End If
	%>
	</td>
	<td align="center"><%= oEzwel.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oEzwel.FItemList(i).FItemID%>','0');"><%= oEzwel.FItemList(i).FoptionCnt %>:<%= oEzwel.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oEzwel.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oEzwel.FItemList(i).FCateMapCnt > 0 Then %>
		��Ī��
	<% Else %>
		<font color="darkred">��Ī�ȵ�</font>
	<% End If %>

	<% If (oEzwel.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oEzwel.FItemList(i).FlastErrStr %>">ERR:<%= oEzwel.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center"><%= oEzwel.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oEzwel.HasPreScroll then %>
		<a href="javascript:goPage('<%= oEzwel.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oEzwel.StartScrollPage to oEzwel.FScrollCount + oEzwel.StartScrollPage - 1 %>
    		<% if i>oEzwel.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oEzwel.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "getRegdate", trigger    : "gDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<% SET oEzwel = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
