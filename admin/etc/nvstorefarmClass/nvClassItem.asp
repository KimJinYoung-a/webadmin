<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/nvstorefarmClass/nvClassCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, nvClassGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, nvClassYes10x10No, nvClassNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing
Dim page, i, research
Dim oNvclass
Dim startMargin, endMargin
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
nvClassGoodNo			= request("nvClassGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
nvClassYes10x10No		= request("nvClassYes10x10No")
nvClassNo10x10Yes		= request("nvClassNo10x10Yes")
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
isextusing				= requestCheckVar(request("isextusing"), 1)

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = ""
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
'������� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If nvClassGoodNo <> "" then
	Dim iA2, arrTemp2, arrnvClassGoodNo
	nvClassGoodNo = replace(nvClassGoodNo,",",chr(10))
	nvClassGoodNo = replace(nvClassGoodNo,chr(13),"")
	arrTemp2 = Split(nvClassGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrnvClassGoodNo = arrnvClassGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	nvClassGoodNo = left(arrnvClassGoodNo,len(arrnvClassGoodNo)-1)
End If

Set oNvclass = new CNvClass
	oNvclass.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oNvclass.FPageSize					= 100
Else
	oNvclass.FPageSize					= 50
End If
	oNvclass.FRectCDL					= request("cdl")
	oNvclass.FRectCDM					= request("cdm")
	oNvclass.FRectCDS					= request("cds")
	oNvclass.FRectItemID				= itemid
	oNvclass.FRectItemName				= itemname
	oNvclass.FRectSellYn				= sellyn
	oNvclass.FRectLimitYn				= limityn
	oNvclass.FRectSailYn				= sailyn
'	oNvclass.FRectonlyValidMargin		= onlyValidMargin
	oNvclass.FRectStartMargin			= startMargin
	oNvclass.FRectEndMargin				= endMargin
	oNvclass.FRectMakerid				= makerid
	oNvclass.FRectNvClassGoodNo			= nvClassGoodNo
	oNvclass.FRectMatchCate				= MatchCate
	oNvclass.FRectIsMadeHand			= isMadeHand
	oNvclass.FRectIsOption				= isOption
	oNvclass.FRectIsReged				= isReged
	oNvclass.FRectNotinmakerid			= notinmakerid
	oNvclass.FRectNotinitemid			= notinitemid
	oNvclass.FRectExcTrans				= exctrans
	oNvclass.FRectPriceOption			= priceOption
	oNvclass.FRectIsSpecialPrice     	= isSpecialPrice
	oNvclass.FRectDeliverytype			= deliverytype
	oNvclass.FRectMwdiv					= mwdiv
	oNvclass.FRectIsextusing			= isextusing

	oNvclass.FRectExtNotReg				= ExtNotReg
	oNvclass.FRectExpensive10x10		= expensive10x10
	oNvclass.FRectdiffPrc				= diffPrc
	oNvclass.FRectNvClassYes10x10No		= nvClassYes10x10No
	oNvclass.FRectNvClassNo10x10Yes		= nvClassNo10x10Yes
	oNvclass.FRectExtSellYn				= extsellyn
	oNvclass.FRectInfoDiv				= infoDiv
	oNvclass.FRectFailCntOverExcept		= ""
	oNvclass.FRectFailCntExists			= failCntExists
	oNvclass.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oNvclass.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oNvclass.FRectOrdType = "BM"
End If

If isReged = "R" Then								'ǰ��ó����� ��ǰ���� ����Ʈ
	oNvclass.getNvClassreqExpireItemList
Else
	oNvclass.getNvClassRegedItemList		'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
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

// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=nvstorefarmclass','popNotInItemid','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

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
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="nvClassYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="nvClassNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.nvClassYes10x10No.checked){
            comp.form.nvClassYes10x10No.checked = false;
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
	if ((comp.name!="nvClassYes10x10No")&&(frm.nvClassYes10x10No.checked)){ frm.nvClassYes10x10No.checked=false }
	if ((comp.name!="nvClassNo10x10Yes")&&(frm.nvClassNo10x10Yes.checked)){ frm.nvClassNo10x10Yes.checked=false }
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
	if ((comp.name!="nvClassYes10x10No")&&(frm.nvClassYes10x10No.checked)){ frm.nvClassYes10x10No.checked=false }
	if ((comp.name!="nvClassNo10x10Yes")&&(frm.nvClassNo10x10Yes.checked)){ frm.nvClassNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
// ���õ� ��ǰ ���
function nvClassSelectRegItemProcess() {
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGITEM";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmClass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function nvClassSelectRegProcess() {
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function nvClassSelectEDITProcess() {
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ��ȸ
function nvClassSelectItemSearchProcess(){
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ȸ
function nvClassSelectOptionSearchProcess(){
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ �ɼ��� ��ȸ �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �̹��� ���
function nvClassSelectImageRegProcess() {
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ �̹����� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "Image";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ɼ� ���
function nvClassSelectOPTRegProcess() {
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ �ɼ��� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

function nvClassSelectDelProcess(){
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

    if (confirm('������ʿ� �����Ͻ� ' + chkSel + '�� ��ǰ �ɼ��� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�ؽ�����ʰ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "DEL";
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}

function nvClassSellYnProcess(chkYn) {
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

	switch(chkYn) {
		case "Y": strSell="�Ǹ�";break;
		case "N": strSell="�Ǹ�����";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/nvstorefarmclass/actNvClassReq.asp"
        document.frmSvArr.submit();
    }
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=nvstorefarmclass&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		&nbsp;
		<a href="http://sell.storefarm.naver.com" target="_blank">������� Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenten | tenbytenstore ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		������� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="nvClassGoodNo" id="itemid"><%=replace(nvClassGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >������� ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >������� ��Ͽ����̻�
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >������� ���۽õ��߿���
			<option value="I" <%= CHkIIF(ExtNotReg="I","selected","") %> >������� �̹����� �Ϸ�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >������� ��ϿϷ�(����)
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>������� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvClassYes10x10No" <%= ChkIIF(nvClassYes10x10No="on","checked","") %> ><font color=red>��������Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="nvClassNo10x10Yes" <%= ChkIIF(nvClassNo10x10Yes="on","checked","") %> ><font color=red>�������ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>5) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
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
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();"> &nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('nvstorefarmclass');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegImgSel" value="�̹���" onClick="nvClassSelectImageRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="yhj0613") OR (session("ssBctID")="hrkang97") Then %>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnRegSel" value="��ǰ" onClick="nvClassSelectRegItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegOptSel" value="�ɼ�" onClick="nvClassSelectOPTRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelSel" value="����" onClick="nvClassSelectDelProcess();">&nbsp;&nbsp;
				<% Else %>
				<input class="button" type="button" id="btnDelSel" value="����" onClick="nvClassSelectDelProcess();">&nbsp;&nbsp;
				<% End If %>
			<% End If %>
				<input class="button" type="button" id="btnReg" value="��ǰ+�ɼ�" onClick="nvClassSelectRegProcess();" style=color:red;font-weight:bold>
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnReg" value="����" onClick="nvClassSelectEDITProcess();" style=color:blue;font-weight:bold>
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				������ǰ ��ȸ :
				<input class="button" type="button" id="btnSchitem" value="��ǰ" onClick="nvClassSelectItemSearchProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnSchOpt" value="�ɼ�" onClick="nvClassSelectOptionSearchProcess();">&nbsp;&nbsp;
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="nvClassSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="17">
		�˻���� : <b><%= FormatNumber(oNvclass.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oNvclass.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">������ʵ����<br>�����������������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�������<br>���ݹ��Ǹ�</td>
	<td width="70">�������<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">ǰ��</td>
	<td width="100">�̹���<br>���ε�</td>
</tr>
<% For i=0 to oNvclass.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oNvclass.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oNvclass.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oNvclass.FItemList(i).FItemID %>','nvstorefarmclass','<%=oNvclass.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oNvclass.FItemList(i).FItemID%>" target="_blank"><%= oNvclass.FItemList(i).FItemID %></a>
		<% If oNvclass.FItemList(i).FNvClassStatcd <> 7 Then %>
		<br><%= oNvclass.FItemList(i).getNvClassStatName %>
		<% End If %>
		<%= oNvclass.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oNvclass.FItemList(i).FMakerid %> <%= oNvclass.FItemList(i).getDeliverytypeName %><br><%= oNvclass.FItemList(i).FItemName %></td>
	<td align="center"><%= oNvclass.FItemList(i).FRegdate %><br><%= oNvclass.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oNvclass.FItemList(i).FNvClassRegdate %><br><%= oNvclass.FItemList(i).FNvClassLastUpdate %></td>
	<td align="right">
		<% If oNvclass.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oNvclass.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oNvclass.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oNvclass.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oNvclass.FItemList(i).Fsellcash = 0 Then
			'//
		elseIf (oNvclass.FItemList(i).FNvClassStatCd > 0) and Not IsNULL(oNvclass.FItemList(i).FNvClassPrice) Then
			If (oNvclass.FItemList(i).FSaleYn = "Y") and (oNvclass.FItemList(i).FSellcash < oNvclass.FItemList(i).FNvClassPrice) Then
				'// ���޸� ���� �Ǹ���
		%>
		<strike><%= CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		<font color="#CC3333"><%= CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).FNvClassPrice*100*100)/100 & "%" %></font>
		<%
			else
				response.write CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%"
			end if
		else
			response.write CLng(10000-oNvclass.FItemList(i).Fbuycash/oNvclass.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oNvclass.FItemList(i).IsSoldOut Then
			If oNvclass.FItemList(i).FSellyn = "N" Then
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
		If (oNvclass.FItemList(i).FNvClassStatCd > 0) Then
			If Not IsNULL(oNvclass.FItemList(i).FNvClassPrice) Then
				If (oNvclass.FItemList(i).Fsellcash <> oNvclass.FItemList(i).FNvClassPrice) Then
	%>
					<strong><%= formatNumber(oNvclass.FItemList(i).FNvClassPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oNvclass.FItemList(i).FNvClassPrice,0)&"<br>"
				End If

				If (oNvclass.FItemList(i).FSellyn="Y" and oNvclass.FItemList(i).FNvClassSellYn<>"Y") or (oNvclass.FItemList(i).FSellyn<>"Y" and oNvclass.FItemList(i).FNvClassSellYn="Y") Then
	%>
					<strong><%= oNvclass.FItemList(i).FNvClassSellYn %></strong>
	<%
				Else
					response.write oNvclass.FItemList(i).FNvClassSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oNvclass.FItemList(i).FNvClassGoodNo)) Then
			Response.Write "<a target='_blank' href='http://storefarm.naver.com/tenbytenclass/products/"&oNvclass.FItemList(i).FNvClassGoodNo&"'>"&oNvclass.FItemList(i).FNvClassGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oNvclass.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oNvclass.FItemList(i).FItemID%>','0');"><%= oNvclass.FItemList(i).FoptionCnt %>:<%= oNvclass.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oNvclass.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= oNvclass.FItemList(i).FinfoDiv %>
		<%
		If (oNvclass.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oNvclass.FItemList(i).FlastErrStr) &"'>ERR:"& oNvclass.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oNvclass.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oNvclass.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>�̵��</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oNvclass.HasPreScroll then %>
		<a href="javascript:goPage('<%= oNvclass.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oNvclass.StartScrollPage to oNvclass.FScrollCount + oNvclass.StartScrollPage - 1 %>
    		<% if i>oNvclass.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oNvclass.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oNvclass = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
