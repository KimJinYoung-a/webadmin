<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, gsshopgoodno, extsellyn, ExtNotReg, isReged, MatchCate, MatchPrddiv, notinmakerid, priceOption, deliverytype, mwdiv
Dim expensive10x10, diffPrc, gsshopYes10x10No, gsshopNo10x10Yes, reqEdit, reqExpire, failCntExists, diffName
Dim page, i, research
Dim ogsshop
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
gsshopgoodno			= request("gsshopgoodno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
gsshopYes10x10No		= request("gsshopYes10x10No")
gsshopNo10x10Yes		= request("gsshopNo10x10Yes")
reqEdit					= request("reqEdit")
diffName				= request("diffName")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")

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
'GSShop ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If gsshopgoodno<>"" then
	Dim iA2, arrTemp2, arrgsshopgoodno
	gsshopgoodno = replace(gsshopgoodno,",",chr(10))
	gsshopgoodno = replace(gsshopgoodno,chr(13),"")
	arrTemp2 = Split(gsshopgoodno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrgsshopgoodno = arrgsshopgoodno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	gsshopgoodno = left(arrgsshopgoodno,len(arrgsshopgoodno)-1)
End If

SET oGSShop = new CGSShop
	oGSShop.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oGSShop.FPageSize					= 50
Else
	oGSShop.FPageSize					= 20
End If
	oGSShop.FRectCDL					= request("cdl")
	oGSShop.FRectCDM					= request("cdm")
	oGSShop.FRectCDS					= request("cds")
	oGSShop.FRectItemID					= itemid
	oGSShop.FRectItemName				= itemname
	oGSShop.FRectSellYn					= sellyn
	oGSShop.FRectLimitYn				= limityn
	oGSShop.FRectSailYn					= sailyn
	oGSShop.FRectonlyValidMargin		= onlyValidMargin
	oGSShop.FRectMakerid				= makerid
	oGSShop.FRectGSShopgoodno			= gsshopgoodno
	oGSShop.FRectMatchCate				= MatchCate
	oGSShop.FRectPrdDivMatch			= MatchPrddiv
	oGSShop.FRectIsMadeHand				= isMadeHand
	oGSShop.FRectIsOption				= isOption
	oGSShop.FRectIsReged				= isReged
	oGSShop.FRectNotinmakerid			= notinmakerid
	oGSShop.FRectPriceOption			= priceOption
	oGSShop.FRectDeliverytype			= deliverytype
	oGSShop.FRectMwdiv					= mwdiv

	oGSShop.FRectExtNotReg				= ExtNotReg
	oGSShop.FRectExpensive10x10			= expensive10x10
	oGSShop.FRectdiffPrc				= diffPrc
	oGSShop.FRectGSShopYes10x10No		= gsshopYes10x10No
	oGSShop.FRectGSShopNo10x10Yes		= gsshopNo10x10Yes
	oGSShop.FRectExtSellYn				= extsellyn
	oGSShop.FRectInfoDiv				= infoDiv
	oGSShop.FRectFailCntOverExcept		= ""
	oGSShop.FRectFailCntExists			= failCntExists
	oGSShop.FRectReqEdit				= reqEdit
	oGSShop.FRectDiffName				= diffName
If (bestOrd = "on") Then
    oGSShop.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oGSShop.FRectOrdType = "BM"
End If
	oGSShop.getGSShopAddOptionRegedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
//�������� �ʼ� �˾�
function pop_safecode(itemcd){
	var popwin=window.open('/admin/etc/gsshop/pop_safecode.asp?itemid='+itemcd+'','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
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
	if ((comp.name!="gsshopYes10x10No")&&(frm.gsshopYes10x10No.checked)){ frm.gsshopYes10x10No.checked=false }
	if ((comp.name!="gsshopNo10x10Yes")&&(frm.gsshopNo10x10Yes.checked)){ frm.gsshopNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
	if ((comp.name!="diffName")&&(frm.diffName.checked)){ frm.diffName.checked=false }
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "J"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "J"
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

    if ((comp.name=="gsshopYes10x10No")&&(comp.checked)){
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
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value = "N";
			comp.form.extsellyn.value = "Y";
    	}
    }

    if ((comp.name=="gsshopNo10x10Yes")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value = "Y";
			comp.form.extsellyn.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.gsshopYes10x10No.checked){
            comp.form.gsshopYes10x10No.checked = false;
        }
        if (comp.checked){
        	document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.ExtNotReg.value="G"
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
			comp.form.ExtNotReg.value="G"
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
			comp.form.ExtNotReg.value="G"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="Y";
			comp.form.extsellyn.value = "Y";
		}
	}
	if (comp.name=="diffName"&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
             comp.form.expensive10x10.checked = false;
        }
		if (comp.checked){
        	document.getElementById("AR").checked=true;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.value="G"
			comp.form.ExtNotReg.disabled = false;
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gsshopYes10x10No")&&(frm.gsshopYes10x10No.checked)){ frm.gsshopYes10x10No.checked=false }
	if ((comp.name!="gsshopNo10x10Yes")&&(frm.gsshopNo10x10Yes.checked)){ frm.gsshopNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
	if ((comp.name!="diffName")&&(frm.diffName.checked)){ frm.diffName.checked=false }
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
function GSShopSelectRegProcess(isreal) {
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

    if (isreal){
        if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "REG";
            document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?\n\n��30�д����� ��ġ ��ϵ˴ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/gsshop/actgsshopReq.asp"
            document.frmSvArr.submit();
        }
    }
}

// ���õ� ��ǰ ���� ����
function GSShopPriceEditProcess() {
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ������ ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditPrice";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

//����� regedoption ���
function Sugi_regedoption() {
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
    document.frmSvArr.target = "xLink";
    document.frmSvArr.ckLimit.value = "<%= limityn %>";
    document.frmSvArr.cmdparam.value = "sugiRegedoption";
    document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
    document.frmSvArr.submit();
}

// ���õ� �̹���(��ǥ �� �����) ����
function GSShopImageEditProcess() {
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� �̹����� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditImage").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditImage";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ���� ����
function GSShopContentsEditProcess() {
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ������ ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditContents").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��� �� �ɼ� �߰�/����
function GSShopOPTEditProcess() {
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� �̹����� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditOPT").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditOPT";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ�� ����
function GSShopItemnameEditProcess() {
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ���� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditItemname";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}

//���ΰ���׸� ����
function GSShopInfodivEditProcess(){
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� ��ǰ���� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditInfoDiv").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditInfoDiv";
        document.frmSvArr.cmdparam.value = "INFODIV";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}


//�⺻���� ����
function GSShopIteminfoEditProcess(){
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

    if (confirm('GSShop�� �����Ͻ� ' + chkSel + '�� �⺻������ ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITINFO";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}


// ���õ� ��ǰ �Ǹſ��� ����
function GSShopSellYnProcess(chkYn) {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��GSShop���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� GSShop���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
    }
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=gsshop&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM6 = window.open("/admin/etc/que/popQueOptionLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function confirmOK(itemcd){
	if (confirm('�ٹ����� ��ǰ�ڵ� : ' + itemcd + '\n���� Ȯ�� �ϼ̽��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditStatCd";
        document.frmSvArr.chgStatItemCode.value = itemcd;
        document.frmSvArr.action = "<%=apiURL%>/outmall/gsshopAddOpt/actgsshopReq.asp"
        document.frmSvArr.submit();
	}
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function chgItemname(idx, iname){
	document.frmUp.target = "xLink";
    document.frmUp.idx.value = idx
    document.frmUp.cName.value = document.getElementById(iname).value;
    document.frmUp.mode.value = "chgName"
    document.frmUp.action = "/admin/etc/optManager/optProc.asp"
    document.frmUp.submit();

}
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
		�귣��&nbsp;&nbsp;&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;<% 'OutmallAdminInfo("gsshop") %>
		&nbsp;
		<a href="https://withgs.gsshop.com/cmm/login" target="_blank">GSShop Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  1003890 | store101010* | 1010tenbyten* ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		GSShop ��ǰ�ڵ� : <textarea rows="2" cols="20" name="gsshopgoodno" id="itemid"><%=replace(gsshopgoodno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >GSShop ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >GSShop ��Ͽ����̻�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >GSShop ��Ͽ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >GSShop ���۽õ��߿���
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >GSShop ����� ���δ��(�ӽ�)
			<option value="G" <%= CHkIIF(ExtNotReg="G","selected","") %> >GSShop ��ϿϷ� ���δ���̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >GSShop ��ϿϷ�(����)
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
		�ٹ����� :
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		�Ǹ�
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
		</select>&nbsp;
		����
		<select name="limityn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
		����
		<select name="sailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
		</select>&nbsp;
		���ظ���(<%= Chkiif(mallid="gsshop", "14.9", "") %>%)
		<select name="onlyValidMargin" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(onlyValidMargin="Y","selected","") %> >�����̻�
			<option value="N" <%= CHkIIF(onlyValidMargin="N","selected","") %> >��������
		</select>&nbsp;
		�ֹ�����
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>&nbsp;
		ǰ��
		<select name="infodiv" class="select">
			<option value="" <%= CHkIIF(infoDiv="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >�Է�
			<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >���Է�
		<%
			For infoLoop = 1 To 35
				If infoLoop < 10 Then
					infoDivValue = "0"&infoLoop
				Else
					infoDivValue = infoLoop
				End If
		%>
			<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
		<% Next %>
		</select>
		<br>
		���޸� &nbsp;&nbsp; :
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		�Ǹ�
		<select name="extsellyn" class="select">
			<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
			<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
			<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
		</select>&nbsp;
		����
		<select name="failCntExists" class="select">
			<option value="" <%= CHkIIF(failCntExists="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >��ϼ�������1ȸ�̻�
			<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >��ϼ�������0ȸ
		</select>&nbsp;
		ī�װ�
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;
		��ǰ�з�
		<select name="MatchPrddiv" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>GSShop ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gsshopYes10x10No" <%= ChkIIF(gsshopYes10x10No="on","checked","") %> ><font color=red>GSShop�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="gsshopNo10x10Yes" <%= ChkIIF(gsshopNo10x10Yes="on","checked","") %> ><font color=red>GSShopǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffName" <%= ChkIIF(diffName="on","checked","") %> ><font color=red>��ǰ/�ɼǸ� �̻�</font>��ǰ����</label>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<form name="frmReg" method="post" action="gsshopItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('gsshop');">&nbsp;&nbsp;
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
	    		<input class="button" type="button" id="btnRegSel" value="��ǰ ���" onClick="GSShopSelectRegProcess(true);">
				<br><br>
				������ǰ ���� :
			    <input class="button" type="button" id="btnEditSel" value="����" onClick="GSShopPriceEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditImage" value="�̹���(��ǥ �� �����)" onClick="GSShopImageEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditContents" value="��ǰ����" onClick="GSShopContentsEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditOPT" value="����&���&�ɼ�&���¼���" onClick="GSShopOPTEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditName" value="��ǰ��" onClick="GSShopItemnameEditProcess();">
   			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfoDiv" value="���ΰ���׸�" onClick="GSShopInfodivEditProcess();">
				<% If (session("ssBctID")="kjy8517") Then %>
				 &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfo" value="�⺻����" onClick="GSShopIteminfoEditProcess();">
				<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ͻ��ߴ�</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="GSShopSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- �׼� �� -->
<!-- ����Ʈ ���� -->
<form name="frmUp" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="idx" value="">
<input type="hidden" name="cName" value="">
<input type="hidden" name="mode" value="">
</form>

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
		�˻���� : <b><%= FormatNumber(oGSShop.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oGSShop.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">��ǰ�ڵ�<br>�ɼ��ڵ�</td>
	<td>��ǰ����</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">GSShop�����<br>GSShop����������</td>
	<td width="70">�ǸŰ�<br><font color="purple">�ɼǰ�</font></td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">GSShop<br>���ݹ��Ǹ�</td>
	<td width="70">GSShop<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">3����<br>�Ǹŷ�</td>
<!--	<td width="50">�귣��<br>����</td> -->
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="100"><font color="BLUE">GS��ǰ�з�</font><br><font color="Green">GS ��������</font><br><font color="BROWN">GS ��������</font></td>
	<td width="80">ǰ��</td>
</tr>
<% For i = 0 To oGSShop.FResultCount - 1 %>
<% If (oGSShop.FItemList(i).FItemName <> oGSShop.FItemList(i).FRegedItemname) OR (oGSShop.FItemList(i).FOptionname <> oGSShop.FItemList(i).FRegedOptionname) Then %>
<tr align="center" bgcolor="GOLD">
<% Else %>
<tr align="center" bgcolor="#FFFFFF">
<% End If %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oGSShop.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oGSShop.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oGSShop.FItemList(i).FItemID %>','GSShop','<%=oGSShop.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oGSShop.FItemList(i).FItemID%>" target="_blank"><%= oGSShop.FItemList(i).FItemID %></a>
		<br><font color="gray"><%=oGSShop.FItemList(i).FItemOption%></font><br>
		<%
			If oGSShop.FItemList(i).getGSShopOptItemStatCd = "���δ��" Then
				response.write "<input type='button' class=button value="&oGSShop.FItemList(i).getGSShopOptItemStatCd&" onclick=confirmOK('"&oGSShop.FItemList(i).FItemID&"')><br>"
			Else
				response.write oGSShop.FItemList(i).getGSShopOptItemStatCd
			End If
		%>
	</td>
	<td align="left">
		<%= oGSShop.FItemList(i).FMakerid %><%= oGSShop.FItemList(i).getDeliverytypeName %>
		<br>(��) : <%= oGSShop.FItemList(i).FItemName %>
		<br>(��) : <%= oGSShop.FItemList(i).FRegedItemname %>
		<br>(��) : <%= oGSShop.FItemList(i).FOptionname %>
		<br>(��) : <%= oGSShop.FItemList(i).FRegedOptionname %>
		<br><input type="text" style="color:red" id="newitemname<%=oGSShop.FItemList(i).FIdx%>" size="50" value="<%= oGSShop.FItemList(i).getRealItemname %>">
		<input type="button" class="button" value="����" onclick="chgItemname('<%= oGSShop.FItemList(i).FIdx %>', 'newitemname<%=oGSShop.FItemList(i).FIdx%>')">
	</td>
	<td align="center"><%= oGSShop.FItemList(i).FRegdate %><br><%= oGSShop.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oGSShop.FItemList(i).FGSShopRegdate %><br><%= oGSShop.FItemList(i).FGSShopLastUpdate %></td>
	<td align="right">
	<% If oGSShop.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oGSShop.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oGSShop.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oGSShop.FItemList(i).FSellcash,0) %>
	<% End If %>
	<br><font color="purple">+<%= oGSShop.FItemList(i).FOptaddprice %></font>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oGSShop.FItemList(i).Fbuycash/oGSShop.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).IsSoldOut Then
			If oGSShop.FItemList(i).FSellyn = "N" Then
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
		If oGSShop.FItemList(i).FItemdiv = "06" OR oGSShop.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oGSShop.FItemList(i).FGSShopStatCd > 0) Then
			If Not IsNULL(oGSShop.FItemList(i).FGSShopPrice) Then
				If (oGSShop.FItemList(i).Fsellcash + oGSShop.FItemList(i).FOptaddprice <> oGSShop.FItemList(i).FGSShopPrice) Then
	%>
					<strong><%= formatNumber(oGSShop.FItemList(i).FGSShopPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oGSShop.FItemList(i).FGSShopPrice,0)&"<br>"
				End If

				If (oGSShop.FItemList(i).FSellyn="Y" and oGSShop.FItemList(i).FGSShopSellYn<>"Y") or (oGSShop.FItemList(i).FSellyn<>"Y" and oGSShop.FItemList(i).FGSShopSellYn="Y") or (oGSShop.FItemList(i).IsSoldOutLimit5Sell and oGSShop.FItemList(i).FGSShopSellYn="Y") Then
	%>
					<strong><%= oGSShop.FItemList(i).FGSShopSellYn %></strong>
	<%
				Else
					response.write oGSShop.FItemList(i).FGSShopSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		'#�ǻ�ǰ��ȣ
		If Not(IsNULL(oGSShop.FItemList(i).FGSShopGoodNo)) Then
	    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.gsshop.com/prd/prd.gs?prdid="&oGSShop.FItemList(i).FGSShopGoodNo&"')>"&oGSShop.FItemList(i).FGSShopGoodNo&"</span>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oGSShop.FItemList(i).FGSShopStatCd="0","(��Ͽ���)","")
		End If
	%>
	</td>
	<td align="center"><%= oGSShop.FItemList(i).Freguserid %></td>
	<td align="center"><%= oGSShop.FItemList(i).FrctSellCNT %></td>
<!--
	<td align="center">
	<%
		If (oGSShop.FItemList(i).FBrandCd = "") OR (oGSShop.FItemList(i).FDeliveryAddrCd = "") OR (oGSShop.FItemList(i).FDeliveryCd = "") Then
	%>
		<font color="darkred">��Ī�ȵ�</font>
	<%
		Else
			response.write "��Ī��"
		End If
	%>
	</td>
-->
	<td align="center">
	<% If oGSShop.FItemList(i).FCateMapCnt > 0 Then %>
	    ��Ī��
	<% Else %>
		<font color="darkred">��Ī�ȵ�</font>
	<% End If %>

	<% If (oGSShop.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oGSShop.FItemList(i).FlastErrStr %>">ERR:<%= oGSShop.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oGSShop.FItemList(i).FDivcode = "" Then
			response.write "��Ī�ȵ�"
		Else
			rw "<font color='BLUE'>��Ī��</font>"
			Select Case oGSShop.FItemList(i).FSafeCode
				Case "1"	response.write "<input type='button' value='�ʼ�' onclick=pop_safecode('"&oGSShop.FItemList(i).FItemid&"'); class='button'>"
					If oGSShop.FItemList(i).FSafeCertGbnCd <> "" Then
						rw "<font color='BLUE'>( Y )</font>"
					Else
						rw "<font color='RED'>( N )</font>"
					End If
				Case "2"	response.write "<input type='button' value='����' onclick=pop_safecode('"&oGSShop.FItemList(i).FItemid&"'); class='button'>"
					If oGSShop.FItemList(i).FSafeCertGbnCd <> "" Then
						rw "<font color='BLUE'>( Y )</font>"
					Else
						rw "<font color='RED'>( N )</font>"
					End If
				Case "3" 	rw "<font color='Green'>����</font>"
			End Select
			Select Case oGSShop.FItemList(i).FIsvat
				Case "01"	rw "<font color='Brown'>�鼼</font>"
				Case "02"	rw "<font color='Brown'>����</font>"
				Case "03" 	rw "<font color='Brown'>����</font>"
				Case "99" 	rw "<font color='Brown'>��ü���</font>"
			End Select
		End If
	%>
	</td>
	<td align="center"><%= oGSShop.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oGSShop.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGSShop.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oGSShop.StartScrollPage to oGSShop.FScrollCount + oGSShop.StartScrollPage - 1 %>
    		<% if i>oGSShop.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oGSShop.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oGSShop = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
