<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/lotte/lotteitemcls.asp"-->
<%
Dim mallid, infoLoop, infoDivValue
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, lotteGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, priceOption, lotteTmpGoodNo, deliverytype, mwdiv
Dim expensive10x10, diffPrc, lotteYes10x10No, lotteNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oLotteitem

mallid					= CMALLNAME
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
lotteGoodNo				= request("lotteGoodNo")
lotteTmpGoodNo			= request("lotteTmpGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
lotteYes10x10No			= request("lotteYes10x10No")
lotteNo10x10Yes			= request("lotteNo10x10Yes")
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
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
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
'�Ե����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If lotteGoodNo <> "" then
	Dim iA2, arrTemp2, arrlotteGoodNo
	lotteGoodNo = replace(lotteGoodNo,",",chr(10))
	lotteGoodNo = replace(lotteGoodNo,chr(13),"")
	arrTemp2 = Split(lotteGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrlotteGoodNo = arrlotteGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	lotteGoodNo = left(arrlotteGoodNo,len(arrlotteGoodNo)-1)
End If

'�Ե����� ������ ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If lotteTmpGoodNo <> "" then
	Dim iA3, arrTemp3, arrlotteTmpGoodNo
	lotteTmpGoodNo = replace(lotteTmpGoodNo,",",chr(10))
	lotteTmpGoodNo = replace(lotteTmpGoodNo,chr(13),"")
	arrTemp3 = Split(lotteTmpGoodNo,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			If Not(isNumeric(trim(arrTemp3(iA3)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp3(iA3) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrlotteTmpGoodNo = arrlotteTmpGoodNo & trim(arrTemp3(iA3)) & ","
			End If
		End If
		iA3 = iA3 + 1
	Loop
	lotteTmpGoodNo = left(arrlotteTmpGoodNo,len(arrlotteTmpGoodNo)-1)
End If

Set oLotteitem = new CLotte
If (session("ssBctID")="kjy8517") Then
	oLotteitem.FPageSize					= 50
Else
	oLotteitem.FPageSize					= 20
End If
	oLotteitem.FCurrPage					= page
	oLotteitem.FRectMakerid					= makerid
	oLotteitem.FRectItemID					= itemid
	oLotteitem.FRectItemName				= itemname
	oLotteitem.FRectLotteGoodNo				= lotteGoodNo
	oLotteitem.FRectLotteTmpGoodNo			= lotteTmpGoodNo
	oLotteitem.FRectCDL						= request("cdl")
	oLotteitem.FRectCDM						= request("cdm")
	oLotteitem.FRectCDS						= request("cds")
	oLotteitem.FRectExtNotReg				= ExtNotReg
	oLotteitem.FRectIsReged					= isReged
	oLotteitem.FRectNotinmakerid			= notinmakerid
	oLotteitem.FRectPriceOption				= priceOption
	oLotteitem.FRectDeliverytype			= deliverytype
	oLotteitem.FRectMwdiv					= mwdiv

	oLotteitem.FRectSellYn					= sellyn
	oLotteitem.FRectLimitYn					= limityn
	oLotteitem.FRectSailYn					= sailyn
	oLotteitem.FRectonlyValidMargin			= onlyValidMargin
	oLotteitem.FRectIsMadeHand				= isMadeHand
	oLotteitem.FRectIsOption				= isOption
	oLotteitem.FRectInfoDiv					= infoDiv
	oLotteitem.FRectExtSellYn				= extsellyn
	oLotteitem.FRectFailCntExists			= failCntExists
	oLotteitem.FRectMatchCate				= MatchCate
	oLotteitem.FRectExpensive10x10			= expensive10x10
	oLotteitem.FRectdiffPrc					= diffPrc
	oLotteitem.FRectLotteYes10x10No		= lotteYes10x10No
	oLotteitem.FRectLotteNo10x10Yes			= lotteNo10x10Yes
	oLotteitem.FRectReqEdit					= reqEdit
If (bestOrd = "on") Then
    oLotteitem.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oLotteitem.FRectOrdType = "BM"
End If
	oLotteitem.getLotteAddOptionRegedItemList		'�� �� ����Ʈ
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteCom","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteCom','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// �Ե����� ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/lotte/popLotteCateList.asp","popCateMan","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
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
	if ((comp.name!="lotteYes10x10No")&&(frm.lotteYes10x10No.checked)){ frm.lotteYes10x10No.checked=false }
	if ((comp.name!="lotteNo10x10Yes")&&(frm.lotteNo10x10Yes.checked)){ frm.lotteNo10x10Yes.checked=false }
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

    if ((comp.name=="lotteYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="lotteNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.lotteYes10x10No.checked){
            comp.form.lotteYes10x10No.checked = false;
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
	if ((comp.name!="lotteYes10x10No")&&(frm.lotteYes10x10No.checked)){ frm.lotteYes10x10No.checked=false }
	if ((comp.name!="lotteNo10x10Yes")&&(frm.lotteNo10x10Yes.checked)){ frm.lotteNo10x10Yes.checked=false }
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
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM6 = window.open("/admin/etc/que/popQueOptionLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}

// ���õ� ��ǰ �Ǹſ��� ����
function LotteSellYnProcess(chkYn) {
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
		case "X": strSell="�Ǹ�����(����)";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� �Ե����Ŀ��� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

        document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function LotteSelectRegProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��Ͽ��� ��ǰ �ϰ� ���
function LotteregIMSI(isreg) {
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

    if (isreg){
        if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
			document.frmSvArr.submit();
        }
    }
}

//���� �űԻ�ǰ��ȸ
function LotteStatCheckProcess(){
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹ� ���¸� Ȯ�� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function LottePriceEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �̹��� ����
function LotteImageEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ �̹����� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}

//���õ� ��ǰ ��ǰ�� ����
function LotteItemnameEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ���� �ϰ� ���� ��û �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
function LotteInfodivEditProcess(){
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ǰ���� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "INFODIV";
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function LotteEditProcess(v) {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		if(v == ""){
			document.getElementById("btnEditSel").disabled=true;
		}else{
			document.getElementById("btnEditSel2").disabled=true;
		}
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT" + v;
		document.frmSvArr.action = "<%=apiURL%>/outmall/lotteComAddOpt/actLotteComReq.asp"
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
		<a href="https://partner.lotte.com/main/Login.lotte" target="_blank">�Ե�����Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 124072 | store101010** ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		�Ե����� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="lotteGoodNo" id="itemid"><%=replace(lotteGoodNo,",",chr(10))%></textarea>
		&nbsp;
		������ ��ǰ�ڵ� : <textarea rows="2" cols="20" name="lotteTmpGoodNo" id="itemid"><%=replace(lotteTmpGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >�Ե����� ��Ͻ���
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >�Ե����� ��Ͽ���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >�Ե����� �ݷ�
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >�Ե����� ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >�Ե����� ��ϿϷ�(����)
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
		���ظ���(<%= Chkiif(mallid="lotteCom", "15", "") %>%)
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>�Ե����� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="lotteYes10x10No" <%= ChkIIF(lotteYes10x10No="on","checked","") %> ><font color=red>�Ե������Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="lotteNo10x10Yes" <%= ChkIIF(lotteNo10x10Yes="on","checked","") %> ><font color=red>�Ե�����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="lotteItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lotteCom');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="LotteSelectRegProcess();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="LottePriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="����&����&�ɼ�&����" onClick="LotteEditProcess('');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditName" value="��ǰ��" onClick="LotteItemnameEditProcess();">
				&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditInfoDiv" value="���ΰ���׸�" onClick="LotteInfodivEditProcess();">
   				&nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditImage" value="�̹���" onClick="LotteImageEditProcess();">
				<br><br>
				���ο��� ��ǰ :
				<input class="button" type="button" id="btnEditSel2" value="����" onClick="LotteEditProcess('2');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnchk" value="�űԻ�ǰ��ȸ" onClick="LotteStatCheckProcess();">
				<br><br>
				��Ͽ��� ��ǰ :
				<input class="button" type="button" id="btnRegImsi" value="���" onClick="LotteregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="����" onClick="LotteregIMSI(false);" >
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
					<option value="X">�Ǹ�����(����)</option><!-- �����ϸ� ���� ���� �� �� ���� -->
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="LotteSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
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
		�˻���� : <b><%= FormatNumber(oLotteitem.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oLotteitem.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">��ǰ�ڵ�<br>�ɼ��ڵ�</td>
	<td>��ǰ����</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">�Ե����ĵ����<br>�Ե���������������</td>
	<td width="70">�ǸŰ�<br><font color="purple">�ɼǰ�</font></td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�Ե�����<br>���ݹ��Ǹ�</td>
	<td width="70">�Ե�����<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i = 0 To oLotteitem.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oLotteitem.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oLotteitem.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oLotteitem.FItemList(i).FItemID %>','lotteCom','<%=oLotteitem.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oLotteitem.FItemList(i).FItemID%>" target="_blank"><%= oLotteitem.FItemList(i).FItemID %></a>
		<br><font color="gray"><%=oLotteitem.FItemList(i).FItemOption%></font><br>
		<% if oLotteitem.FItemList(i).FLimitYn="Y" then %><br><%= oLotteitem.FItemList(i).getLimitHtmlStr %></font><% end if %>
	</td>
	<td align="left">
		<%= oLotteitem.FItemList(i).FMakerid %><%= oLotteitem.FItemList(i).getDeliverytypeName %>
		<br>(��) : <%= oLotteitem.FItemList(i).FItemName %>
		<br>(��) : <%= oLotteitem.FItemList(i).FRegedItemname %>
		<br>(��) : <%= oLotteitem.FItemList(i).FOptionname %>
		<br>(��) : <%= oLotteitem.FItemList(i).FRegedOptionname %>
		<br><input type="text" style="color:red" id="newitemname<%=oLotteitem.FItemList(i).FIdx%>" size="50" value="<%= oLotteitem.FItemList(i).getRealItemname %>">
		<input type="button" class="button" value="����" onclick="chgItemname('<%= oLotteitem.FItemList(i).FIdx %>', 'newitemname<%=oLotteitem.FItemList(i).FIdx%>')">
	</td>
	<td align="center"><%= oLotteitem.FItemList(i).FRegdate %><br><%= oLotteitem.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oLotteitem.FItemList(i).FLotteRegdate %><br><%= oLotteitem.FItemList(i).FLotteLastUpdate %></td>

	<td align="right">
	<% If oLotteitem.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oLotteitem.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %>
	<% End If %>
	<br><font color="purple">+<%= oLotteitem.FItemList(i).FOptaddprice %></font>
	</td>
	<td align="center">
	<%
		If oLotteitem.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oLotteitem.FItemList(i).IsSoldOut Then
			If oLotteitem.FItemList(i).FSellyn = "N" Then
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
		If oLotteitem.FItemList(i).FItemdiv = "06" OR oLotteitem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oLotteitem.FItemList(i).FLotteStatCd > 0) Then
			If Not IsNULL(oLotteitem.FItemList(i).FLottePrice) Then
				If (oLotteitem.FItemList(i).Fsellcash + oLotteitem.FItemList(i).FOptaddprice <> oLotteitem.FItemList(i).FLottePrice) Then
	%>
					<strong><%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oLotteitem.FItemList(i).FLottePrice,0)&"<br>"
				End If

				If (oLotteitem.FItemList(i).FSellyn="Y" and oLotteitem.FItemList(i).FLotteSellYn<>"Y") or (oLotteitem.FItemList(i).FSellyn<>"Y" and oLotteitem.FItemList(i).FLotteSellYn="Y") Then
	%>
					<strong><%= oLotteitem.FItemList(i).FLotteSellYn %></strong>
	<%
				Else
					response.write oLotteitem.FItemList(i).FLotteSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
    <%
    	'#�ǻ�ǰ��ȣ
    	if Not(IsNULL(oLotteitem.FItemList(i).FLotteGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotte.com/goods/viewGoodsDetail.lotte?goods_no="&oLotteitem.FItemList(i).FLotteGoodNo&"'>"&oLotteitem.FItemList(i).FLotteGoodNo&"</a>"
		else
			'#�ӽû�ǰ��ȣ
			if Not(IsNULL(oLotteitem.FItemList(i).FLotteTmpGoodNo)) then
				if oLotteitem.FItemList(i).FLotteStatCd<>"30" then
					Response.Write oLotteitem.FItemList(i).getLotteItemStatCd & "<br>(" & oLotteitem.FItemList(i).FLotteTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if
	%>
	</td>
	<td align="center"><%= oLotteitem.FItemList(i).Freguserid %></td>
	<td align="center"><%= oLotteitem.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oLotteitem.FItemList(i).FCateMapCnt > 0 Then %>
	    ��Ī��
	<% Else %>
		<font color="darkred">��Ī�ȵ�</font>
	<% End If %>

	<% If (oLotteitem.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oLotteitem.FItemList(i).FlastErrStr %>">ERR:<%= oLotteitem.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
    <td align="center"><%= oLotteitem.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oLotteitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotteitem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLotteitem.StartScrollPage to oLotteitem.FScrollCount + oLotteitem.StartScrollPage - 1 %>
    		<% if i>oLotteitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLotteitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oLotteitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->