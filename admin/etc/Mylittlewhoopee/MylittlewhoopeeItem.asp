<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/Mylittlewhoopee/Mylittlewhoopeecls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, MylittlewhoopeeGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, MylittlewhoopeeYes10x10No, MylittlewhoopeeNo10x10Yes, reqEdit, reqExpire, failCntExists, scheduleNotInItemid, isextusing, cisextusing, rctsellcnt
Dim page, i, research
Dim oMylittlewhoopee
Dim startMargin, endMargin
Dim purchasetype
page    					= request("page")
research					= request("research")
itemid  					= request("itemid")
makerid						= request("makerid")
itemname					= request("itemname")
bestOrd						= request("bestOrd")
bestOrdMall					= request("bestOrdMall")
sellyn						= request("sellyn")
limityn						= request("limityn")
sailyn						= request("sailyn")
onlyValidMargin				= request("onlyValidMargin")
startMargin					= request("startMargin")
endMargin					= request("endMargin")
isMadeHand					= request("isMadeHand")
isOption					= request("isOption")
infoDiv						= request("infoDiv")
morningJY					= request("morningJY")
extsellyn					= request("extsellyn")
MylittlewhoopeeGoodNo		= request("MylittlewhoopeeGoodNo")
ExtNotReg					= request("ExtNotReg")
isReged						= request("isReged")
MatchCate					= request("MatchCate")
expensive10x10				= request("expensive10x10")
diffPrc						= request("diffPrc")
MylittlewhoopeeYes10x10No	= request("MylittlewhoopeeYes10x10No")
MylittlewhoopeeNo10x10Yes	= request("MylittlewhoopeeNo10x10Yes")
reqEdit						= request("reqEdit")
reqExpire					= request("reqExpire")
failCntExists				= request("failCntExists")
optAddPrcRegTypeNone		= request("optAddPrcRegTypeNone")
notinmakerid				= request("notinmakerid")
priceOption					= request("priceOption")
isSpecialPrice				= request("isSpecialPrice")
deliverytype				= request("deliverytype")
mwdiv						= request("mwdiv")
notinitemid					= requestCheckVar(request("notinitemid"), 1)
exctrans					= requestCheckVar(request("exctrans"), 1)
scheduleNotInItemid			= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing					= requestCheckVar(request("isextusing"), 1)
cisextusing					= requestCheckVar(request("cisextusing"), 1)
rctsellcnt					= requestCheckVar(request("rctsellcnt"), 1)
purchasetype				= request("purchasetype")

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
'������� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If MylittlewhoopeeGoodNo <> "" then
	Dim iA2, arrTemp2, arrMylittlewhoopeeGoodNo
	MylittlewhoopeeGoodNo = replace(MylittlewhoopeeGoodNo,",",chr(10))
	MylittlewhoopeeGoodNo = replace(MylittlewhoopeeGoodNo,chr(13),"")
	arrTemp2 = Split(MylittlewhoopeeGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrMylittlewhoopeeGoodNo = arrMylittlewhoopeeGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	MylittlewhoopeeGoodNo = left(arrMylittlewhoopeeGoodNo,len(arrMylittlewhoopeeGoodNo)-1)
End If

Set oMylittlewhoopee = new CMylittlewhoopee
	oMylittlewhoopee.FCurrPage						= page
If (session("ssBctID")="kjy8517") Then
	oMylittlewhoopee.FPageSize						= 100
Else
	oMylittlewhoopee.FPageSize						= 50
End If
	oMylittlewhoopee.FRectCDL						= request("cdl")
	oMylittlewhoopee.FRectCDM						= request("cdm")
	oMylittlewhoopee.FRectCDS						= request("cds")
	oMylittlewhoopee.FRectItemID					= itemid
	oMylittlewhoopee.FRectItemName					= itemname
	oMylittlewhoopee.FRectSellYn					= sellyn
	oMylittlewhoopee.FRectLimitYn					= limityn
	oMylittlewhoopee.FRectSailYn					= sailyn
	oMylittlewhoopee.FRectStartMargin				= startMargin
	oMylittlewhoopee.FRectEndMargin					= endMargin
	oMylittlewhoopee.FRectMakerid					= makerid
	oMylittlewhoopee.FRectMylittlewhoopeeGoodNo		= MylittlewhoopeeGoodNo
	oMylittlewhoopee.FRectMatchCate					= MatchCate
	oMylittlewhoopee.FRectIsMadeHand				= isMadeHand
	oMylittlewhoopee.FRectIsOption					= isOption
	oMylittlewhoopee.FRectIsReged					= isReged
	oMylittlewhoopee.FRectNotinmakerid				= notinmakerid
	oMylittlewhoopee.FRectNotinitemid				= notinitemid
	oMylittlewhoopee.FRectExcTrans					= exctrans
	oMylittlewhoopee.FRectPriceOption				= priceOption
	oMylittlewhoopee.FRectIsSpecialPrice  	   		= isSpecialPrice
	oMylittlewhoopee.FRectDeliverytype				= deliverytype
	oMylittlewhoopee.FRectMwdiv						= mwdiv
	oMylittlewhoopee.FRectScheduleNotInItemid		= scheduleNotInItemid
	oMylittlewhoopee.FRectIsextusing				= isextusing
	oMylittlewhoopee.FRectCisextusing				= cisextusing
	oMylittlewhoopee.FRectRctsellcnt				= rctsellcnt

	oMylittlewhoopee.FRectExtNotReg					= ExtNotReg
	oMylittlewhoopee.FRectExpensive10x10			= expensive10x10
	oMylittlewhoopee.FRectdiffPrc					= diffPrc
	oMylittlewhoopee.FRectMylittlewhoopeeYes10x10No = MylittlewhoopeeYes10x10No
	oMylittlewhoopee.FRectMylittlewhoopeeNo10x10Yes = MylittlewhoopeeNo10x10Yes
	oMylittlewhoopee.FRectExtSellYn					= extsellyn
	oMylittlewhoopee.FRectInfoDiv					= infoDiv
	oMylittlewhoopee.FRectFailCntOverExcept			= ""
	oMylittlewhoopee.FRectFailCntExists				= failCntExists
	oMylittlewhoopee.FRectReqEdit					= reqEdit
	oMylittlewhoopee.FRectPurchasetype				= purchasetype
If (bestOrd = "on") Then
    oMylittlewhoopee.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oMylittlewhoopee.FRectOrdType = "BM"
End If

If isReged = "R" Then								'ǰ��ó����� ��ǰ���� ����Ʈ
	oMylittlewhoopee.getMylittlewhoopeereqExpireItemList
Else
	oMylittlewhoopee.getMylittlewhoopeeRegedItemList		'�� �� ����Ʈ
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

// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=Mylittlewhoopee","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=Mylittlewhoopee','popNotInItemid','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ���
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=Mylittlewhoopee','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� �귣��(EP)
function NotInMakeridEP(){
    var popwin = window.open("/admin/etc/potal/notinmakerid.asp?mallid=naverEP","popNotInMakeridEP","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ(EP)
function NotInItemidEP(){
	var popwin2=window.open('/admin/etc/potal/notinitemid.asp?mallid=naverEP','popNotInItemidEP','width=1200,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=Mylittlewhoopee','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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

    if ((comp.name=="MylittlewhoopeeYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="MylittlewhoopeeNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.MylittlewhoopeeYes10x10No.checked){
            comp.form.MylittlewhoopeeYes10x10No.checked = false;
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
	if ((comp.name!="MylittlewhoopeeYes10x10No")&&(frm.MylittlewhoopeeYes10x10No.checked)){ frm.MylittlewhoopeeYes10x10No.checked=false }
	if ((comp.name!="MylittlewhoopeeNo10x10Yes")&&(frm.MylittlewhoopeeNo10x10Yes.checked)){ frm.MylittlewhoopeeNo10x10Yes.checked=false }
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
	if ((comp.name!="MylittlewhoopeeYes10x10No")&&(frm.MylittlewhoopeeYes10x10No.checked)){ frm.MylittlewhoopeeYes10x10No.checked=false }
	if ((comp.name!="MylittlewhoopeeNo10x10Yes")&&(frm.MylittlewhoopeeNo10x10Yes.checked)){ frm.MylittlewhoopeeNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
// ���õ� ��ǰ ���
function MylittlewhoopeeSelectRegItemProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function MylittlewhoopeeSelectRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function MylittlewhoopeeSelectEDITProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ��ȸ
function MylittlewhoopeeSelectItemSearchProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ȸ
function MylittlewhoopeeSelectOptionSearchProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �̹��� ���
function MylittlewhoopeeSelectImageRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ɼ� ���
function MylittlewhoopeeSelectOPTRegProcess() {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}

function MylittlewhoopeeSelectDelProcess(){
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}

function MylittlewhoopeeSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp"
        document.frmSvArr.submit();
    }
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// ������� ī�װ��� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/nvstorefarm/popnvstorefarmCateList.asp","popnvstorefarm","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//�����ڵ� �˻�
function NvstorefarmCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "<%=apiURL%>/outmall/Mylittlewhoopee/actMylittlewhoopeeReq.asp?cmdparam=nvstorefarmCommonCode&CommCD="+ccd+"";
	}
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=Mylittlewhoopee&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('http://scm.10x10.co.kr/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
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
		������� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="MylittlewhoopeeGoodNo" id="itemid"><%= replace(replace(MylittlewhoopeeGoodNo,",",chr(10)), "'", "")%></textarea>
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
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">�̵��<font color="<%= CHKIIF(makerid="" and itemid="", "#000000", "#AAAAAA") %>">(�ֱ� 3���� ��ϻ�ǰ��)</font></label>&nbsp;
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="MylittlewhoopeeYes10x10No" <%= ChkIIF(MylittlewhoopeeYes10x10No="on","checked","") %> ><font color=red>��������Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="MylittlewhoopeeNo10x10Yes" <%= ChkIIF(MylittlewhoopeeNo10x10Yes="on","checked","") %> ><font color=red>�������ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>5) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>

<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, <del>�ǸŰ�(���ΰ�) 1���� �̸�</del>, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ֹ����ۻ�ǰ, �ֹ����۹�����ǰ, �ǸŰ�(���ΰ�) 1õ�� �̸�, �Ϻ� ǰ��(ȭ��ǰ, ��ǰ(����깰), ������ǰ, �ǰ���ɽ�ǰ) ��ǰ, �ɼǰ� 0�� �Ǹ��� ��ǰ�� ����(�ɼ� �������� 5�� ���� ����)

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
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ī�װ���" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">&nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('Mylittlewhoopee');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ī�װ���" onclick="pop_CateManager();">
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
				<input class="button" type="button" id="btnRegImgSel" value="�̹���" onClick="MylittlewhoopeeSelectImageRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="yhj0613") OR (session("ssBctID")="hrkang97") Then %>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnRegSel" value="��ǰ" onClick="MylittlewhoopeeSelectRegItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegOptSel" value="�ɼ�" onClick="MylittlewhoopeeSelectOPTRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelSel" value="����" onClick="MylittlewhoopeeSelectDelProcess();">&nbsp;&nbsp;
				<% Else %>
				<input class="button" type="button" id="btnDelSel" value="����" onClick="MylittlewhoopeeSelectDelProcess();">&nbsp;&nbsp;
				<% End If %>
			<% End If %>
				<input class="button" type="button" id="btnReg" value="��ǰ+�ɼ�" onClick="MylittlewhoopeeSelectRegProcess();" style=color:red;font-weight:bold>
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnReg" value="����" onClick="MylittlewhoopeeSelectEDITProcess();" style=color:blue;font-weight:bold>
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				������ǰ ��ȸ :
				<input class="button" type="button" id="btnSchitem" value="��ǰ" onClick="MylittlewhoopeeSelectItemSearchProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnSchOpt" value="�ɼ�" onClick="MylittlewhoopeeSelectOptionSearchProcess();">&nbsp;&nbsp;
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="GetAddressBookList">�Ǹ����ּ�
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="MylittlewhoopeeCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="MylittlewhoopeeSellYnProcess(frmReg.chgSellYn.value);">
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
		�˻���� : <b><%= FormatNumber(oMylittlewhoopee.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMylittlewhoopee.FTotalPage,0) %></b>
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
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�������<br>���ݹ��Ǹ�</td>
	<td width="70">�������<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ���<br>��Ī����</td>
	<td width="80">ǰ��</td>
	<td width="100">�̹���<br>���ε�</td>
</tr>
<% For i=0 to oMylittlewhoopee.FResultCount - 1 %>
<tr align="center" <% If oMylittlewhoopee.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oMylittlewhoopee.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oMylittlewhoopee.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oMylittlewhoopee.FItemList(i).FItemID %>','Mylittlewhoopee','<%=oMylittlewhoopee.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oMylittlewhoopee.FItemList(i).FItemID%>" target="_blank"><%= oMylittlewhoopee.FItemList(i).FItemID %></a>
		<% If oMylittlewhoopee.FItemList(i).FMylittlewhoopeeStatcd <> 7 Then %>
		<br><%= oMylittlewhoopee.FItemList(i).getMylittlewhoopeeStatName %>
		<% End If %>
		<%= oMylittlewhoopee.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oMylittlewhoopee.FItemList(i).FMakerid %> <%= oMylittlewhoopee.FItemList(i).getDeliverytypeName %><br><%= oMylittlewhoopee.FItemList(i).FItemName %></td>
	<td align="center"><%= oMylittlewhoopee.FItemList(i).FRegdate %><br><%= oMylittlewhoopee.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oMylittlewhoopee.FItemList(i).FMylittlewhoopeeRegdate %><br><%= oMylittlewhoopee.FItemList(i).FMylittlewhoopeeLastUpdate %></td>
	<td align="right">
		<% If oMylittlewhoopee.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oMylittlewhoopee.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oMylittlewhoopee.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oMylittlewhoopee.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oMylittlewhoopee.FItemList(i).Fsellcash = 0 Then
		elseif (oMylittlewhoopee.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oMylittlewhoopee.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oMylittlewhoopee.FItemList(i).FOrgSuplycash/oMylittlewhoopee.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oMylittlewhoopee.FItemList(i).Fbuycash/oMylittlewhoopee.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oMylittlewhoopee.FItemList(i).Fbuycash/oMylittlewhoopee.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oMylittlewhoopee.FItemList(i).IsSoldOut Then
			If oMylittlewhoopee.FItemList(i).FSellyn = "N" Then
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
		If oMylittlewhoopee.FItemList(i).FItemdiv = "06" OR oMylittlewhoopee.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oMylittlewhoopee.FItemList(i).FMylittlewhoopeeStatCd > 0) Then
			If Not IsNULL(oMylittlewhoopee.FItemList(i).FMylittlewhoopeePrice) Then
				If (oMylittlewhoopee.FItemList(i).Fsellcash <> oMylittlewhoopee.FItemList(i).FMylittlewhoopeePrice) Then
	%>
					<strong><%= formatNumber(oMylittlewhoopee.FItemList(i).FMylittlewhoopeePrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oMylittlewhoopee.FItemList(i).FMylittlewhoopeePrice,0)&"<br>"
				End If

				If Not IsNULL(oMylittlewhoopee.FItemList(i).FSpecialPrice) Then
					If (now() >= oMylittlewhoopee.FItemList(i).FStartDate) And (now() <= oMylittlewhoopee.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oMylittlewhoopee.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oMylittlewhoopee.FItemList(i).FSellyn="Y" and oMylittlewhoopee.FItemList(i).FMylittlewhoopeeSellYn<>"Y") or (oMylittlewhoopee.FItemList(i).FSellyn<>"Y" and oMylittlewhoopee.FItemList(i).FMylittlewhoopeeSellYn="Y") Then
	%>
					<strong><%= oMylittlewhoopee.FItemList(i).FMylittlewhoopeeSellYn %></strong>
	<%
				Else
					response.write oMylittlewhoopee.FItemList(i).FMylittlewhoopeeSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oMylittlewhoopee.FItemList(i).FMylittlewhoopeeGoodNo)) Then
			Response.Write "<a target='_blank' href='http://storefarm.naver.com/mylittlewhoopee/products/"&oMylittlewhoopee.FItemList(i).FMylittlewhoopeeGoodNo&"'>"&oMylittlewhoopee.FItemList(i).FMylittlewhoopeeGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oMylittlewhoopee.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oMylittlewhoopee.FItemList(i).FItemID%>','0');"><%= oMylittlewhoopee.FItemList(i).FoptionCnt %>:<%= oMylittlewhoopee.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oMylittlewhoopee.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oMylittlewhoopee.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oMylittlewhoopee.FItemList(i).FinfoDiv %>
		<%
		If (oMylittlewhoopee.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oMylittlewhoopee.FItemList(i).FlastErrStr) &"'>ERR:"& oMylittlewhoopee.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oMylittlewhoopee.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oMylittlewhoopee.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>�̵��</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oMylittlewhoopee.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMylittlewhoopee.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oMylittlewhoopee.StartScrollPage to oMylittlewhoopee.FScrollCount + oMylittlewhoopee.StartScrollPage - 1 %>
    		<% if i>oMylittlewhoopee.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oMylittlewhoopee.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oMylittlewhoopee = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->