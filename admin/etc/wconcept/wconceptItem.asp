<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wconcept/wconceptcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, wconceptKeepSell, isSpecialPrice
Dim bestOrdMall, wconceptGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchBrand, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, wconceptYes10x10No, wconceptNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, wconceptGoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim oWconcept, xl, kjypageSize, purchasetype
Dim startMargin, endMargin
page    				= request("page")
kjypageSize				= request("kjypageSize")
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
wconceptGoodNo			= request("wconceptGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchBrand				= request("MatchBrand")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
wconceptYes10x10No		= request("wconceptYes10x10No")
wconceptNo10x10Yes		= request("wconceptNo10x10Yes")
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
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
purchasetype 			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchBrand= ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = ""
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

'wconcept ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If wconceptGoodNo <> "" then
	Dim iA2, arrTemp2, arrwconceptGoodNo
	wconceptGoodNo = replace(wconceptGoodNo,",",chr(10))
	wconceptGoodNo = replace(wconceptGoodNo,chr(13),"")
	arrTemp2 = Split(wconceptGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrwconceptGoodNo = arrwconceptGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	wconceptGoodNo = left(arrwconceptGoodNo,len(arrwconceptGoodNo)-1)
End If

Set oWconcept = new Cwconcept
	oWconcept.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oWconcept.FPageSize					= kjypageSize
Else
	oWconcept.FPageSize					= 100
End If
	oWconcept.FRectCDL					= request("cdl")
	oWconcept.FRectCDM					= request("cdm")
	oWconcept.FRectCDS					= request("cds")
	oWconcept.FRectItemID				= itemid
	oWconcept.FRectItemName				= itemname
	oWconcept.FRectSellYn				= sellyn
	oWconcept.FRectLimitYn				= limityn
	oWconcept.FRectSailYn				= sailyn
	oWconcept.FRectStartMargin			= startMargin
	oWconcept.FRectEndMargin			= endMargin
	oWconcept.FRectMakerid				= makerid
	oWconcept.FRectwconceptGoodNo		= wconceptGoodNo
	oWconcept.FRectMatchCate			= MatchCate
	oWconcept.FRectMatchBrand			= MatchBrand
	oWconcept.FRectIsMadeHand			= isMadeHand
	oWconcept.FRectIsOption				= isOption
	oWconcept.FRectIsReged				= isReged
	oWconcept.FRectNotinmakerid			= notinmakerid
	oWconcept.FRectNotinitemid			= notinitemid
	oWconcept.FRectExcTrans				= exctrans
	oWconcept.FRectPriceOption			= priceOption
	oWconcept.FRectIsSpecialPrice     	= isSpecialPrice
	oWconcept.FRectDeliverytype			= deliverytype
	oWconcept.FRectMwdiv				= mwdiv
	oWconcept.FRectScheduleNotInItemid	= scheduleNotInItemid
	oWconcept.FRectIsextusing			= isextusing
	oWconcept.FRectCisextusing			= cisextusing
	oWconcept.FRectRctsellcnt			= rctsellcnt
	oWconcept.FRectPurchasetype 		= purchasetype

	oWconcept.FRectExtNotReg			= ExtNotReg
	oWconcept.FRectExpensive10x10		= expensive10x10
	oWconcept.FRectdiffPrc				= diffPrc
	oWconcept.FRectwconceptYes10x10No	= wconceptYes10x10No
	oWconcept.FRectwconceptNo10x10Yes	= wconceptNo10x10Yes
	oWconcept.FRectwconceptKeepSell		= wconceptKeepSell
	oWconcept.FRectExtSellYn			= extsellyn
	oWconcept.FRectInfoDiv				= infoDiv
	oWconcept.FRectFailCntOverExcept	= ""
	oWconcept.FRectFailCntExists		= failCntExists
	oWconcept.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oWconcept.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oWconcept.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oWconcept.getwconceptreqExpireItemList
Else
	oWconcept.getwconceptRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=wconceptList"& replace(DATE(), "-", "") &"_xl.xls"
Else
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=wconcept1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=wconcept1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=wconcept1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=wconcept1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="wconceptKeepSell")&&(frm.wconceptKeepSell.checked)){ frm.wconceptKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="wconceptYes10x10No")&&(frm.wconceptYes10x10No.checked)){ frm.wconceptYes10x10No.checked=false }
	if ((comp.name!="wconceptNo10x10Yes")&&(frm.wconceptNo10x10Yes.checked)){ frm.wconceptNo10x10Yes.checked=false }
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

    if ((comp.name=="wconceptYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="wconceptNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.wconceptYes10x10No.checked){
            comp.form.wconceptYes10x10No.checked = false;
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
	if ((comp.name!="wconceptYes10x10No")&&(frm.wconceptYes10x10No.checked)){ frm.wconceptYes10x10No.checked=false }
	if ((comp.name!="wconceptNo10x10Yes")&&(frm.wconceptNo10x10Yes.checked)){ frm.wconceptNo10x10Yes.checked=false }
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
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/wconcept/popwconceptCateList.asp","popCatewconceptmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

//�����ڵ� �˻�
function fnwconceptCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;

	if(ccd == ''){
		alert('�����ڵ带 �Է��ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "/admin/etc/wconcept/actwconceptReq.asp?cmdparam=wconceptCommonCode&CommCD="+ccd+"";
	}
}

// ���õ� ��ǰ ���
function wconceptSelectRegProcess() {
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

    if (confirm('wconcept�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

function wconceptSelectRegStepProcess(v) {
	var chkSel=0;
	var strStep;
	switch(v) {
		case 1 :
			strStep="REGSTEP1";
			break;
		case 2 :
			strStep="REGSTEP2";
			break;
		case 3 :
			strStep="REGSTEP3";
			break;
		case 4 :
			strStep="REGSTEP4";
			break;
		case 5 :
			strStep="REGSTEP5";
			break;
		case 6 :
			strStep="REGSTEP6";
			break;

	}
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

    if (confirm('wconcept�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = strStep;
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function wconceptSellYnProcess(chkYn) {
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
		case "X": strSell="������ ��ǰ����";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
       	document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function wconceptEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function wconceptSelectPriceEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//������ ����
function wconceptSelectContentEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ �������� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CONTENT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//�̹��� ����
function wconceptSelectImageEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ �̹����� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "IMAGE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//�߰� �̹��� ����
function wconceptSelectAddImageEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ �߰��̹����� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "ADDIMAGE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//�ɼ� ����
function wconceptSelectOptionEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ �ɼ��� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}

//������� ����
function wconceptSelectInfoCodeProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ �ɼ��� ���� �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "INFOCODE";
        document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
        document.frmSvArr.submit();
    }
}


//��ǰ + �ɼ� ��ȸ
function wconceptChkStatProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}

//��ǰ ��ȸ
function wconceptChkItemProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKITEM";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}

//�ɼ� ��ȸ
function wconceptChkOptProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� �ɼ��� ��ȸ �Ͻðڽ��ϱ�?\n\n��wconcept���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKOPT";
		document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
		document.frmSvArr.submit();
    }
}


//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=wconcept1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function wconceptEditProcessAuto() {
	var cnt = <%= oWconcept.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "/admin/etc/wconcept/actwconceptReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oWconcept.FResultCount %>;
	if (cnt === 0) {
		// 45�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		wconceptEditProcessAuto();
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
function popXL()
{
    frmXL.submit();
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
		<a href="https://newpin.wconcept.co.kr/Auth/Login" target="_blank">wconceptAdmin�ٷΰ���</a>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		wconcept ��ǰ�ڵ� : <textarea rows="2" cols="20" name="wconceptGoodNo" id="itemid"><%=replace(wconceptGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >wconcept ��Ͻõ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >wconcept ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >wconcept ��ϿϷ�(����)
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
		<% If (session("ssBctID")="kjy8517") Then %>
			<input class="text" size="5" type="text" name="kjypageSize" value="<%= kjypageSize %>">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>wconcept ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptYes10x10No" <%= ChkIIF(wconceptYes10x10No="on","checked","") %> ><font color=red>wconcept�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptNo10x10Yes" <%= ChkIIF(wconceptNo10x10Yes="on","checked","") %> ><font color=red>wconceptǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="wconceptKeepSell" <%= ChkIIF(wconceptKeepSell="on","checked","") %> ><font color=red>�Ǹ�����</font> �ؾ��� ��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, ��۹���� �ù�(�Ϲ�) �ƴѰ�, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ɼǰ� 0�� �Ǹ��� ��ǰ�� ����(�ɼ� �������� 5�� ���� ����), �ɼǰ��� �ǸŰ� 100% �̻��� ��ǰ

<p />
<% end if %>
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
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('wconcept1010');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="wconceptSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP1)" onClick="wconceptSelectRegStepProcess(1);">&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP2)" onClick="wconceptSelectRegStepProcess(2);">&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP3)" onClick="wconceptSelectRegStepProcess(3);">&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP4)" onClick="wconceptSelectRegStepProcess(4);">&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP5)" onClick="wconceptSelectRegStepProcess(5);">&nbsp;&nbsp;
				<input class="button" type="button" value="���(STEP6)" onClick="wconceptSelectRegStepProcess(6);">&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" value="����" onClick="wconceptEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="����" onClick="wconceptSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="������" onClick="wconceptSelectContentEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="�̹���" onClick="wconceptSelectImageEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="�߰��̹���" onClick="wconceptSelectAddImageEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="�ɼ�" onClick="wconceptSelectOptionEditProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" value="���" onClick="wconceptSelectInfoCodeProcess();">&nbsp;&nbsp;
			<% End If %>
				<br><br>
				������ǰ ��ȸ :
				<input class="button" type="button" value="��ȸ" onClick="wconceptChkStatProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="��ǰ" onClick="wconceptChkItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="�ɼ�" onClick="wconceptChkOptProcess();">
				<br><br>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				�����ڵ� �˻� :
				<Select name="CommCD" id="CommCD" class="select">
					<option value="BRAND">�귣��</option>
					<option value="ORIGIN">������</option>
					<option value="CATEGORY">ī�װ�</option>
					<% If (session("ssBctID")="kjy8517") Then %>
					<option value="PRODUCT_NOTICE_INFO">ī�װ�ǰ��</option>
					<% End If %>
				</Select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="fnwconceptCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="wconceptSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= FormatNumber(oWconcept.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oWconcept.FTotalPage,0) %></b>
	</td>
	<td align="right">
		<input type="button" class="button" value="�����ޱ�" onClick="popXL()">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
<% If (xl <> "Y") Then %>
	<td width="50">�̹���</td>
<% End If %>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">wconcept�����<br>wconcept����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">wconcept<br>���ݹ��Ǹ�</td>
	<td width="70">wconcept<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oWconcept.FResultCount - 1 %>
<tr align="center" <% If oWconcept.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oWconcept.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oWconcept.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oWconcept.FItemList(i).FItemID %>','wconcept1010','<%=oWconcept.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oWconcept.FItemList(i).FItemID%>" target="_blank"><%= oWconcept.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oWconcept.FItemList(i).FwconceptStatCd <> 7 Then
	%>
		<br><%= oWconcept.FItemList(i).getwconceptStatName %>
	<%
			End If
			response.write oWconcept.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oWconcept.FItemList(i).FMakerid %> <%= oWconcept.FItemList(i).getDeliverytypeName %><br><%= oWconcept.FItemList(i).FItemName %></td>
	<td align="center"><%= oWconcept.FItemList(i).FRegdate %><br><%= oWconcept.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oWconcept.FItemList(i).FwconceptRegdate %><br><%= oWconcept.FItemList(i).FwconceptLastUpdate %></td>
	<td align="right">
		<% If oWconcept.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oWconcept.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oWconcept.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oWconcept.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oWconcept.FItemList(i).Fsellcash = 0 Then
		elseif (oWconcept.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oWconcept.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oWconcept.FItemList(i).FOrgSuplycash/oWconcept.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oWconcept.FItemList(i).Fbuycash/oWconcept.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oWconcept.FItemList(i).Fbuycash/oWconcept.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oWconcept.FItemList(i).IsSoldOut Then
			If oWconcept.FItemList(i).FSellyn = "N" Then
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
		If oWconcept.FItemList(i).FItemdiv = "06" OR oWconcept.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oWconcept.FItemList(i).FwconceptStatCd > 0) Then
			If Not IsNULL(oWconcept.FItemList(i).FwconceptPrice) Then
				If (oWconcept.FItemList(i).Mustprice <> oWconcept.FItemList(i).FwconceptPrice) Then
	%>
					<strong><%= formatNumber(oWconcept.FItemList(i).FwconceptPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oWconcept.FItemList(i).FwconceptPrice,0)&"<br>"
				End If

				If Not IsNULL(oWconcept.FItemList(i).FSpecialPrice) Then
					If (now() >= oWconcept.FItemList(i).FStartDate) And (now() <= oWconcept.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oWconcept.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oWconcept.FItemList(i).FSellyn="Y" and oWconcept.FItemList(i).FwconceptSellYn<>"Y") or (oWconcept.FItemList(i).FSellyn<>"Y" and oWconcept.FItemList(i).FwconceptSellYn="Y") Then
	%>
					<strong><%= oWconcept.FItemList(i).FwconceptSellYn %></strong>
	<%
				Else
					response.write oWconcept.FItemList(i).FwconceptSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oWconcept.FItemList(i).FwconceptGoodNo <> "" Then %>
			<a target="_blank" href="https://www.wconcept.co.kr/product/<%=oWconcept.FItemList(i).FwconceptGoodNo%>"><%=oWconcept.FItemList(i).FwconceptGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oWconcept.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oWconcept.FItemList(i).FItemID%>','0');"><%= oWconcept.FItemList(i).FoptionCnt %>:<%= oWconcept.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oWconcept.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= Chkiif(oWconcept.FItemList(i).FCateMapCnt > 0, "��Ī��(ī)", "<font color='darkred'>��Ī�ȵ�(ī)</font>") %><br />
		<%= Chkiif(oWconcept.FItemList(i).FBrandMapCnt > 0, "��Ī��(��)", "<font color='darkred'>��Ī�ȵ�(��)</font>") %><br />
	</td>
	<td align="center">
		<%= oWconcept.FItemList(i).FinfoDiv %>
		<%
		If (oWconcept.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oWconcept.FItemList(i).FlastErrStr) &"'>ERR:"& oWconcept.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% wconceptGoodNoArray = wconceptGoodNoArray & oWconcept.FItemList(i).FwconceptGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= wconceptGoodNoArray %></textarea>
	<button onclick="copyArr();">Copy</button>
	<script>
	function copyArr() {
		var tt = document.getElementById("itemidArr");
		tt.select();
		document.execCommand("copy");
	}
	</script>
<% End If %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oWconcept.HasPreScroll then %>
		<a href="javascript:goPage('<%= oWconcept.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oWconcept.StartScrollPage to oWconcept.FScrollCount + oWconcept.StartScrollPage - 1 %>
    		<% if i>oWconcept.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oWconcept.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value= <%= page %>>
	<input type="hidden" name="research" value= <%= research %>>
	<input type="hidden" name="itemid" value= <%= itemid %>>
	<input type="hidden" name="makerid" value= <%= makerid %>>
	<input type="hidden" name="itemname" value= <%= itemname %>>
	<input type="hidden" name="bestOrd" value= <%= bestOrd %>>
	<input type="hidden" name="bestOrdMall" value= <%= bestOrdMall %>>
	<input type="hidden" name="sellyn" value= <%= sellyn %>>
	<input type="hidden" name="limityn" value= <%= limityn %>>
	<input type="hidden" name="sailyn" value= <%= sailyn %>>
	<input type="hidden" name="onlyValidMargin" value= <%= onlyValidMargin %>>
	<input type="hidden" name="startMargin" value= <%= startMargin %>>
	<input type="hidden" name="endMargin" value= <%= endMargin %>>
	<input type="hidden" name="isMadeHand" value= <%= isMadeHand %>>
	<input type="hidden" name="isOption" value= <%= isOption %>>
	<input type="hidden" name="infoDiv" value= <%= infoDiv %>>
	<input type="hidden" name="morningJY" value= <%= morningJY %>>
	<input type="hidden" name="extsellyn" value= <%= extsellyn %>>
	<input type="hidden" name="wconceptGoodNo" value= <%= wconceptGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchBrand" value= <%= MatchBrand %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="wconceptYes10x10No" value= <%= wconceptYes10x10No %>>
	<input type="hidden" name="wconceptNo10x10Yes" value= <%= wconceptNo10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oWconcept = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
