<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/skstoa/skstoaCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, skstoaGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, setMargin, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, skstoaYes10x10No, skstoaNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, isextusing, cisextusing, rctsellcnt
Dim page, i, research, reglevel
Dim oSkstoa, xl, skstoaGoodNoArray, scheduleNotInItemid
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
skstoaGoodNo			= request("skstoaGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
skstoaYes10x10No		= request("skstoaYes10x10No")
skstoaNo10x10Yes		= request("skstoaNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
setMargin				= request("setMargin")
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
reglevel				= request("reglevel")
purchasetype			= request("purchasetype")
xl 						= request("xl")

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

'skstoa ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If skstoaGoodNo <> "" then
	Dim iA2, arrTemp2, arrskstoaGoodNo
	skstoaGoodNo = replace(skstoaGoodNo,",",chr(10))
	skstoaGoodNo = replace(skstoaGoodNo,chr(13),"")
	arrTemp2 = Split(skstoaGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrskstoaGoodNo = arrskstoaGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	skstoaGoodNo = left(arrskstoaGoodNo,len(arrskstoaGoodNo)-1)
End If

Set oSkstoa = new CSkstoa
	oSkstoa.FCurrPage				= page
	oSkstoa.FPageSize				= 100
	oSkstoa.FRectCDL				= request("cdl")
	oSkstoa.FRectCDM				= request("cdm")
	oSkstoa.FRectCDS				= request("cds")
	oSkstoa.FRectItemID				= itemid
	oSkstoa.FRectItemName			= itemname
	oSkstoa.FRectSellYn				= sellyn
	oSkstoa.FRectLimitYn			= limityn
	oSkstoa.FRectSailYn				= sailyn
	oSkstoa.FRectStartMargin		= startMargin
	oSkstoa.FRectEndMargin			= endMargin
	oSkstoa.FRectMakerid			= makerid
	oSkstoa.FRectSkstoaGoodNo		= skstoaGoodNo
	oSkstoa.FRectMatchCate			= MatchCate
	oSkstoa.FRectIsMadeHand			= isMadeHand
	oSkstoa.FRectIsOption			= isOption
	oSkstoa.FRectIsReged			= isReged
	oSkstoa.FRectNotinmakerid		= notinmakerid
	oSkstoa.FRectNotinitemid		= notinitemid
	oSkstoa.FRectExcTrans			= exctrans
	oSkstoa.FRectPriceOption		= priceOption
	oSkstoa.FRectIsSpecialPrice		= isSpecialPrice
	oSkstoa.FRectDeliverytype		= deliverytype
	oSkstoa.FRectMwdiv				= mwdiv
	oSkstoa.FRectSetMargin			= setMargin
	oSkstoa.FRectScheduleNotInItemid	= scheduleNotInItemid
	oSkstoa.FRectIsextusing			= isextusing
	oSkstoa.FRectCisextusing		= cisextusing
	oSkstoa.FRectRctsellcnt			= rctsellcnt
	oSkstoa.FRectReglevel			= reglevel

	oSkstoa.FRectExtNotReg			= ExtNotReg
	oSkstoa.FRectExpensive10x10		= expensive10x10
	oSkstoa.FRectdiffPrc			= diffPrc
	oSkstoa.FRectskstoaYes10x10No	= skstoaYes10x10No
	oSkstoa.FRectskstoaNo10x10Yes	= skstoaNo10x10Yes
	oSkstoa.FRectExtSellYn			= extsellyn
	oSkstoa.FRectInfoDiv			= infoDiv
	oSkstoa.FRectFailCntOverExcept	= ""
	oSkstoa.FRectFailCntExists		= failCntExists
	oSkstoa.FRectReqEdit			= reqEdit
	oSkstoa.FRectPurchasetype		= purchasetype
If (bestOrd = "on") Then
    oSkstoa.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oSkstoa.FRectOrdType = "BM"
End If

If isReged = "R" Then					'ǰ��ó����� ��ǰ���� ����Ʈ
	oSkstoa.getskstoareqExpireItemList
Else
	oSkstoa.getskstoaRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=skstoaList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=skstoa","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=skstoa','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=skstoa','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//���� ���� ��ǰ
function popMarginItemList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginItemList.asp?mallid=skstoa','popMarginItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=skstoa','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/skstoa/popskstoaCateList.asp","popCateskstoamanager","width=1200,height=600,scrollbars=yes,resizable=yes");
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
	if ((comp.name!="skstoaYes10x10No")&&(frm.skstoaYes10x10No.checked)){ frm.skstoaYes10x10No.checked=false }
	if ((comp.name!="skstoaNo10x10Yes")&&(frm.skstoaNo10x10Yes.checked)){ frm.skstoaNo10x10Yes.checked=false }
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

    if ((comp.name=="skstoaYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="skstoaNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.skstoaYes10x10No.checked){
            comp.form.skstoaYes10x10No.checked = false;
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
	if ((comp.name!="skstoaYes10x10No")&&(frm.skstoaYes10x10No.checked)){ frm.skstoaYes10x10No.checked=false }
	if ((comp.name!="skstoaNo10x10Yes")&&(frm.skstoaNo10x10Yes.checked)){ frm.skstoaNo10x10Yes.checked=false }
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
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//��ǰ ����
function skstoaDeleteProcess(){
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
    if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n11���� ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=skstoa&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// ���õ� ��ǰ �ϰ� ���
function skstoaSelectRegProcess() {
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

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� �ӽû�ǰ��ȸ
function skstoaSelectConfirmProcess() {
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

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ӽû�ǰ��ȸ �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CONFIRM";
		document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ���
function skstoaSelectSugiRegProcess(v) {
	var chkSel=0;
	var strAct;
	var apiAct;
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

	switch(v) {
		case "1": strAct="����";apiAct="REGAddItem";break;
		case "2": strAct="�����";apiAct="REGContent";break;
		case "3": strAct="��ǰ";apiAct="REGOpt";break;
		case "4": strAct="�̹���";apiAct="REGImg";break;
		case "5": strAct="���";apiAct="REGGosi";break;
		case "6": strAct="��������";apiAct="REGCert";break;
		case "7": strAct="���ο�û";apiAct="REGConfirm";break;
	}

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ�� '+strAct+' ��� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = apiAct;
		document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function skstoaSelectSugiEditProcess(v) {
	var chkSel=0;
	var strAct;
	var apiAct;
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

	switch(v) {
		case "1": strAct="�����";apiAct="EDITContent";break;
		case "2": strAct="�̹���";apiAct="EDITImage";break;
		case "3": strAct="���";apiAct="EDITGosi";break;
		case "4": strAct="��������";apiAct="EDITCert";break;
	}

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ�� '+strAct+' ���� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = apiAct;
		document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function skstoaSellYnProcess(chkYn) {
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
		case "X": strSell="�Ǹ�����(��������)";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
        document.frmSvArr.submit();
    }
}

//����
function skstoaSelectEditProcess() {
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

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
        document.frmSvArr.submit();
    }
}

//���� ����
function skstoaSelectPriceEditProcess() {
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

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
        document.frmSvArr.submit();
    }
}

//��� ����
function skstoaSelectQtyProcess() {
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

    if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� ���� �Ͻðڽ��ϱ�?\n\n��skstoa���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITQTY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
        document.frmSvArr.submit();
    }
}

//��ǰ �� ��ȸ
function skstoaSelectViewProcess() {
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

	if (confirm('skstoa�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/skstoa/actskstoaReq.asp";
        document.frmSvArr.submit();
    }
}

function popXL()
{
    frmXL.submit();
}

function popPassword(){
    var popwin = window.open("/admin/etc/skstoa/popPassword.asp","popPassword","width=400,height=50,scrollbars=yes,resizable=yes");
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
		<a href="https://scm.skstoa.com" target="_blank">skstoa Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 112644 | E112644 | Tenbyten10 ]</font>"
			End If
		%>
		<input type="button" class="button" value="��й�ȣ" onclick="popPassword();">
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		skstoa ��ǰ�ڵ� : <textarea rows="2" cols="20" name="skstoaGoodNo" id="itemid"><%= replace(replace(skstoaGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >skstoa ��Ͻ���
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >skstoa ��Ͽ���
			<option value="E" <%= CHkIIF(ExtNotReg="E","selected","") %> >skstoa ���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >skstoa �ݷ�
			<option value="X" <%= CHkIIF(ExtNotReg="X","selected","") %> >skstoa ����� �ӽ�(����)
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >skstoa ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >skstoa ��ϿϷ�(����)
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
		��Ϸ��� :
		<select name="reglevel" class="select">
			<option value="">����
			<% For i = 1 to 7 %>
				<option value="<%= i %>" <%= CHkIIF(CSTR(i) = CSTR(reglevel),"selected","") %> ><%= i %>
			<% Next %>
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>skstoa ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="skstoaYes10x10No" <%= ChkIIF(skstoaYes10x10No="on","checked","") %> ><font color=red>skstoa�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="skstoaNo10x10Yes" <%= ChkIIF(skstoaNo10x10Yes="on","checked","") %> ><font color=red>skstoaǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, ������ �ƴ� �� �ǸŰ� 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ֹ����۹��� ��ǰ, �ɼ��߰��ݾ� �ִ� ��ǰ<br />
* ��Ϸ��� : �����������(1), ��������(2), ��ǰ���(3), �̹���URL(4), �������(5), ��������(Optional)(6), ���ο�û(7)<br />
<p />
<form name="frmReg" method="post" action="skstoaItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();">&nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="���Ը�������(��ǰ)" onclick="popMarginItemList();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">&nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('skstoa');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="skstoaSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="��ȸ" onClick="skstoaSelectConfirmProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" value="����" onClick="skstoaSelectSugiRegProcess('1');">&nbsp;
				<input class="button" type="button" value="�����" onClick="skstoaSelectSugiRegProcess('2');">&nbsp;
				<input class="button" type="button" value="��ǰ" onClick="skstoaSelectSugiRegProcess('3');">&nbsp;
				<input class="button" type="button" value="�̹���" onClick="skstoaSelectSugiRegProcess('4');">&nbsp;
				<input class="button" type="button" value="���" onClick="skstoaSelectSugiRegProcess('5');">&nbsp;
				<input class="button" type="button" value="��������" onClick="skstoaSelectSugiRegProcess('6');">&nbsp;
				<input class="button" type="button" value="���ο�û" onClick="skstoaSelectSugiRegProcess('7');">&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSelPrice" value="����" onClick="skstoaSelectPriceEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="����" onClick="skstoaSelectEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="��ȸ" onClick="skstoaSelectViewProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" value="�����" onClick="skstoaSelectSugiEditProcess('1');">&nbsp;&nbsp;
				<input class="button" type="button" value="�̹���" onClick="skstoaSelectSugiEditProcess('2');">&nbsp;&nbsp;
				<input class="button" type="button" value="���" onClick="skstoaSelectSugiEditProcess('3');">&nbsp;&nbsp;
				<input class="button" type="button" value="��������" onClick="skstoaSelectSugiEditProcess('4');">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelOption" value="���" onClick="skstoaSelectQtyProcess();">&nbsp;&nbsp;
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<!-- <option value="Y">�Ǹ���</option> -->
					<option value="X">�Ǹ�����(��������))</option><!-- ���������ϸ� ���� ���� �� �� ���� / ���� ������ �� �ȴٰ� �� -->
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="skstoaSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="grpVal" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= FormatNumber(oSkstoa.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSkstoa.FTotalPage,0) %></b>
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
	<td width="140">skstoa�����<br>skstoa����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">skstoa<br>���ݹ��Ǹ�</td>
	<td width="70">skstoa<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="50">���븶��</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oSkstoa.FResultCount - 1 %>
<tr align="center" <% If oSkstoa.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oSkstoa.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oSkstoa.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oSkstoa.FItemList(i).FItemID %>','skstoa','<%=oSkstoa.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oSkstoa.FItemList(i).FItemID%>" target="_blank"><%= oSkstoa.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			response.write oSkstoa.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oSkstoa.FItemList(i).FMakerid %> <%= oSkstoa.FItemList(i).getDeliverytypeName %><br><%= oSkstoa.FItemList(i).FItemName %></td>
	<td align="center"><%= oSkstoa.FItemList(i).FRegdate %><br><%= oSkstoa.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oSkstoa.FItemList(i).FSkstoaRegdate %><br><%= oSkstoa.FItemList(i).FSkstoaLastUpdate %></td>
	<td align="right">
		<% If oSkstoa.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oSkstoa.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oSkstoa.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oSkstoa.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oSkstoa.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oSkstoa.FItemList(i).Fbuycash/oSkstoa.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oSkstoa.FItemList(i).Fbuycash/oSkstoa.FItemList(i).FSkstoaPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oSkstoa.FItemList(i).Fbuycash/oSkstoa.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oSkstoa.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oSkstoa.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oSkstoa.FItemList(i).FOrgSuplycash/oSkstoa.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oSkstoa.FItemList(i).Fbuycash/oSkstoa.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oSkstoa.FItemList(i).Fbuycash/oSkstoa.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oSkstoa.FItemList(i).IsSoldOut Then
			If oSkstoa.FItemList(i).FSellyn = "N" Then
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
		If oSkstoa.FItemList(i).FItemdiv = "06" OR oSkstoa.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oSkstoa.FItemList(i).FSkstoaStatCd > 0) Then
			If Not IsNULL(oSkstoa.FItemList(i).FSkstoaPrice) Then
				If (oSkstoa.FItemList(i).Mustprice <> oSkstoa.FItemList(i).FSkstoaPrice) Then
	%>
					<strong><%= formatNumber(oSkstoa.FItemList(i).FSkstoaPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oSkstoa.FItemList(i).FSkstoaPrice,0)&"<br>"
				End If

				If Not IsNULL(oSkstoa.FItemList(i).FSpecialPrice) Then
					If (now() >= oSkstoa.FItemList(i).FStartDate) And (now() <= oSkstoa.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oSkstoa.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oSkstoa.FItemList(i).FSellyn="Y" and oSkstoa.FItemList(i).FSkstoaSellYn<>"Y") or (oSkstoa.FItemList(i).FSellyn<>"Y" and oSkstoa.FItemList(i).FSkstoaSellYn="Y") Then
	%>
					<strong><%= oSkstoa.FItemList(i).FSkstoaSellYn %></strong>
	<%
				Else
					response.write oSkstoa.FItemList(i).FSkstoaSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
    <%
    	'#�ǻ�ǰ��ȣ
    	if (oSkstoa.FItemList(i).FSkstoaGoodNo <> "") then
        	Response.Write "<a target='_blank' href='http://www.skstoa.com/display/goods/"&oSkstoa.FItemList(i).FSkstoaGoodNo&"'>"&oSkstoa.FItemList(i).FSkstoaGoodNo&"</a>"
		else
			'#�ӽû�ǰ��ȣ
			if oSkstoa.FItemList(i).FSkstoaTmpGoodNo <> "" then
				Response.Write oSkstoa.FItemList(i).getskstoaStatName & "<br>(" & oSkstoa.FItemList(i).FSkstoaTmpGoodNo & ")"
			end if
		end if
	%>
	</td>
	<td align="center"><%= oSkstoa.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oSkstoa.FItemList(i).FItemID%>','0');"><%= oSkstoa.FItemList(i).FoptionCnt %>:<%= oSkstoa.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oSkstoa.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= oSkstoa.FItemList(i).FSetMargin %></td>
	<td align="center">
	<%
		If oSkstoa.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If

		If oSkstoa.FItemList(i).FReglevel < 5 Then
			response.write " <br />���� : " & oSkstoa.FItemList(i).FReglevel
		End If
	%>
	</td>
	<td align="center">
		<%= oSkstoa.FItemList(i).FinfoDiv %>
		<%
		If (oSkstoa.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oSkstoa.FItemList(i).FlastErrStr) &"'>ERR:"& oSkstoa.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% skstoaGoodNoArray = skstoaGoodNoArray & oSkstoa.FItemList(i).FSkstoaGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= skstoaGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oSkstoa.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSkstoa.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSkstoa.StartScrollPage to oSkstoa.FScrollCount + oSkstoa.StartScrollPage - 1 %>
    		<% if i>oSkstoa.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSkstoa.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
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
	<input type="hidden" name="skstoaGoodNo" value= <%= skstoaGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="skstoaYes10x10No" value= <%= skstoaYes10x10No %>>
	<input type="hidden" name="skstoaNo10x10Yes" value= <%= skstoaNo10x10Yes %>>
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
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% Set oSkstoa = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->