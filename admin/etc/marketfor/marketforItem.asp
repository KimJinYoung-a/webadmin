<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/marketfor/marketforCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, marketforGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, marketforYes10x10No, marketforNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, isextusing, cisextusing, rctsellcnt
Dim page, i, research
Dim oMarketfor, xl, marketforGoodNoArray
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
marketforGoodNo			= request("marketforGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
marketforYes10x10No		= request("marketforYes10x10No")
marketforNo10x10Yes		= request("marketforNo10x10Yes")
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

'marketfor ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If marketforGoodNo <> "" then
	Dim iA2, arrTemp2, arrmarketforGoodNo
	marketforGoodNo = replace(marketforGoodNo,",",chr(10))
	marketforGoodNo = replace(marketforGoodNo,chr(13),"")
	arrTemp2 = Split(marketforGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrmarketforGoodNo = arrmarketforGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	marketforGoodNo = left(arrmarketforGoodNo,len(arrmarketforGoodNo)-1)
End If

Set oMarketfor = new Cmarketfor
	oMarketfor.FCurrPage				= page
	oMarketfor.FPageSize				= 100
	oMarketfor.FRectCDL					= request("cdl")
	oMarketfor.FRectCDM					= request("cdm")
	oMarketfor.FRectCDS					= request("cds")
	oMarketfor.FRectItemID				= itemid
	oMarketfor.FRectItemName			= itemname
	oMarketfor.FRectSellYn				= sellyn
	oMarketfor.FRectLimitYn				= limityn
	oMarketfor.FRectSailYn				= sailyn
	oMarketfor.FRectStartMargin			= startMargin
	oMarketfor.FRectEndMargin			= endMargin
	oMarketfor.FRectMakerid				= makerid
	oMarketfor.FRectmarketforGoodNo		= marketforGoodNo
	oMarketfor.FRectMatchCate			= MatchCate
	oMarketfor.FRectIsMadeHand			= isMadeHand
	oMarketfor.FRectIsOption			= isOption
	oMarketfor.FRectIsReged				= isReged
	oMarketfor.FRectNotinmakerid		= notinmakerid
	oMarketfor.FRectNotinitemid			= notinitemid
	oMarketfor.FRectExcTrans			= exctrans
	oMarketfor.FRectPriceOption			= priceOption
	oMarketfor.FRectIsSpecialPrice		= isSpecialPrice
	oMarketfor.FRectDeliverytype		= deliverytype
	oMarketfor.FRectMwdiv				= mwdiv
	oMarketfor.FRectIsextusing			= isextusing
	oMarketfor.FRectCisextusing			= cisextusing
	oMarketfor.FRectRctsellcnt			= rctsellcnt


	oMarketfor.FRectExtNotReg			= ExtNotReg
	oMarketfor.FRectExpensive10x10		= expensive10x10
	oMarketfor.FRectdiffPrc				= diffPrc
	oMarketfor.FRectmarketforYes10x10No	= marketforYes10x10No
	oMarketfor.FRectmarketforNo10x10Yes	= marketforNo10x10Yes
	oMarketfor.FRectExtSellYn			= extsellyn
	oMarketfor.FRectInfoDiv				= infoDiv
	oMarketfor.FRectFailCntOverExcept	= ""
	oMarketfor.FRectFailCntExists		= failCntExists
	oMarketfor.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oMarketfor.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oMarketfor.FRectOrdType = "BM"
End If

If isReged = "R" Then					'ǰ��ó����� ��ǰ���� ����Ʈ
	oMarketfor.getMarketforreqExpireItemList
Else
	oMarketfor.getMarketforRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=marketforList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=marketfor","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=marketfor','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=marketfor','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/marketfor/popMarketforCateList.asp","popCatemarketformanager","width=1200,height=600,scrollbars=yes,resizable=yes");
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
	if ((comp.name!="marketforYes10x10No")&&(frm.marketforYes10x10No.checked)){ frm.marketforYes10x10No.checked=false }
	if ((comp.name!="marketforNo10x10Yes")&&(frm.marketforNo10x10Yes.checked)){ frm.marketforNo10x10Yes.checked=false }
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

    if ((comp.name=="marketforYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="marketforNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.marketforYes10x10No.checked){
            comp.form.marketforYes10x10No.checked = false;
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
	if ((comp.name!="marketforYes10x10No")&&(frm.marketforYes10x10No.checked)){ frm.marketforYes10x10No.checked=false }
	if ((comp.name!="marketforNo10x10Yes")&&(frm.marketforNo10x10Yes.checked)){ frm.marketforNo10x10Yes.checked=false }
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
function marketforDeleteProcess(){
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
			document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=market_for&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// ���õ� ��ǰ �ϰ� ���
function marketforSelectRegProcess() {
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

    if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��marketfor���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function marketforSellYnProcess(chkYn) {
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
		case "X": strSell="�Ǹ�����(����)";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}

//����
function marketforSelectEditProcess() {
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

    if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?\n\n��marketfor���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}

//���� ����
function marketforSelectPriceEditProcess() {
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

    if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��marketfor���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}

//��� ����
function marketforSelectQtyProcess() {
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

    if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� ���� �Ͻðڽ��ϱ�?\n\n��marketfor���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "QTY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ �� ��ȸ
function marketforSelectViewProcess() {
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

	if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ ���� ����
function marketforSelectOptStatProcess() {
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

	if (confirm('marketfor�� �����Ͻ� ' + chkSel + '�� ��ǰ ���� ���� �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/marketfor/actmarketforReq.asp"
        document.frmSvArr.submit();
    }
}
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
		<a href="https://store.marketfor.com/cm/main/login_SO.wsp" target="_blank">�Ե�On Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ LD304013 | store101010* ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		marketfor ��ǰ�ڵ� : <textarea rows="2" cols="20" name="marketforGoodNo" id="itemid"><%= replace(replace(marketforGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >marketfor ��Ͻ���
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >marketfor ��Ͽ���
			<option value="E" <%= CHkIIF(ExtNotReg="E","selected","") %> >marketfor ���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >marketfor �ݷ�
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >marketfor ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >marketfor ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>marketfor ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="marketforYes10x10No" <%= ChkIIF(marketforYes10x10No="on","checked","") %> ><font color=red>marketfor�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="marketforNo10x10Yes" <%= ChkIIF(marketforNo10x10Yes="on","checked","") %> ><font color=red>marketforǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
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

<p />
<form name="frmReg" method="post" action="marketforItem.asp" style="margin:0px;">
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
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('market_for');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="marketforSelectRegProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="����" onClick="marketforSelectEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="����" onClick="marketforSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelOption" value="���" onClick="marketforSelectQtyProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="��ȸ" onClick="marketforSelectViewProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="�ɼǻ���" onClick="marketforSelectOptStatProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="��ǰ����" onClick="marketforDeleteProcess();" style=font-weight:bold>
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
					<option value="X">�Ǹ�����(����)</option><!-- �����ϸ� ���� ���� �� �� ���� -->
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="marketforSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="16">
		�˻���� : <b><%= FormatNumber(oMarketfor.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMarketfor.FTotalPage,0) %></b>
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
	<td width="140">marketfor�����<br>marketfor����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">marketfor<br>���ݹ��Ǹ�</td>
	<td width="70">marketfor<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oMarketfor.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oMarketfor.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oMarketfor.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oMarketfor.FItemList(i).FItemID %>','market_for','<%=oMarketfor.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oMarketfor.FItemList(i).FItemID%>" target="_blank"><%= oMarketfor.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oMarketfor.FItemList(i).FmarketforStatcd <> 7 Then
	%>
		<br><%= oMarketfor.FItemList(i).getmarketforStatName %>
	<%
			End If
			response.write oMarketfor.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oMarketfor.FItemList(i).FMakerid %> <%= oMarketfor.FItemList(i).getDeliverytypeName %><br><%= oMarketfor.FItemList(i).FItemName %></td>
	<td align="center"><%= oMarketfor.FItemList(i).FRegdate %><br><%= oMarketfor.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oMarketfor.FItemList(i).FmarketforRegdate %><br><%= oMarketfor.FItemList(i).FmarketforLastUpdate %></td>
	<td align="right">
		<% If oMarketfor.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oMarketfor.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oMarketfor.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oMarketfor.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oMarketfor.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oMarketfor.FItemList(i).Fbuycash/oMarketfor.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oMarketfor.FItemList(i).Fbuycash/oMarketfor.FItemList(i).FmarketforPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oMarketfor.FItemList(i).Fbuycash/oMarketfor.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oMarketfor.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oMarketfor.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oMarketfor.FItemList(i).FOrgSuplycash/oMarketfor.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oMarketfor.FItemList(i).Fbuycash/oMarketfor.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oMarketfor.FItemList(i).Fbuycash/oMarketfor.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oMarketfor.FItemList(i).IsSoldOut Then
			If oMarketfor.FItemList(i).FSellyn = "N" Then
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
		If oMarketfor.FItemList(i).FItemdiv = "06" OR oMarketfor.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oMarketfor.FItemList(i).FmarketforStatCd > 0) Then
			If Not IsNULL(oMarketfor.FItemList(i).FmarketforPrice) Then
				If (oMarketfor.FItemList(i).Mustprice <> oMarketfor.FItemList(i).FmarketforPrice) Then
	%>
					<strong><%= formatNumber(oMarketfor.FItemList(i).FmarketforPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oMarketfor.FItemList(i).FmarketforPrice,0)&"<br>"
				End If

				If Not IsNULL(oMarketfor.FItemList(i).FSpecialPrice) Then
					If (now() >= oMarketfor.FItemList(i).FStartDate) And (now() <= oMarketfor.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oMarketfor.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oMarketfor.FItemList(i).FSellyn="Y" and oMarketfor.FItemList(i).FmarketforSellYn<>"Y") or (oMarketfor.FItemList(i).FSellyn<>"Y" and oMarketfor.FItemList(i).FmarketforSellYn="Y") Then
	%>
					<strong><%= oMarketfor.FItemList(i).FmarketforSellYn %></strong>
	<%
				Else
					response.write oMarketfor.FItemList(i).FmarketforSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oMarketfor.FItemList(i).FmarketforGoodNo <> "" Then %>
			<a target="_blank" href="https://www.marketfor.com/p/product/<%=oMarketfor.FItemList(i).FmarketforGoodNo%>"><%=oMarketfor.FItemList(i).FmarketforGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oMarketfor.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oMarketfor.FItemList(i).FItemID%>','0');"><%= oMarketfor.FItemList(i).FoptionCnt %>:<%= oMarketfor.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oMarketfor.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oMarketfor.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oMarketfor.FItemList(i).FinfoDiv %>
		<%
		If (oMarketfor.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oMarketfor.FItemList(i).FlastErrStr) &"'>ERR:"& oMarketfor.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% marketforGoodNoArray = marketforGoodNoArray & oMarketfor.FItemList(i).FmarketforGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= marketforGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oMarketfor.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMarketfor.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oMarketfor.StartScrollPage to oMarketfor.FScrollCount + oMarketfor.StartScrollPage - 1 %>
    		<% if i>oMarketfor.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oMarketfor.HasNextScroll then %>
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
	<input type="hidden" name="marketforGoodNo" value= <%= marketforGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="marketforYes10x10No" value= <%= marketforYes10x10No %>>
	<input type="hidden" name="marketforNo10x10Yes" value= <%= marketforNo10x10Yes %>>
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
<% Set oMarketfor = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->