<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/gmarket/gmarketcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, addOptErr, isSpecialPrice
Dim bestOrdMall, gmarketGoodNo, g9GoodNo, extsellyn, ExtNotReg, isReged, MatchCate, MatchBrand, optAddPrcRegTypeNone, notinmakerid, notinitemid, MatchG9, sellpriceChk, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, gmarketYes10x10No, gmarketNo10x10Yes, gmarketKeepSell, reqEdit, reqExpire, failCntExists, priceOption, isextusing, scheduleNotInItemid
Dim page, i, research, GmarketGoodNoArray, cisextusing, rctsellcnt
Dim oGmarket, xl
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
gmarketGoodNo			= request("gmarketGoodNo")
g9GoodNo				= request("g9GoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchBrand				= request("MatchBrand")
MatchG9					= request("MatchG9")
sellpriceChk			= request("sellpriceChk")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
addOptErr				= request("addOptErr")
gmarketYes10x10No		= request("gmarketYes10x10No")
gmarketNo10x10Yes		= request("gmarketNo10x10Yes")
gmarketKeepSell			= request("gmarketKeepSell")
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
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
purchasetype			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchBrand = ""
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

'Gmarket ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If gmarketGoodNo <> "" then
	Dim iA2, arrTemp2, arrgmarketGoodNo
	gmarketGoodNo = replace(gmarketGoodNo,",",chr(10))
	gmarketGoodNo = replace(gmarketGoodNo,chr(13),"")
	arrTemp2 = Split(gmarketGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrgmarketGoodNo = arrgmarketGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	gmarketGoodNo = left(arrgmarketGoodNo,len(arrgmarketGoodNo)-1)
End If

'G9 ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If g9GoodNo <> "" then
	Dim iA3, arrTemp3, arrg9GoodNo
	g9GoodNo = replace(g9GoodNo,",",chr(10))
	g9GoodNo = replace(g9GoodNo,chr(13),"")
	arrTemp3 = Split(g9GoodNo,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			arrg9GoodNo = arrg9GoodNo& "'"& trim(arrTemp3(iA3)) & "',"
		End If
		iA3 = iA3 + 1
	Loop
	g9GoodNo = left(arrg9GoodNo,len(arrg9GoodNo)-1)
End If

Set oGmarket = new CGmarket
	oGmarket.FCurrPage					= page
	oGmarket.FPageSize					= 100
	oGmarket.FRectCDL					= request("cdl")
	oGmarket.FRectCDM					= request("cdm")
	oGmarket.FRectCDS					= request("cds")
	oGmarket.FRectItemID				= itemid
	oGmarket.FRectItemName				= itemname
	oGmarket.FRectSellYn				= sellyn
	oGmarket.FRectLimitYn				= limityn
	oGmarket.FRectSailYn				= sailyn
'	oGmarket.FRectonlyValidMargin		= onlyValidMargin
	oGmarket.FRectStartMargin			= startMargin
	oGmarket.FRectEndMargin				= endMargin
	oGmarket.FRectMakerid				= makerid
	oGmarket.FRectGmarketGoodNo			= gmarketGoodNo
	oGmarket.FRectG9GoodNo				= g9GoodNo
	oGmarket.FRectMatchCate				= MatchCate
	oGmarket.FRectMatchBrand			= MatchBrand
	oGmarket.FRectMatchG9				= MatchG9
	oGmarket.FRectSellpriceChk			= sellpriceChk
	oGmarket.FRectIsMadeHand			= isMadeHand
	oGmarket.FRectIsOption				= isOption
	oGmarket.FRectIsReged				= isReged
	oGmarket.FRectNotinmakerid			= notinmakerid
	oGmarket.FRectNotinitemid			= notinitemid
	oGmarket.FRectExcTrans				= exctrans
	oGmarket.FRectPriceOption			= priceOption
	oGmarket.FRectIsSpecialPrice     	= isSpecialPrice
	oGmarket.FRectAddOptErr				= addOptErr
	oGmarket.FRectDeliverytype			= deliverytype
	oGmarket.FRectMwdiv					= mwdiv
	oGmarket.FRectScheduleNotInItemid	= scheduleNotInItemid
	oGmarket.FRectIsextusing			= isextusing
	oGmarket.FRectCisextusing			= cisextusing
	oGmarket.FRectRctsellcnt			= rctsellcnt

	oGmarket.FRectExtNotReg				= ExtNotReg
	oGmarket.FRectExpensive10x10		= expensive10x10
	oGmarket.FRectdiffPrc				= diffPrc
	oGmarket.FRectGmarketYes10x10No		= gmarketYes10x10No
	oGmarket.FRectGmarketNo10x10Yes		= gmarketNo10x10Yes
	oGmarket.FRectGmarketKeepSell		= gmarketKeepSell
	oGmarket.FRectExtSellYn				= extsellyn
	oGmarket.FRectInfoDiv				= infoDiv
	oGmarket.FRectFailCntOverExcept		= ""
	oGmarket.FRectFailCntExists			= failCntExists
	oGmarket.FRectReqEdit				= reqEdit
	oGmarket.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oGmarket.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oGmarket.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oGmarket.getGmarketreqExpireItemList
Else
	oGmarket.getGmarketRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=gmarket1010List"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=gmarket1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=gmarket1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=gmarket1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="gmarketKeepSell")&&(frm.gmarketKeepSell.checked)){ frm.gmarketKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gmarketYes10x10No")&&(frm.gmarketYes10x10No.checked)){ frm.gmarketYes10x10No.checked=false }
	if ((comp.name!="gmarketNo10x10Yes")&&(frm.gmarketNo10x10Yes.checked)){ frm.gmarketNo10x10Yes.checked=false }
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

    if ((comp.name=="gmarketYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="gmarketNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.gmarketYes10x10No.checked){
            comp.form.gmarketYes10x10No.checked = false;
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

    if ((comp.name=="gmarketKeepSell")&&(comp.checked)){
        if (comp.form.gmarketYes10x10No.checked){
            comp.form.gmarketYes10x10No.checked = false;
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

	if (comp.name=="addOptErr"){
		if (comp.checked){
			document.getElementById("AR").checked=false;
			document.getElementById("NR").checked=false;
			document.getElementById("RR").checked=false;
			document.getElementById("QR").checked=false;
			comp.form.ExtNotReg.disabled = false;
			comp.form.priceOption.value = "Y";
			comp.form.ExtNotReg.value="W"
			comp.form.sellyn.value="A";
			comp.form.onlyValidMargin.value="";
			comp.form.extsellyn.value = "N";
		}
	}

	if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
	if ((comp.name!="gmarketKeepSell")&&(frm.gmarketKeepSell.checked)){ frm.gmarketKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="gmarketYes10x10No")&&(frm.gmarketYes10x10No.checked)){ frm.gmarketYes10x10No.checked=false }
	if ((comp.name!="gmarketNo10x10Yes")&&(frm.gmarketNo10x10Yes.checked)){ frm.gmarketNo10x10Yes.checked=false }
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
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/gmarket/popgmarketcateList.asp","popCateGmarketmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//�귣�� ����
function pop_BrandManager() {
	var pCM2 = window.open("/admin/etc/gmarket/popgmarketbrandList.asp","popBrandGmarketmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//�귣�� ����
function pop_AddMakerBrand(){
	var pCM3 = window.open("/admin/etc/gmarket/popgmarketBrand.asp","popAddMakerBrand","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM3.focus();
}
//����� ����
function pop_AddAdressBook(){
	var pCM4 = window.open("/admin/etc/gmarket/popgmarketAddress.asp","popAddMakerAddress","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM3.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=gmarket1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ǿ���
function G9SpecialList(){
	var popwin=window.open('/admin/etc/gmarket/g9SpecialItem.asp','specialItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ���õ� ��ǰ �⺻���� ���
function GmarketSelectRegProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �⺻������ ��� �Ͻðڽ��ϱ�?\n\nG�����ڵ带 ���Ϲޱ� ���� �⺻���� ����Դϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddItem";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ɼ����� ���
function GmarketSelectOPTProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ɼ������� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ������� ���
function GmarketSelectInfoCdRegProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ������ø� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGInfoCd";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ����, ��� ���
function GmarketSelectPriceRegProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ����&��� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGPrice";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �⺻���� + �ɼ� + ������� ���
function GmarketREGProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �Ǹ��� ����
function GmarketOnSaleEditProcess() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ���¸� �Ǹ������� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOnSale";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �Ǹ��� ����
function GmarketCategory() {
    if (confirm('Gmarket�� ī�װ� get? ')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CATE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �Ǹ��� ����2
function GmarketOnSaleEdit2Process() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ���¸� �Ǹ������� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOnSale2";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ OnSale + �ɼ� ���
function GmarketREG2Process() {
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG2";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function GmarketSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� ����
function GmarketEditInfoProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ������ ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditInfo";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� ���� + ��ǰ�� ����
function GmarketEditPolicyProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ��ǰ����+��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITPOLICY";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//�ɼ� ����
function GmarketEditPriceOPTProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� �ɼ��� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//�̹��� ����
function GmarketEditImgProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� �̹����� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditImg";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//���� ����
function GmarketEditSafeProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� ������ ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditCert";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� + ���� + �ɼ����� ����
function GmarketEditProcess(){
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

    if (confirm('Gmarket�� �����Ͻ� ' + chkSel + '�� �⺻���� + ���� + �ɼ� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}
//G9 ��ǰ ���
function GmarketG9SelectRegProcess(){
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

    if (confirm('G9�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGG9";
        document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ ����
function GmarketDeleteProcess(){
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
    if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n������ ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
			document.frmSvArr.submit();
		}
    }
}









//�����ڵ� �˻�
function fngmarketCommCDSubmit() {
	var ccd;
	var goodsGrpCd;
	ccd = document.getElementById('CommCD').value;
	//goodsGrpCd = $("#goodsGrpCd option:selected").val();

	goodsGrpCd = $("#goodsGrpCd").val();
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		
		xLink.location.href = "/admin/etc/gmarket/actgmarketReq.asp?cmdparam=ebayCommonCode&CommCD="+ccd+"&goodsGrpCd="+goodsGrpCd;
	}
}
function jsByValue(s){
	if((s == "brand") || (s == "maker") || (s =="placepolicy" || s == "infocodedtl" || s == "mastercode" || s == "sitecode" || s == "addon")) {
		$("#goodsGrpCd_span").show();
	}else{
		$("#goodsGrpCd_span").hide();
	}
}












//����, ��ȿ�� ��� �˾�
function popgmarketDate(iitemid){
    var pdate = window.open("/admin/etc/gmarket/popGmarketDate.asp?itemid="+iitemid+'&mallid=gmarket',"popgmarketDate","width=500,height=200,scrollbars=yes,resizable=yes");
	pdate.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=gmarket1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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
		<a href="http://www.esmplus.com/Home/Home" target="_blank">G���� Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ 10x10store | Cube1010!* ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		G���� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="gmarketGoodNo" id="itemid"><%= replace(replace(gmarketGoodNo,",",chr(10)), "'", "")%></textarea>
		&nbsp;
		G9 ��ǰ�ڵ� : <textarea rows="2" cols="20" name="g9GoodNo" id="itemid"><%= replace(replace(g9GoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >G���� ��ϼ���_OnSale��
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >G���� ���۽õ� �� ����
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >G���� ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >G���� ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>G���� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="addOptErr" <%= ChkIIF(addOptErr="on","checked","") %> >�߰��ݾ׵�Ͽ���</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketYes10x10No" <%= ChkIIF(gmarketYes10x10No="on","checked","") %> ><font color=red>G�����Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketNo10x10Yes" <%= ChkIIF(gmarketNo10x10Yes="on","checked","") %> ><font color=red>G����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="gmarketKeepSell" <%= ChkIIF(gmarketKeepSell="on","checked","") %> ><font color=red>�Ǹ�����</font> �ؾ��� ��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>

<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : ��ǰ���� IFRAME TAG ����� ��ǰ, �ɼǰ��� �ǸŰ� 50% �̻��� ��ǰ, �ɼǼ� 100�� �ʰ� ��ǰ, �ɼǰ� 0�� �Ǹ��� ��ǰ�� ����(�ɼ� �������� 5�� ���� ����)<br />

<p />

<form name="frmReg" method="post" action="gmarketitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('gmarket1010');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ī�װ�" onclick="pop_CateManager();"> &nbsp;
				<!--
				<input class="button" type="button" value="�귣��" onclick="pop_BrandManager();">
				<input class="button" type="button" value="�����" onclick="pop_AddAdressBook();"> &nbsp;
				<input class="button" type="button" value="�귣��" onclick="pop_AddMakerBrand();">
				 -->
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
				<input class="button" type="button" id="btnRegSel" value="�⺻����" onClick="GmarketSelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnInfocdSel" value="��ǰ���" onClick="GmarketSelectInfoCdRegProcess();">&nbsp;&nbsp;
				<!-- <input class="button" type="button" id="btnPriceSel" value="����/���" onClick="GmarketSelectPriceRegProcess();">&nbsp;&nbsp; -->
				<input class="button" type="button" id="btnOPTSel" value="�ɼ�����" onClick="GmarketSelectOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnREG" value="�⺻+���+�ɼ�" onClick="GmarketREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnOnSale" value="OnSale����" onClick="GmarketOnSaleEditProcess();" style=color:red;font-weight:bold>
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
					<input class="button" type="button" id="btnCate" value="ī�װ�" onClick="GmarketCategory();">
				<% End If %>
				<br><br>
				�߰��ݾ� ����  :
				<input class="button" type="button" id="btnEditSale" value="OnSale����" onClick="GmarketOnSaleEdit2Process();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOPTSel" value="�ɼ�����" onClick="GmarketSelectOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnREG" value="OnSale+�ɼ�" onClick="GmarketREG2Process();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditInfo" value="�⺻����" onClick="GmarketEditInfoProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEditPrice" value="����+�ɼ�" onClick="GmarketEditPriceOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEditImg" value="�̹���" onClick="GmarketEditImgProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEditSafe" value="����" onClick="GmarketEditSafeProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEdit" value="�⺻+����+�ɼ�" onClick="GmarketEditProcess();" style=color:blue;font-weight:bold>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="nys1006") OR (session("ssBctID")="z0516") OR (session("ssBctID")="hrkang97") Then %>
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnOPdit" value="�⺻+��ǰ��" onClick="GmarketEditPolicyProcess();">
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="��ǰ����" onClick="GmarketDeleteProcess();" style=font-weight:bold>
			<% End If %>
				<br><br>
				G9��ǰ ���&nbsp;&nbsp; :
				<input class="button" type="button" id="btnRegG9Sel" value="���" onClick="GmarketG9SelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" value="������ǿ���" onclick="G9SpecialList();">



			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD" onChange="jsByValue(this.value);">
					<option value="">- Choice -</option>
					<option value="category">ESMī�װ�</option>
					<option value="brand">�귣������� �귣��ã��</option>
					<option value="maker">����������� �귣��ã��</option>
					<option value="address">�Ǹ����ּҷ�</option>
					<option value="locaddress">���������</option>
					<option value="placepolicy">�������� ������ۺ��ڵ�ã��</option>
					<option value="dispatchpolicy">�߼���å��ȸ</option>
					<option value="parcel">�ù���ڵ�</option>
					<option value="origin">�������ڵ�</option>
					<option value="infocode">������û�ǰ��</option>
					<option value="infocodedtl">������û�</option>
					<option value="mastercode">�������ڵ���ȸ(G�����ڵ��)</option>
					<option value="sitecode">G�����ڵ���ȸ(�������ڵ��)</option>
					<option value="addon">�߰�������ȸ</option>
				</select>
				<span id="goodsGrpCd_span" style="display:none;">
					<input type="text" name="goodsGrpCd" id="goodsGrpCd">
				</span>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="fngmarketCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ�</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="GmarketSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<!-- ����Ʈ ���� -->
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oGmarket.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oGmarket.FTotalPage,0) %></b>
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
	<td width="140">Gmarket�����<br>Gmarket����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Gmarket<br>���ݹ��Ǹ�</td>
	<td width="70">Gmarket<br>��ǰ��ȣ</td>
	<td width="70">G9<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
	<td width="100">��|��|��<br>�Ǹŷ� ������</td>
</tr>
<% For i=0 to oGmarket.FResultCount - 1 %>
<tr align="center" <% If oGmarket.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oGmarket.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oGmarket.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oGmarket.FItemList(i).FItemID %>','gmarket1010','<%=oGmarket.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oGmarket.FItemList(i).FItemID%>" target="_blank"><%= oGmarket.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oGmarket.FItemList(i).FGmarketStatcd <> 7 Then
	%>
		<br><%= oGmarket.FItemList(i).getGmarketStatName %>
	<%
			End If
			response.write oGmarket.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oGmarket.FItemList(i).FMakerid %> <%= oGmarket.FItemList(i).getDeliverytypeName %><br><%= oGmarket.FItemList(i).FItemName %></td>
	<td align="center"><%= oGmarket.FItemList(i).FRegdate %><br><%= oGmarket.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oGmarket.FItemList(i).FGmarketRegdate %><br><%= oGmarket.FItemList(i).FGmarketLastUpdate %></td>
	<td align="right">
		<% If oGmarket.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oGmarket.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oGmarket.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oGmarket.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oGmarket.FItemList(i).Fsellcash = 0 Then
		elseif (oGmarket.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oGmarket.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oGmarket.FItemList(i).FOrgSuplycash/oGmarket.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oGmarket.FItemList(i).Fbuycash/oGmarket.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oGmarket.FItemList(i).Fbuycash/oGmarket.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oGmarket.FItemList(i).IsSoldOut Then
			If oGmarket.FItemList(i).FSellyn = "N" Then
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
		If oGmarket.FItemList(i).FItemdiv = "06" OR oGmarket.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oGmarket.FItemList(i).FGmarketStatCd > 0) Then
			If Not IsNULL(oGmarket.FItemList(i).FGmarketPrice) Then
				If (oGmarket.FItemList(i).Mustprice <> oGmarket.FItemList(i).FGmarketPrice) Then
	%>
					<strong><%= formatNumber(oGmarket.FItemList(i).FGmarketPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oGmarket.FItemList(i).FGmarketPrice,0)&"<br>"
				End If

				If Not IsNULL(oGmarket.FItemList(i).FSpecialPrice) Then
					If (now() >= oGmarket.FItemList(i).FStartDate) And (now() <= oGmarket.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oGmarket.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oGmarket.FItemList(i).FSellyn="Y" and oGmarket.FItemList(i).FGmarketSellYn<>"Y") or (oGmarket.FItemList(i).FSellyn<>"Y" and oGmarket.FItemList(i).FGmarketSellYn="Y") Then
	%>
					<strong><%= oGmarket.FItemList(i).FGmarketSellYn %></strong>
	<%
				Else
					response.write oGmarket.FItemList(i).FGmarketSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oGmarket.FItemList(i).FGmarketGoodNo)) Then
			Response.Write "<a target='_blank' href='https://item.gmarket.co.kr/Item?goodscode="&oGmarket.FItemList(i).FGmarketGoodNo&"'>"&oGmarket.FItemList(i).FGmarketGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oGmarket.FItemList(i).FG9GoodNo)) Then
			Response.Write "<a target='_blank' href='http://www.g9.co.kr/Display/VIP/Index/"&oGmarket.FItemList(i).FG9GoodNo&"'>"&oGmarket.FItemList(i).FG9GoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oGmarket.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oGmarket.FItemList(i).FItemID%>','0');"><%= oGmarket.FItemList(i).FoptionCnt %>:<%= oGmarket.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oGmarket.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oGmarket.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��(ī)"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�(ī)</font>"
		End If

		' If oGmarket.FItemList(i).FBrandCode > 0 Then
		' 	response.write "<br />��Ī��(��)"
		' Else
		' 	response.write "<br /><font color='darkred'>��Ī�ȵ�(��)</font>"
		' End If
	%>
	</td>
	<td align="center">
		<%= oGmarket.FItemList(i).FinfoDiv %>
		<%
		If (oGmarket.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oGmarket.FItemList(i).FlastErrStr) &"'>ERR:"& oGmarket.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oGmarket.FItemList(i).FAPIadditem="Y","<font color='BLUE'>"&oGmarket.FItemList(i).FAPIadditem&"</font>", "<font color='RED'>"&oGmarket.FItemList(i).FAPIadditem&"</font>") %>&nbsp;|
		<%= Chkiif(oGmarket.FItemList(i).FAPIaddgosi="Y","<font color='BLUE'>"&oGmarket.FItemList(i).FAPIaddgosi&"</font>", "<font color='RED'>"&oGmarket.FItemList(i).FAPIaddgosi&"</font>") %>&nbsp;|
		<%= Chkiif(oGmarket.FItemList(i).FAPIaddopt="Y","<font color='BLUE'>"&oGmarket.FItemList(i).FAPIaddopt&"</font>", "<font color='RED'>"&oGmarket.FItemList(i).FAPIaddopt&"</font>") %>
		<br>
		<%= oGmarket.FItemList(i).FDisplayDate %>
	</td>
</tr>
<% GmarketGoodNoArray = GmarketGoodNoArray & oGmarket.FItemList(i).FGmarketGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= GmarketGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oGmarket.HasPreScroll then %>
		<a href="javascript:goPage('<%= oGmarket.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oGmarket.StartScrollPage to oGmarket.FScrollCount + oGmarket.StartScrollPage - 1 %>
    		<% if i>oGmarket.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oGmarket.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
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
	<input type="hidden" name="gmarketGoodNo" value= <%= gmarketGoodNo %>>
	<input type="hidden" name="g9GoodNo" value= <%= g9GoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchBrand" value= <%= MatchBrand %>>
	<input type="hidden" name="MatchG9" value= <%= MatchG9 %>>
	<input type="hidden" name="sellpriceChk" value= <%= sellpriceChk %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="addOptErr" value= <%= addOptErr %>>
	<input type="hidden" name="gmarketYes10x10No" value= <%= gmarketYes10x10No %>>
	<input type="hidden" name="gmarketNo10x10Yes" value= <%= gmarketNo10x10Yes %>>
	<input type="hidden" name="gmarketKeepSell" value= <%= gmarketKeepSell %>>
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
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oGmarket = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
