<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/auction/auctioncls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, auctionGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, auctionYes10x10No, auctionNo10x10Yes, auctionKeepSell, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice
Dim page, i, research
Dim oAuction, AuctionGoodNoArray, isextusing, cisextusing, rctsellcnt, scheduleNotInItemid
dim startsell, stopsell, xl
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
auctionGoodNo			= request("auctionGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
auctionYes10x10No		= request("auctionYes10x10No")
auctionNo10x10Yes		= request("auctionNo10x10Yes")
auctionKeepSell			= request("auctionKeepSell")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
startsell				= requestCheckVar(request("startsell"), 1)
stopsell				= requestCheckVar(request("stopsell"), 1)
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
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"

	if (stopsell = "Y") then
		'// �Ǹ����� ��� ��ǰ���
		ExtNotReg = "D"
		sellyn = "N"
		extsellyn = "Y"
		auctionYes10x10No = "on"
		onlyValidMargin = ""
	elseif (startsell = "Y") then
		'// �Ǹ���ȯ ��� ��ǰ���
		ExtNotReg = "D"
		sellyn = "Y"
		extsellyn = "N"
		onlyValidMargin = "Y"
		notinmakerid = "N"
		notinitemid = "N"
		auctionNo10x10Yes = "on"
	end if
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
'���� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If auctionGoodNo <> "" then
	Dim iA2, arrTemp2, arrauctionGoodNo
	auctionGoodNo = replace(auctionGoodNo,",",chr(10))
	auctionGoodNo = replace(auctionGoodNo,chr(13),"")
	arrTemp2 = Split(auctionGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrauctionGoodNo = arrauctionGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	auctionGoodNo = left(arrauctionGoodNo,len(arrauctionGoodNo)-1)
End If

Set oAuction = new CAuction
	oAuction.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oAuction.FPageSize					= 100

Else
	oAuction.FPageSize					= 50
End If
	oAuction.FRectCDL					= request("cdl")
	oAuction.FRectCDM					= request("cdm")
	oAuction.FRectCDS					= request("cds")
	oAuction.FRectItemID				= itemid
	oAuction.FRectItemName				= itemname
	oAuction.FRectSellYn				= sellyn
	oAuction.FRectLimitYn				= limityn
	oAuction.FRectSailYn				= sailyn
'	oAuction.FRectonlyValidMargin		= onlyValidMargin
	oAuction.FRectStartMargin			= startMargin
	oAuction.FRectEndMargin				= endMargin
	oAuction.FRectMakerid				= makerid
	oAuction.FRectAuctionGoodNo			= auctionGoodNo
	oAuction.FRectMatchCate				= MatchCate
	oAuction.FRectIsMadeHand			= isMadeHand
	oAuction.FRectIsOption				= isOption
	oAuction.FRectIsReged				= isReged
	oAuction.FRectNotinmakerid			= notinmakerid
	oAuction.FRectNotinitemid			= notinitemid
	oAuction.FRectExcTrans				= exctrans
	oAuction.FRectPriceOption			= priceOption
	oAuction.FRectIsSpecialPrice       	= isSpecialPrice
	oAuction.FRectDeliverytype			= deliverytype
	oAuction.FRectMwdiv					= mwdiv
	oAuction.FRectScheduleNotInItemid	= scheduleNotInItemid
	oAuction.FRectIsextusing			= isextusing
	oAuction.FRectCisextusing			= cisextusing
	oAuction.FRectRctsellcnt			= rctsellcnt

	oAuction.FRectExtNotReg				= ExtNotReg
	oAuction.FRectExpensive10x10		= expensive10x10
	oAuction.FRectdiffPrc				= diffPrc
	oAuction.FRectAuctionYes10x10No		= auctionYes10x10No
	oAuction.FRectAuctionNo10x10Yes		= auctionNo10x10Yes
	oAuction.FRectAuctionKeepSell		= auctionKeepSell
	oAuction.FRectExtSellYn				= extsellyn
	oAuction.FRectInfoDiv				= infoDiv
	oAuction.FRectFailCntOverExcept		= ""
	oAuction.FRectFailCntExists			= failCntExists
	oAuction.FRectReqEdit				= reqEdit
	oAuction.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oAuction.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oAuction.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oAuction.getAuctionreqExpireItemList
Else
	oAuction.getAuctionRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=auction1010List"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=auction1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=auction1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=auction1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=auction1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="auctionKeepSell")&&(frm.auctionKeepSell.checked)){ frm.auctionKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="auctionYes10x10No")&&(frm.auctionYes10x10No.checked)){ frm.auctionYes10x10No.checked=false }
	if ((comp.name!="auctionNo10x10Yes")&&(frm.auctionNo10x10Yes.checked)){ frm.auctionNo10x10Yes.checked=false }
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

    if ((comp.name=="auctionYes10x10No")&&(comp.checked)){
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
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="auctionNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.auctionYes10x10No.checked){
            comp.form.auctionYes10x10No.checked = false;
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

    if ((comp.name=="auctionKeepSell")&&(comp.checked)){
        if (comp.form.auctionYes10x10No.checked){
            comp.form.auctionYes10x10No.checked = false;
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
	if ((comp.name!="auctionKeepSell")&&(frm.auctionKeepSell.checked)){ frm.auctionKeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="auctionYes10x10No")&&(frm.auctionYes10x10No.checked)){ frm.auctionYes10x10No.checked=false }
	if ((comp.name!="auctionNo10x10Yes")&&(frm.auctionNo10x10Yes.checked)){ frm.auctionNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/auction/popauctioncateList.asp","popCateAuctionmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// ���õ� ��ǰ �⺻���� ���
function AuctionSelectRegProcess() {
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �⺻������ ��� �Ͻðڽ��ϱ�?\n\n�����ڵ带 ���Ϲޱ� ���� �⺻���� ����Դϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddItem";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ɼ����� ���
function AuctionSelectOPTProcess() {
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ɼ������� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGAddOPT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ������� ���
function AuctionSelectInfoCdRegProcess() {
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ������ø� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGInfoCd";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �⺻���� + �ɼ� + ������� ���
function AuctionREGProcess() {
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ������� ���
function AuctionOnSaleEditProcess() {
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ���¸� �Ǹ������� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REGOnSale";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function AuctionSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� ����
function AuctionEditInfoProcess(){
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� ��ǰ������ ���� �Ͻðڽ��ϱ�?\n\n�ػ�ǰ��, ����, �̹���, ��ǰ�󼼵��� �����˴ϴ�')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditInfo";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//�ɼ� ��ȸ
function AuctionViewProcess(){
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� �ɼ��� ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//�ɼ� ����
function AuctionEditOPTProcess(){
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� �ɼ��� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� + �ɼ����� ����
function AuctionEditProcess(){
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

    if (confirm('Auction�� �����Ͻ� ' + chkSel + '�� �⺻���� + �ɼ� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ ����
function AuctionDeleteProcess(){
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
    if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n���� ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
			document.frmSvArr.submit();
		}
    }
}

//�����ڵ� �˻�
function AuctionCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "<%=apiURL%>/outmall/auction/actauctionReq.asp?cmdparam=auctionCommonCode&CommCD="+ccd+"";
	}
}


//�����ڵ� �˻�
function fnAuctionCommCDSubmit() {
	var ccd;
	var goodsGrpCd;
	ccd = document.getElementById('CommCD2').value;
	//goodsGrpCd = $("#goodsGrpCd option:selected").val();

	goodsGrpCd = $("#goodsGrpCd").val();
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "/admin/etc/auction/actauctionReq.asp?cmdparam=ebayCommonCode&CommCD="+ccd+"&goodsGrpCd="+goodsGrpCd;
	}
}
function jsByValue(s){
	if((s == "brand") || (s == "maker") || (s =="placepolicy" || s == "infocodedtl" || s == "mastercode" || s == "sitecode")) {
		$("#goodsGrpCd_span").show();
	}else{
		$("#goodsGrpCd_span").hide();
	}
}

//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=auction1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
//����, ��ȿ�� ��� �˾�
function popAuctionDate(iitemid){
    var pdate = window.open("/admin/etc/auction/popAuctionDate.asp?itemid="+iitemid+'&mallid=auction1010',"popAuctionDate","width=500,height=200,scrollbars=yes,resizable=yes");
	pdate.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		<a href="http://www.esmplus.com/Home/Home" target="_blank">����Admin�ٷΰ���</a>
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
		���� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="auctionGoodNo" id="itemid"><%= replace(replace(auctionGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >���� ��ϼ���_OnSale��
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >���� ���۽õ� �� ����
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >���� ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >���� ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>���� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionYes10x10No" <%= ChkIIF(auctionYes10x10No="on","checked","") %> ><font color=red>�����Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionNo10x10Yes" <%= ChkIIF(auctionNo10x10Yes="on","checked","") %> ><font color=red>����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="auctionKeepSell" <%= ChkIIF(auctionKeepSell="on","checked","") %> ><font color=red>�Ǹ�����</font> �ؾ��� ��ǰ����</label>
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
* �������ܻ�ǰ2 : ��ǰ���� IFRAME TAG ����� ��ǰ

<p />

<form name="frmReg" method="post" action="auctionitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('auction1010');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="�⺻����" onClick="AuctionSelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOPTSel" value="�ɼ�����" onClick="AuctionSelectOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnInfocdSel" value="��ǰ���" onClick="AuctionSelectInfoCdRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnREG" value="�⺻+�ɼ�+���" onClick="AuctionREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnOnSale" value="OnSale����" onClick="AuctionOnSaleEditProcess();" style=color:red;font-weight:bold>
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditInfo" value="�⺻����(����)" onClick="AuctionEditInfoProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEditOPT" value="�ɼ�����" onClick="AuctionEditOPTProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOEdit" value="�⺻+�ɼ�" onClick="AuctionEditProcess();" style=color:blue;font-weight:bold>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="nys1006") OR (session("ssBctID")="z0516") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="��ǰ����" onClick="AuctionDeleteProcess();" style=font-weight:bold>
			<% End If %>
				<br><br>
				������ǰ ��ȸ :
				<input class="button" type="button" id="btnViewItem" value="�ɼ���ȸ" onClick="AuctionViewProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="GetShippingPlaceCode">�������ڵ�
					<option value="GetNationCode">�������ڵ�
					<option value="GetDeliveryList">��ۻ�(�ù�)��ȸ
					<option value="GetPaidOrderList">�ֹ�Test
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="AuctionCommCDSubmit();" >

				2.0 : 
				<select name="CommCD2" class="select" id="CommCD2" onChange="jsByValue(this.value);">
					<option value="">- Choice -
					<option value="mastercode">�������ڵ���ȸ(�����ڵ��)</option>
					<option value="sitecode">�����ڵ���ȸ(�������ڵ��)</option>
				</select>
				<span id="goodsGrpCd_span" style="display:none;">
					<input type="text" name="goodsGrpCd" id="goodsGrpCd">
				</span>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="fnAuctionCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ�</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="AuctionSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="17">
		�˻���� : <b><%= FormatNumber(oAuction.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oAuction.FTotalPage,0) %></b>
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
	<td width="140">Auction�����<br>Auction����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Auction<br>���ݹ��Ǹ�</td>
	<td width="70">Auction<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="80">ǰ��</td>
	<td width="100">��|��|��<br>OnSale������</td>
</tr>
<% For i=0 to oAuction.FResultCount - 1 %>
<tr align="center" <% If oAuction.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oAuction.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oAuction.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oAuction.FItemList(i).FItemID %>','auction1010','<%=oAuction.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oAuction.FItemList(i).FItemID%>" target="_blank"><%= oAuction.FItemList(i).FItemID %></a>
		<%
			If (xl <> "Y") Then
				If oAuction.FItemList(i).FAuctionStatcd <> 7 Then
		%>
		<br><%= oAuction.FItemList(i).getAuctionStatName %>
		<%
				End If
				response.write oAuction.FItemList(i).getLimitHtmlStr
			End If
		%>
	</td>
	<td align="left"><%= oAuction.FItemList(i).FMakerid %> <%= oAuction.FItemList(i).getDeliverytypeName %><br><%= oAuction.FItemList(i).FItemName %></td>
	<td align="center"><%= oAuction.FItemList(i).FRegdate %><br><%= oAuction.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oAuction.FItemList(i).FAuctionRegdate %><br><%= oAuction.FItemList(i).FAuctionLastUpdate %></td>
	<td align="right">
		<% If oAuction.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oAuction.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oAuction.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oAuction.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oAuction.FItemList(i).Fsellcash = 0 Then
		elseif (oAuction.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oAuction.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oAuction.FItemList(i).FOrgSuplycash/oAuction.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oAuction.FItemList(i).Fbuycash/oAuction.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oAuction.FItemList(i).Fbuycash/oAuction.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oAuction.FItemList(i).IsSoldOut Then
			If oAuction.FItemList(i).FSellyn = "N" Then
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
		If oAuction.FItemList(i).FItemdiv = "06" OR oAuction.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oAuction.FItemList(i).FAuctionStatCd > 0) Then
			If Not IsNULL(oAuction.FItemList(i).FAuctionPrice) Then
				If (oAuction.FItemList(i).Mustprice <> oAuction.FItemList(i).FAuctionPrice) Then
	%>
					<strong><%= formatNumber(oAuction.FItemList(i).FAuctionPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oAuction.FItemList(i).FAuctionPrice,0)&"<br>"
				End If

				If Not IsNULL(oAuction.FItemList(i).FSpecialPrice) Then
					If (now() >= oAuction.FItemList(i).FStartDate) And (now() <= oAuction.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oAuction.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oAuction.FItemList(i).FSellyn="Y" and oAuction.FItemList(i).FAuctionSellYn<>"Y") or (oAuction.FItemList(i).FSellyn<>"Y" and oAuction.FItemList(i).FAuctionSellYn="Y") Then
	%>
					<strong><%= oAuction.FItemList(i).FAuctionSellYn %></strong>
	<%
				Else
					response.write oAuction.FItemList(i).FAuctionSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oAuction.FItemList(i).FAuctionGoodNo)) Then
			Response.Write "<a target='_blank' href='http://itempage3.auction.co.kr/detailview.aspx?itemNo="&oAuction.FItemList(i).FAuctionGoodNo&"'>"&oAuction.FItemList(i).FAuctionGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oAuction.FItemList(i).Freguserid %></td>
	<td align="center">
		<a href="javascript:popManageOptAddPrc('<%=oAuction.FItemList(i).FItemID%>','0');"><%= oAuction.FItemList(i).FoptionCnt %>:<%= oAuction.FItemList(i).FregedOptCnt %></a>
		<br>
		<input type="button" class="button" value="����" onclick="javascript:popAuctionDate('<%=oAuction.FItemList(i).FItemID%>');">
	</td>
	<td align="center"><%= oAuction.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oAuction.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oAuction.FItemList(i).FinfoDiv %>
		<%
		If (oAuction.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oAuction.FItemList(i).FlastErrStr) &"'>ERR:"& oAuction.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oAuction.FItemList(i).FAPIadditem="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIadditem&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIadditem&"</font>") %>&nbsp;|
		<%= Chkiif(oAuction.FItemList(i).FAPIaddopt="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIaddopt&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIaddopt&"</font>") %>&nbsp;|
		<%= Chkiif(oAuction.FItemList(i).FAPIaddgosi="Y","<font color='BLUE'>"&oAuction.FItemList(i).FAPIaddgosi&"</font>", "<font color='RED'>"&oAuction.FItemList(i).FAPIaddgosi&"</font>") %>
		<br>
		<%= oAuction.FItemList(i).FOnSaleRegdate %>
	</td>
</tr>
<% AuctionGoodNoArray = AuctionGoodNoArray & oAuction.FItemList(i).FAuctionGoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= AuctionGoodNoArray %></textarea>
<% End If %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oAuction.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAuction.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oAuction.StartScrollPage to oAuction.FScrollCount + oAuction.StartScrollPage - 1 %>
    		<% if i>oAuction.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oAuction.HasNextScroll then %>
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
	<input type="hidden" name="auctionGoodNo" value= <%= auctionGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="auctionYes10x10No" value= <%= auctionYes10x10No %>>
	<input type="hidden" name="auctionNo10x10Yes" value= <%= auctionNo10x10Yes %>>
	<input type="hidden" name="auctionKeepSell" value= <%= auctionKeepSell %>>
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
	<input type="hidden" name="startsell" value= <%= startsell %>>
	<input type="hidden" name="stopsell" value= <%= stopsell %>>
	<input type="hidden" name="rctsellcnt" value= <%= rctsellcnt %>>
	<input type="hidden" name="purchasetype" value= <%= purchasetype %>>
</form>
<% SET oAuction = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
