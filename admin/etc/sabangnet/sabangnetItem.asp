<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/sabangnet/sabangnetcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, sabangnetGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, sabangnetYes10x10No, sabangnetNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, isextusing, cisextusing, rctsellcnt
Dim page, i, research
Dim oSabangnet, scheduleNotInItemid
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
sabangnetGoodNo			= request("sabangnetGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
sabangnetYes10x10No		= request("sabangnetYes10x10No")
sabangnetNo10x10Yes		= request("sabangnetNo10x10Yes")
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
purchasetype			= request("purchasetype")

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

If (session("ssBctID")="kjy8517") Then
	'itemid = "1562224"	'��ǰ
	'itemid = "1770667"
	'itemid = "1795718"	'�ɼ�
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
If sabangnetGoodNo <> "" then
	Dim iA2, arrTemp2, arrsabangnetGoodNo
	sabangnetGoodNo = replace(sabangnetGoodNo,",",chr(10))
	sabangnetGoodNo = replace(sabangnetGoodNo,chr(13),"")
	arrTemp2 = Split(sabangnetGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrsabangnetGoodNo = arrsabangnetGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	sabangnetGoodNo = left(arrsabangnetGoodNo,len(arrsabangnetGoodNo)-1)
End If

Set oSabangnet = new CSabangnet
	oSabangnet.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oSabangnet.FPageSize					= 100
Else
	oSabangnet.FPageSize					= 50
End If
	oSabangnet.FRectCDL						= request("cdl")
	oSabangnet.FRectCDM						= request("cdm")
	oSabangnet.FRectCDS						= request("cds")
	oSabangnet.FRectItemID					= itemid
	oSabangnet.FRectItemName				= itemname
	oSabangnet.FRectSellYn					= sellyn
	oSabangnet.FRectLimitYn					= limityn
	oSabangnet.FRectSailYn					= sailyn
'	oSabangnet.FRectonlyValidMargin			= onlyValidMargin
	oSabangnet.FRectStartMargin				= startMargin
	oSabangnet.FRectEndMargin				= endMargin
	oSabangnet.FRectMakerid					= makerid
	oSabangnet.FRectsabangnetGoodNo			= sabangnetGoodNo
	oSabangnet.FRectMatchCate				= MatchCate
	oSabangnet.FRectIsMadeHand				= isMadeHand
	oSabangnet.FRectIsOption				= isOption
	oSabangnet.FRectIsReged					= isReged
	oSabangnet.FRectNotinmakerid			= notinmakerid
	oSabangnet.FRectNotinitemid				= notinitemid
	oSabangnet.FRectExcTrans				= exctrans
	oSabangnet.FRectPriceOption				= priceOption
	oSabangnet.FRectDeliverytype			= deliverytype
	oSabangnet.FRectMwdiv					= mwdiv
	oSabangnet.FRectIsextusing				= isextusing
	oSabangnet.FRectCisextusing				= cisextusing
	oSabangnet.FRectRctsellcnt				= rctsellcnt
	oSabangnet.FRectIsSpecialPrice     		= isSpecialPrice
	oSabangnet.FRectScheduleNotInItemid		= scheduleNotInItemid

	oSabangnet.FRectExtNotReg				= ExtNotReg
	oSabangnet.FRectExpensive10x10			= expensive10x10
	oSabangnet.FRectdiffPrc					= diffPrc
	oSabangnet.FRectsabangnetYes10x10No		= sabangnetYes10x10No
	oSabangnet.FRectsabangnetNo10x10Yes		= sabangnetNo10x10Yes
	oSabangnet.FRectExtSellYn				= extsellyn
	oSabangnet.FRectInfoDiv					= infoDiv
	oSabangnet.FRectFailCntOverExcept		= ""
	oSabangnet.FRectFailCntExists			= failCntExists
	oSabangnet.FRectReqEdit					= reqEdit
	oSabangnet.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oSabangnet.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oSabangnet.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oSabangnet.getsabangnetreqExpireItemList
Else
	oSabangnet.getSabangnetRegedItemList		'�� �� ����Ʈ
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=sabangnet","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=sabangnet','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=sabangnet','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="sabangnetYes10x10No")&&(frm.sabangnetYes10x10No.checked)){ frm.sabangnetYes10x10No.checked=false }
	if ((comp.name!="sabangnetNo10x10Yes")&&(frm.sabangnetNo10x10Yes.checked)){ frm.sabangnetNo10x10Yes.checked=false }
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

    if ((comp.name=="sabangnetYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="sabangnetNo10x10Yes")&&(comp.checked)){
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
			comp.form.notinmakerid.value = "N";
			comp.form.notinitemid.value = "N";
    	}
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.sabangnetYes10x10No.checked){
            comp.form.sabangnetYes10x10No.checked = false;
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
	if ((comp.name!="sabangnetYes10x10No")&&(frm.sabangnetYes10x10No.checked)){ frm.sabangnetYes10x10No.checked=false }
	if ((comp.name!="sabangnetNo10x10Yes")&&(frm.sabangnetNo10x10Yes.checked)){ frm.sabangnetNo10x10Yes.checked=false }
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
// ���� ī�װ��� ���ݿ� ���..���� 1ȸ�� �س����� �� �� ����
function sabangnetCategoryRegProcess(){
	if (confirm('sabangnet�� 10x10�� ����ī�װ��� ����Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "Category";
		document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
		document.frmSvArr.submit();
	}
}
// ������� �׸� ��ȸ
function sabangnetGosiInfoSelectProcess(){
	if (confirm('sabangnet�� ��ǰ ��� ������ ȣ���ϰڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "GosiInfo";
		document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
		document.frmSvArr.submit();
	}
}
// ���õ� ��ǰ ���
function sabangnetREGProcess() {
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

    if (confirm('sabangnet�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function sabangnetSellYnProcess(chkYn) {
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
		case "N": strSell="�Ͻ�����";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
        document.frmSvArr.submit();
    }
}

//�⺻���� ����
function sabangnetEditPriceProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ���� �� �ǸŻ��¸� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ ���θ��� DATA ����
function sabangnetShopDataProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ���θ��� DATA�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "SDATA";
        document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
        document.frmSvArr.submit();
    }
}
//�⺻���� + �ɼ����� ����
function sabangnetEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ü ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/sabangnet/actsabangnetReq.asp"
        document.frmSvArr.submit();
    }
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=sabangnet&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
		<a href="http://s400.sabangnet.co.kr" target="_blank">����Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenbyten | store101010*! ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		���� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="sabangnetGoodNo" id="itemid"><%= replace(replace(sabangnetGoodNo,",",chr(10)), "'", "")%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="sabangnetYes10x10No" <%= ChkIIF(sabangnetYes10x10No="on","checked","") %> ><font color=red>�����Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="sabangnetNo10x10Yes" <%= ChkIIF(sabangnetNo10x10Yes="on","checked","") %> ><font color=red>����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>

<p />

* �������ܻ�ǰ : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, �ǸŰ�(���ΰ�) 1���� �̸�

<p />

<form name="frmReg" method="post" action="sabangnetitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('sabangnet');">&nbsp;&nbsp;
	<% If (session("ssBctID")="kjy8517") Then %>
		<!--
		<label><input class="button" type="button" value="CateGory" onclick="sabangnetCategoryRegProcess()"></label>&nbsp;&nbsp;
		<label><input class="button" type="button" value="GosiInfo" onclick="sabangnetGosiInfoSelectProcess()"></label>
		 -->
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
				<input class="button" type="button" id="btnREG" value="���" onClick="sabangnetREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEdit" value="����" onClick="sabangnetEditProcess();" style=color:blue;font-weight:bold>&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditPrice" value="���� �� ����" onClick="sabangnetEditPriceProcess();">
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnEditPrice" value="���θ�DATA����" onClick="sabangnetShopDataProcess();">
			<% End If %>

			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ͻ�����</option>
					<option value="Y">�Ǹ�</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="sabangnetSellYnProcess(frmReg.chgSellYn.value);">
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(oSabangnet.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSabangnet.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">���� �����<br>���� ����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">����<br>���ݹ��Ǹ�</td>
	<td width="70">����<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oSabangnet.FResultCount - 1 %>
<tr align="center" <% If oSabangnet.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oSabangnet.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oSabangnet.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oSabangnet.FItemList(i).FItemID %>','sabangnet','<%=oSabangnet.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oSabangnet.FItemList(i).FItemID%>" target="_blank"><%= oSabangnet.FItemList(i).FItemID %></a>
		<% If oSabangnet.FItemList(i).FsabangnetStatcd <> 7 Then %>
		<br><%= oSabangnet.FItemList(i).getsabangnetStatName %>
		<% End If %>
		<%= oSabangnet.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oSabangnet.FItemList(i).FMakerid %> <%= oSabangnet.FItemList(i).getDeliverytypeName %><br><%= oSabangnet.FItemList(i).FItemName %></td>
	<td align="center"><%= oSabangnet.FItemList(i).FRegdate %><br><%= oSabangnet.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oSabangnet.FItemList(i).FsabangnetRegdate %><br><%= oSabangnet.FItemList(i).FsabangnetLastUpdate %></td>
	<td align="right">
		<% If oSabangnet.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oSabangnet.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oSabangnet.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oSabangnet.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oSabangnet.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-oSabangnet.FItemList(i).Fbuycash/oSabangnet.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oSabangnet.FItemList(i).IsSoldOut Then
			If oSabangnet.FItemList(i).FSellyn = "N" Then
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
		If oSabangnet.FItemList(i).FItemdiv = "06" OR oSabangnet.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oSabangnet.FItemList(i).FsabangnetStatCd > 0) Then
			If Not IsNULL(oSabangnet.FItemList(i).FsabangnetPrice) Then
				If (oSabangnet.FItemList(i).Fsellcash <> oSabangnet.FItemList(i).FsabangnetPrice) Then
	%>
					<strong><%= formatNumber(oSabangnet.FItemList(i).FsabangnetPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oSabangnet.FItemList(i).FsabangnetPrice,0)&"<br>"
				End If

				If Not IsNULL(oSabangnet.FItemList(i).FSpecialPrice) Then
					If (now() >= oSabangnet.FItemList(i).FStartDate) And (now() <= oSabangnet.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oSabangnet.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oSabangnet.FItemList(i).FSellyn="Y" and oSabangnet.FItemList(i).FsabangnetSellYn<>"Y") or (oSabangnet.FItemList(i).FSellyn<>"Y" and oSabangnet.FItemList(i).FsabangnetSellYn="Y") Then
	%>
					<strong><%= oSabangnet.FItemList(i).FsabangnetSellYn %></strong>
	<%
				Else
					response.write oSabangnet.FItemList(i).FsabangnetSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oSabangnet.FItemList(i).FsabangnetGoodNo)) Then
			Response.Write "<a target='_blank' href='http://s400.sabangnet.co.kr/mall_join/product_preview.html?product_id="&oSabangnet.FItemList(i).FsabangnetGoodNo&"&compayny_id=mw58297'>"&oSabangnet.FItemList(i).FsabangnetGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oSabangnet.FItemList(i).Freguserid %></td>
	<td align="center">
		<a href="javascript:popManageOptAddPrc('<%=oSabangnet.FItemList(i).FItemID%>','0');"><%= oSabangnet.FItemList(i).FoptionCnt %>:<%= oSabangnet.FItemList(i).FregedOptCnt %></a>
	</td>
	<td align="center"><%= oSabangnet.FItemList(i).FrctSellCNT %></td>
	<td align="center">
		<%= oSabangnet.FItemList(i).FinfoDiv %>
		<%
		If (oSabangnet.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oSabangnet.FItemList(i).FlastErrStr) &"'>ERR:"& oSabangnet.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oSabangnet.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSabangnet.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSabangnet.StartScrollPage to oSabangnet.FScrollCount + oSabangnet.StartScrollPage - 1 %>
    		<% if i>oSabangnet.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSabangnet.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oSabangnet = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
