<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kakaogift/kakaogiftcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, kakaogiftGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, morningJY, deliverytype, mwdiv
Dim expensive10x10, diffPrc, kakaogiftYes10x10No, kakaogiftNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, isextusing, cisextusing, rctsellcnt
Dim okakaogift, itemidarr, notinitemid, exctrans, isSpecialPrice
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
kakaogiftGoodNo			= request("kakaogiftGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
kakaogiftYes10x10No		= request("kakaogiftYes10x10No")
kakaogiftNo10x10Yes		= request("kakaogiftNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)

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
'	itemid = "1291678"
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

'KakaoGift ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If kakaogiftGoodNo <> "" then
	Dim iA2, arrTemp2, arrkakaogiftGoodNo
	kakaogiftGoodNo = replace(kakaogiftGoodNo,",",chr(10))
	kakaogiftGoodNo = replace(kakaogiftGoodNo,chr(13),"")
	arrTemp2 = Split(kakaogiftGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrkakaogiftGoodNo = arrkakaogiftGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	kakaogiftGoodNo = left(arrkakaogiftGoodNo,len(arrkakaogiftGoodNo)-1)
End If

Set okakaogift = new Ckakaogift
	okakaogift.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	okakaogift.FPageSize					= 50
Else
	okakaogift.FPageSize					= 20
End If
	okakaogift.FRectCDL					= request("cdl")
	okakaogift.FRectCDM					= request("cdm")
	okakaogift.FRectCDS					= request("cds")
	okakaogift.FRectItemID				= itemid
	okakaogift.FRectItemName				= itemname
	okakaogift.FRectSellYn				= sellyn
	okakaogift.FRectLimitYn				= limityn
	okakaogift.FRectSailYn				= sailyn
'	okakaogift.FRectonlyValidMargin		= onlyValidMargin
	okakaogift.FRectStartMargin			= startMargin
	okakaogift.FRectEndMargin			= endMargin
	okakaogift.FRectMakerid				= makerid
	okakaogift.FRectkakaogiftGoodNo		= kakaogiftGoodNo
	okakaogift.FRectMatchCate			= MatchCate
	okakaogift.FRectIsMadeHand			= isMadeHand
	okakaogift.FRectIsOption				= isOption
	okakaogift.FRectIsReged				= isReged
	okakaogift.FRectNotinmakerid			= notinmakerid
	okakaogift.FRectPriceOption			= priceOption
	okakaogift.FRectDeliverytype			= deliverytype
	okakaogift.FRectMwdiv				= mwdiv
	okakaogift.FRectIsextusing			= isextusing
	okakaogift.FRectCisextusing			= cisextusing
	okakaogift.FRectRctsellcnt			= rctsellcnt

	okakaogift.FRectExtNotReg			= ExtNotReg
	okakaogift.FRectExpensive10x10		= expensive10x10
	okakaogift.FRectdiffPrc				= diffPrc
	okakaogift.FRectkakaogiftYes10x10No	= kakaogiftYes10x10No
	okakaogift.FRectkakaogiftNo10x10Yes	= kakaogiftNo10x10Yes
	okakaogift.FRectExtSellYn			= extsellyn
	okakaogift.FRectInfoDiv				= infoDiv
	okakaogift.FRectFailCntOverExcept	= ""
	okakaogift.FRectFailCntExists		= failCntExists
	okakaogift.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    okakaogift.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    okakaogift.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	okakaogift.getkakaogiftreqExpireItemList
Else
	okakaogift.getkakaogiftRegedItemList		'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=kakaogift","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=kakaogift','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="kakaogiftYes10x10No")&&(frm.kakaogiftYes10x10No.checked)){ frm.kakaogiftYes10x10No.checked=false }
	if ((comp.name!="kakaogiftNo10x10Yes")&&(frm.kakaogiftNo10x10Yes.checked)){ frm.kakaogiftNo10x10Yes.checked=false }
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

    if ((comp.name=="kakaogiftYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="kakaogiftNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.kakaogiftYes10x10No.checked){
            comp.form.kakaogiftYes10x10No.checked = false;
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
	if ((comp.name!="kakaogiftYes10x10No")&&(frm.kakaogiftYes10x10No.checked)){ frm.kakaogiftYes10x10No.checked=false }
	if ((comp.name!="kakaogiftNo10x10Yes")&&(frm.kakaogiftNo10x10Yes.checked)){ frm.kakaogiftNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/kakaogift/popkakaogiftcateList.asp","popCatekakaogiftmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//Que �α� ����Ʈ �˾�
function pop_songjang(mallid) {
	var pCM6 = window.open("/admin/etc/kakaogift/popsongjangList.asp","pop_songjang","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}
//�Ǹ��ߴܿ��List
function pop_maySoldout() {
	var pCM5 = window.open("/admin/etc/kakaogift/popMaySoldoutList.asp","pop_maySoldout","width=1400,height=300,scrollbars=yes,resizable=yes");
	pCM5.focus();
}

// ���õ� ��ǰ �ϰ� ���
function kakaogiftSelectChkProcess() {
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

    if (confirm('KakaoGift�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ ��� �Ͻðڽ��ϱ�?')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHK";
		document.frmSvArr.action = "<%=apiURL%>/outmall/kakaogift/actkakaogiftReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function kakaogiftSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/kakaogift/actkakaogiftReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function kakaogiftriceEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��KakaoGift���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/kakaogift/actkakaogiftReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function kakaogiftEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� ��� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/kakaogift/actkakaogiftReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function kakaogiftExcelUpload() {
	var winFile = window.open("/admin/etc/kakaogift/popRegFile.asp","popFile","width=600, height=300 ,scrollbars=yes,resizable=yes");
	winFile.focus();
}

// ���õ� ��ǰ �ϰ� ����
/*
function kakaogiftSelectDeliveryProcess() {

    if (confirm('�ù�� �ڵ带 ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "Delivery";
		document.frmSvArr.action = "<%=apiURL%>/outmall/kakaogift/actkakaogiftReq.asp"
		document.frmSvArr.submit();
    }
}
*/

//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=kakaogift&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
		<a href="https://sell.kakao.com/dashboard/index" target="_blank">KakaoGiftAdmin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then

			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		KakaoGift ��ǰ�ڵ� : <textarea rows="2" cols="20" name="kakaogiftGoodNo" id="itemid"><%=replace(kakaogiftGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >KakaoGift ��Ͻõ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >KakaoGift ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >KakaoGift ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>KakaoGift ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="kakaogiftYes10x10No" <%= ChkIIF(kakaogiftYes10x10No="on","checked","") %> ><font color=red>KakaoGift�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="kakaogiftNo10x10Yes" <%= ChkIIF(kakaogiftNo10x10Yes="on","checked","") %> ><font color=red>KakaoGiftǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
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
			<td align="left">
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="����List" onclick="pop_songjang();">&nbsp;&nbsp;
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('kakaogift');">&nbsp;&nbsp;
				<input class="button" type="button" value="�Ǹ��ߴܿ��List" onclick="pop_maySoldout();">&nbsp;&nbsp;
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
				���� üũ :
				<input class="button" type="button" id="btnRegSel" value="üũ" onClick="kakaogiftSelectChkProcess();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnStock" value="����" onClick="kakaogiftEditProcess();">&nbsp;&nbsp;
				<br><br>
				��ǰ����Ʈ ���ε�(RPA��) :
				<input class="button" type="button" id="btnStock" value="EXCEL���ε�" onClick="kakaogiftExcelUpload();">&nbsp;&nbsp;
				<% if (FALSE) then %>>
				<br><br>
				�����ڵ� ��ȸ :
				<input class="button" type="button" id="btnStock" value="�ù��" onClick="kakaogiftSelectDeliveryProcess();">&nbsp;&nbsp;
			    <% end if%>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="kakaogiftSellYnProcess(frmReg.chgSellYn.value);">
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
	<td colspan="20">
		�˻���� : <b><%= FormatNumber(okakaogift.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(okakaogift.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td>KakaoGift���ǰ��<br>��ϵȻ�ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">KakaoGift�����<br>KakaoGift����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">KakaoGift<br>�߰���ۺ�</td>
	<td width="70">KakaoGift<br>�߰��ݾ�</td>
	<td width="70">KakaoGift<br>���ݹ��Ǹ�</td>
	<td width="70">KakaoGift<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to okakaogift.FResultCount - 1 %>
<% itemidarr = itemidarr+CStr(okakaogift.FItemList(i).FItemID) + "," %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= okakaogift.FItemList(i).FItemID %>"></td>
	<td><img src="<%= okakaogift.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= okakaogift.FItemList(i).FItemID %>','kakaogift','<%=okakaogift.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=okakaogift.FItemList(i).FItemID%>" target="_blank"><%= okakaogift.FItemList(i).FItemID %></a>
		<% If okakaogift.FItemList(i).FkakaogiftStatcd <> 7 Then %>
		<br><%= okakaogift.FItemList(i).getkakaogiftStatName %>
		<% End If %>
		<%= okakaogift.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= okakaogift.FItemList(i).FMakerid %> <%= okakaogift.FItemList(i).getDeliverytypeName %><br><%= okakaogift.FItemList(i).FItemName %></td>
	<td align="left">
	    <%= okakaogift.FItemList(i).Fkakaoitemname %>
	    <br>
	    <%= okakaogift.FItemList(i).Fregitemname %>
	</td>
	<td align="center"><%= okakaogift.FItemList(i).FRegdate %><br><%= okakaogift.FItemList(i).FLastupdate %></td>
	<td align="center"><%= okakaogift.FItemList(i).FkakaogiftRegdate %><br><%= okakaogift.FItemList(i).FkakaogiftLastUpdate %></td>
	<td align="right">
		<% If okakaogift.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(okakaogift.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(okakaogift.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(okakaogift.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If okakaogift.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-okakaogift.FItemList(i).Fbuycash/okakaogift.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If okakaogift.FItemList(i).IsSoldOut Then
			If okakaogift.FItemList(i).FSellyn = "N" Then
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
		If okakaogift.FItemList(i).FItemdiv = "06" OR okakaogift.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center"><%= formatNumber(okakaogift.FItemList(i).FaddDlvPrice,0) %></td>
	<td align="center"><%= formatNumber(okakaogift.FItemList(i).FaddKakaoPrice,0) %></td>
	<td align="center">
	<%
		If (okakaogift.FItemList(i).FkakaogiftStatCd > 0) Then
			If Not IsNULL(okakaogift.FItemList(i).FkakaogiftPrice) Then
				If (okakaogift.FItemList(i).Mustprice <> okakaogift.FItemList(i).FkakaogiftPrice) Then
	%>
					<strong><%= formatNumber(okakaogift.FItemList(i).FkakaogiftPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(okakaogift.FItemList(i).FkakaogiftPrice,0)&"<br>"
				End If

				If (okakaogift.FItemList(i).FSellyn="Y" and okakaogift.FItemList(i).FkakaogiftSellYn<>"Y") or (okakaogift.FItemList(i).FSellyn<>"Y" and okakaogift.FItemList(i).FkakaogiftSellYn="Y") Then
	%>
					<strong><%= okakaogift.FItemList(i).FkakaogiftSellYn %></strong>
	<%
				Else
					response.write okakaogift.FItemList(i).FkakaogiftSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If okakaogift.FItemList(i).FkakaogiftGoodNo <> "" Then %>
			<a target="_blank" href="http://www.kakaogift.com/Detail?PrStCd=<%=okakaogift.FItemList(i).FkakaogiftGoodNo%>&ColorCd=ZZ9"><%=okakaogift.FItemList(i).FkakaogiftGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= okakaogift.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=okakaogift.FItemList(i).FItemID%>','0');"><%= okakaogift.FItemList(i).FoptionCnt %>:<%= okakaogift.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= okakaogift.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If okakaogift.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= okakaogift.FItemList(i).FinfoDiv %>
		<%
		If (okakaogift.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(okakaogift.FItemList(i).FlastErrStr) &"'>ERR:"& okakaogift.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if okakaogift.HasPreScroll then %>
		<a href="javascript:goPage('<%= okakaogift.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + okakaogift.StartScrollPage to okakaogift.FScrollCount + okakaogift.StartScrollPage - 1 %>
    		<% if i>okakaogift.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if okakaogift.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<textarea cols="80" rows="3"><%=itemidarr%></textarea>

<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->