<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/halfclub/halfclubcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, halfclubGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, halfclubYes10x10No, halfclubNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice
Dim page, i, research
Dim oHalfclub
Dim startMargin, endMargin, isextusing
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
halfclubGoodNo			= request("halfclubGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
halfclubYes10x10No		= request("halfclubYes10x10No")
halfclubNo10x10Yes		= request("halfclubNo10x10Yes")
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

'����Ŭ�� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If halfclubGoodNo <> "" then
	Dim iA2, arrTemp2, arrhalfclubGoodNo
	halfclubGoodNo = replace(halfclubGoodNo,",",chr(10))
	halfclubGoodNo = replace(halfclubGoodNo,chr(13),"")
	arrTemp2 = Split(halfclubGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrhalfclubGoodNo = arrhalfclubGoodNo & trim("'"&arrTemp2(iA2)&"'") & ","
		End If
		iA2 = iA2 + 1
	Loop
	halfclubGoodNo = left(arrhalfclubGoodNo,len(arrhalfclubGoodNo)-1)
End If

Set oHalfclub = new CHalfclub
	oHalfclub.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oHalfclub.FPageSize					= 100
Else
	oHalfclub.FPageSize					= 50
End If
	oHalfclub.FRectCDL					= request("cdl")
	oHalfclub.FRectCDM					= request("cdm")
	oHalfclub.FRectCDS					= request("cds")
	oHalfclub.FRectItemID				= itemid
	oHalfclub.FRectItemName				= itemname
	oHalfclub.FRectSellYn				= sellyn
	oHalfclub.FRectLimitYn				= limityn
	oHalfclub.FRectSailYn				= sailyn
	oHalfclub.FRectonlyValidMargin		= onlyValidMargin
	oHalfclub.FRectMakerid				= makerid
	oHalfclub.FRectHalfclubGoodNo		= halfclubGoodNo
	oHalfclub.FRectMatchCate			= MatchCate
	oHalfclub.FRectIsMadeHand			= isMadeHand
	oHalfclub.FRectIsOption				= isOption
	oHalfclub.FRectIsReged				= isReged
	oHalfclub.FRectNotinmakerid			= notinmakerid
	oHalfclub.FRectNotinitemid			= notinitemid
	oHalfclub.FRectExcTrans				= exctrans
	oHalfclub.FRectPriceOption			= priceOption
	oHalfclub.FRectIsSpecialPrice       = isSpecialPrice
	oHalfclub.FRectDeliverytype			= deliverytype
	oHalfclub.FRectMwdiv				= mwdiv

	oHalfclub.FRectExtNotReg			= ExtNotReg
	oHalfclub.FRectExpensive10x10		= expensive10x10
	oHalfclub.FRectdiffPrc				= diffPrc
	oHalfclub.FRectHalfClubYes10x10No	= halfclubYes10x10No
	oHalfclub.FRectHalfClubNo10x10Yes	= halfclubNo10x10Yes
	oHalfclub.FRectExtSellYn			= extsellyn
	oHalfclub.FRectInfoDiv				= infoDiv
	oHalfclub.FRectFailCntOverExcept	= ""
	oHalfclub.FRectFailCntExists		= failCntExists
	oHalfclub.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oHalfclub.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHalfclub.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oHalfclub.getHalfclubreqExpireItemList
Else
	oHalfclub.getHalfclubRegedItemList		'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=halfclub","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=halfclub','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="halfclubYes10x10No")&&(frm.halfclubYes10x10No.checked)){ frm.halfclubYes10x10No.checked=false }
	if ((comp.name!="halfclubNo10x10Yes")&&(frm.halfclubNo10x10Yes.checked)){ frm.halfclubNo10x10Yes.checked=false }
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

    if ((comp.name=="halfclubYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="halfclubNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.halfclubYes10x10No.checked){
            comp.form.halfclubYes10x10No.checked = false;
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
	if ((comp.name!="halfclubYes10x10No")&&(frm.halfclubYes10x10No.checked)){ frm.halfclubYes10x10No.checked=false }
	if ((comp.name!="halfclubNo10x10Yes")&&(frm.halfclubNo10x10Yes.checked)){ frm.halfclubNo10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/halfclub/pophalfclubcateList.asp","popCatehalfclubmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// ���õ� ��ǰ �ϰ� ���
function halfclubSelectRegProcess() {
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

    if (confirm('����Ŭ���� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n������Ŭ������ ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function halfclubSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function halfclubriceEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n������Ŭ������ ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function halfclubEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n������Ŭ������ ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function halfclubSelectDeliveryProcess() {

    if (confirm('�ù�� �ڵ带 ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "Delivery";
		document.frmSvArr.action = "<%=apiURL%>/outmall/halfclub/acthalfclubReq.asp"
		document.frmSvArr.submit();
    }
}

//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=halfclub&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
		<a href="http://scm.halfclub.com/Home/Default.aspx" target="_blank">����Ŭ��Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then
				response.write "<font color='GREEN'>[ 10x10 | store10x10** ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		����Ŭ�� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="halfclubGoodNo" id="itemid"><%=replace(halfclubGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >����Ŭ�� ��Ͻõ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >����Ŭ�� ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >����Ŭ�� ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>����Ŭ�� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="halfclubYes10x10No" <%= ChkIIF(halfclubYes10x10No="on","checked","") %> ><font color=red>����Ŭ���Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="halfclubNo10x10Yes" <%= ChkIIF(halfclubNo10x10Yes="on","checked","") %> ><font color=red>����Ŭ��ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>

<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����), ������ �ø�ó��(����Ŭ���� ������ �Ⱦ�)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ֹ����۹��� ��ǰ<br />

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
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('halfclub');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="halfclubSelectRegProcess();">
				<br><br>
				������ǰ ���� :
				<!--
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="halfclubriceEditProcess();">&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnStock" value="����" onClick="halfclubEditProcess();">&nbsp;&nbsp;
				<br><br>
				�����ڵ� ��ȸ :
				<input class="button" type="button" id="btnStock" value="�ù��" onClick="halfclubSelectDeliveryProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="halfclubSellYnProcess(frmReg.chgSellYn.value);">
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
		�˻���� : <b><%= FormatNumber(oHalfclub.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHalfclub.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">����Ŭ�������<br>����Ŭ������������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">����Ŭ��<br>���ݹ��Ǹ�</td>
	<td width="70">����Ŭ��<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oHalfclub.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHalfclub.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHalfclub.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHalfclub.FItemList(i).FItemID %>','halfclub','<%=oHalfclub.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oHalfclub.FItemList(i).FItemID%>" target="_blank"><%= oHalfclub.FItemList(i).FItemID %></a>
		<% If oHalfclub.FItemList(i).FHalfclubStatcd <> 7 Then %>
		<br><%= oHalfclub.FItemList(i).getHalfclubStatName %>
		<% End If %>
		<%= oHalfclub.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oHalfclub.FItemList(i).FMakerid %> <%= oHalfclub.FItemList(i).getDeliverytypeName %><br><%= oHalfclub.FItemList(i).FItemName %></td>
	<td align="center"><%= oHalfclub.FItemList(i).FRegdate %><br><%= oHalfclub.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHalfclub.FItemList(i).FHalfClubRegdate %><br><%= oHalfclub.FItemList(i).FHalfClubLastUpdate %></td>
	<td align="right">
		<% If oHalfclub.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oHalfclub.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oHalfclub.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oHalfclub.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oHalfclub.FItemList(i).Fsellcash = 0 Then
			'//
		elseIf (oHalfclub.FItemList(i).FHalfclubStatCd > 0) and Not IsNULL(oHalfclub.FItemList(i).FHalfClubPrice) Then
			If (oHalfclub.FItemList(i).FSaleYn = "Y") and (oHalfclub.FItemList(i).FSellcash < oHalfclub.FItemList(i).FHalfClubPrice) Then
				'// ���޸� ���� �Ǹ���
		%>
		<strike><%= CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		<font color="#CC3333"><%= CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).FHalfClubPrice*100*100)/100 & "%" %></font>
		<%
			else
				response.write CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%"
			end if
		else
			response.write CLng(10000-oHalfclub.FItemList(i).Fbuycash/oHalfclub.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oHalfclub.FItemList(i).IsSoldOut Then
			If oHalfclub.FItemList(i).FSellyn = "N" Then
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
		If oHalfclub.FItemList(i).FItemdiv = "06" OR oHalfclub.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oHalfclub.FItemList(i).FHalfclubStatCd > 0) Then
			If Not IsNULL(oHalfclub.FItemList(i).FHalfClubPrice) Then
				If (oHalfclub.FItemList(i).Mustprice <> oHalfclub.FItemList(i).FHalfClubPrice) Then
	%>
					<strong><%= formatNumber(oHalfclub.FItemList(i).FHalfClubPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oHalfclub.FItemList(i).FHalfClubPrice,0)&"<br>"
				End If

				If Not IsNULL(oHalfclub.FItemList(i).FSpecialPrice) Then
					If (now() >= oHalfclub.FItemList(i).FStartDate) And (now() <= oHalfclub.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oHalfclub.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oHalfclub.FItemList(i).FSellyn="Y" and oHalfclub.FItemList(i).FHalfClubSellYn<>"Y") or (oHalfclub.FItemList(i).FSellyn<>"Y" and oHalfclub.FItemList(i).FHalfClubSellYn="Y") Then
	%>
					<strong><%= oHalfclub.FItemList(i).FHalfClubSellYn %></strong>
	<%
				Else
					response.write oHalfclub.FItemList(i).FHalfClubSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oHalfclub.FItemList(i).FHalfclubGoodNo <> "" Then %>
			<a target="_blank" href="http://www.halfclub.com/Detail?PrStCd=<%=oHalfclub.FItemList(i).FHalfclubGoodNo%>&ColorCd=ZZ9"><%=oHalfclub.FItemList(i).FHalfclubGoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= oHalfclub.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHalfclub.FItemList(i).FItemID%>','0');"><%= oHalfclub.FItemList(i).FoptionCnt %>:<%= oHalfclub.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHalfclub.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oHalfclub.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oHalfclub.FItemList(i).FinfoDiv %>
		<%
		If (oHalfclub.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oHalfclub.FItemList(i).FlastErrStr) &"'>ERR:"& oHalfclub.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oHalfclub.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHalfclub.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHalfclub.StartScrollPage to oHalfclub.FScrollCount + oHalfclub.StartScrollPage - 1 %>
    		<% if i>oHalfclub.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHalfclub.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>

<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
