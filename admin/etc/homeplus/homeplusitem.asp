<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, homeplusGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, dftMatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid
Dim expensive10x10, diffPrc, homeplusYes10x10No, homeplusNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oHomeplus

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
homeplusGoodNo			= request("homeplusGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
dftMatchCate			= request("dftMatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
homeplusYes10x10No		= request("homeplusYes10x10No")
homeplusNo10x10Yes		= request("homeplusNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")

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
'Homeplus ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If homeplusGoodNo <> "" then
	Dim iA2, arrTemp2, arrhomeplusGoodNo
	homeplusGoodNo = replace(homeplusGoodNo,",",chr(10))
	homeplusGoodNo = replace(homeplusGoodNo,chr(13),"")
	arrTemp2 = Split(homeplusGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrhomeplusGoodNo = arrhomeplusGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	homeplusGoodNo = left(arrhomeplusGoodNo,len(arrhomeplusGoodNo)-1)
End If

SET oHomeplus = new CHomeplus
	oHomeplus.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oHomeplus.FPageSize					= 50
Else
	oHomeplus.FPageSize					= 20
End If
	oHomeplus.FRectCDL					= request("cdl")
	oHomeplus.FRectCDM					= request("cdm")
	oHomeplus.FRectCDS					= request("cds")
	oHomeplus.FRectItemID				= itemid
	oHomeplus.FRectItemName				= itemname
	oHomeplus.FRectSellYn				= sellyn
	oHomeplus.FRectLimitYn				= limityn
	oHomeplus.FRectSailYn				= sailyn
	oHomeplus.FRectonlyValidMargin		= onlyValidMargin
	oHomeplus.FRectMakerid				= makerid
	oHomeplus.FRectHomeplusGoodNo		= homeplusGoodNo
	oHomeplus.FRectMatchCate			= MatchCate
	oHomeplus.FRectDftMatchCate			= dftMatchCate
	oHomeplus.FRectIsMadeHand			= isMadeHand
	oHomeplus.FRectIsOption				= isOption
	oHomeplus.FRectIsReged				= isReged
	oHomeplus.FRectNotinmakerid			= notinmakerid

	oHomeplus.FRectExtNotReg			= ExtNotReg
	oHomeplus.FRectExpensive10x10		= expensive10x10
	oHomeplus.FRectdiffPrc				= diffPrc
	oHomeplus.FRectHomeplusYes10x10No	= homeplusYes10x10No
	oHomeplus.FRectHomeplusNo10x10Yes	= homeplusNo10x10Yes
	oHomeplus.FRectExtSellYn			= extsellyn
	oHomeplus.FRectInfoDiv				= infoDiv
	oHomeplus.FRectFailCntOverExcept	= ""
	oHomeplus.FRectFailCntExists		= failCntExists
	oHomeplus.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oHomeplus.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHomeplus.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oHomeplus.getHomeplusreqExpireItemList
Else
	oHomeplus.getHomeplusRegedItemList			'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=homeplus","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=homeplus','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_prdDivManager() {
	var pCM1 = window.open("/admin/etc/homeplus/pophomeplusprdDivList.asp","popCatehomeplus","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM1.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/homeplus/pophomepluscateList.asp","popCatehomeplusmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
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
	if ((comp.name!="homeplusYes10x10No")&&(frm.homeplusYes10x10No.checked)){ frm.homeplusYes10x10No.checked=false }
	if ((comp.name!="homeplusNo10x10Yes")&&(frm.homeplusNo10x10Yes.checked)){ frm.homeplusNo10x10Yes.checked=false }
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

    if ((comp.name=="homeplusYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="homeplusNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.homeplusYes10x10No.checked){
            comp.form.homeplusYes10x10No.checked = false;
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
	if ((comp.name!="homeplusYes10x10No")&&(frm.homeplusYes10x10No.checked)){ frm.homeplusYes10x10No.checked=false }
	if ((comp.name!="homeplusNo10x10Yes")&&(frm.homeplusNo10x10Yes.checked)){ frm.homeplusNo10x10Yes.checked=false }
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
function homeplusCateAPI() {
    if (confirm('Ȩ�÷��� ī�װ�API�� �����Ͻðڽ��ϱ�?\n������ ��ϵ� ī�װ��� ������ �� �ֽ��ϴ�.')){
    	document.getElementById("btncate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CategoryView";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ���
function HomeplusSelectRegProcess() {
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

    if (confirm('Homeplus�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��Homeplus���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �Ǹſ��� ����
function HomeplusSellYnProcess(chkYn) {
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
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Homeplus���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ���� ����
function HomeplusSelectEditProcess() {
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

    if (confirm('Homeplus�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n�ػ�ǰ��, ī�װ�, �������, �̹���, ��ǰ���� ���� �����˴ϴ�.\n\n�ػ�ǰ����, �ǸŻ��´� �������� �ʽ��ϴ�.')){
		document.getElementById("btnEditSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "ITEMNAME";
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ������ ����
function HomeplusSelectEditItemProcess() {
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

    if (confirm('Homeplus�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n���ش� ������ �߰�/����/�Ǹ���/�Ǹ�����/���� ������ ����˴ϴ�.\n\n�ػ�ǰ��, ī�װ�, �������, �̹���, ��ǰ������ �������� �ʽ��ϴ�.')){
        document.getElementById("btnEditOptSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �̹��� ����
function HomeplusSelectImgEditProcess() {
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

    if (confirm('Homeplus�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �̹����� ���� �Ͻðڽ��ϱ�?\n\n��Homeplus���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditImgSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditImg";
        document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ��ȸ
function HomeplusSelectViewProcess() {
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

    if (confirm('Homeplus�� �����Ͻ� ' + chkSel + '�� ��ǰ������ ��ȸ �Ͻðڽ��ϱ�?\n\n��Homeplus���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=homeplus&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
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
		<a href="https://bos.homeplus.co.kr:446/LoginForm.jsp" target="_blank">Homeplus Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  292811 | tenbyten10*$ ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		Homeplus ��ǰ�ڵ� : <textarea rows="2" cols="20" name="homeplusGoodNo" id="itemid"><%=replace(homeplusGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Homeplus ��Ͻ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >Homeplus ���۽õ��� ����
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >Homeplus ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Homeplus ��ϿϷ�(����)
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
		����ī�װ�
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;

		����ī�װ�
		<select name="dftMatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(dftMatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(dftMatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="homeplusYes10x10No" <%= ChkIIF(homeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="homeplusNo10x10Yes" <%= ChkIIF(homeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplusǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="homeplusitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
			<% If (session("ssBctID")="kjy8517") Then %>
				<input class="button" type="button" id="btncate" value="ī�װ�API" onclick="homeplusCateAPI();"> &nbsp;
			<% End If %>
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();"> &nbsp;
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('homeplus');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="���� �� ����ī�װ�" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="���� ī�װ�" onclick="pop_CateManager();">
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
				������ǰ ��� : <input class="button" type="button" id="btnRegSel" value="���" onClick="HomeplusSelectRegProcess();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="���� ����" onClick="HomeplusSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditOptSel" value="������ ����" onClick="HomeplusSelectEditItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditImgSel" value="�̹��� ����" onClick="HomeplusSelectImgEditProcess();">&nbsp;&nbsp;
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="cogusdk") Then %>
				<br><br>
				������ǰ ��ȸ :
				<input class="button" type="button" id="btnViewSel" value="���� ��ȸ" onClick="HomeplusSelectViewProcess();">&nbsp;&nbsp;
				<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="HomeplusSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<!-- �׼� �� -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oHomeplus.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHomeplus.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">Homeplus�����<br>Homeplus����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Homeplus<br>���ݹ��Ǹ�</td>
	<td width="70">Homeplus<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="100">Homeplus<br>����ī�װ�</td>
	<td width="100">Homeplus<br>����ī�װ�</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oHomeplus.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHomeplus.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHomeplus.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHomeplus.FItemList(i).FItemID %>','homeplus','<%=oHomeplus.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oHomeplus.FItemList(i).FItemID%>" target="_blank"><%= oHomeplus.FItemList(i).FItemID %></a>
		<% If oHomeplus.FItemList(i).FHomeplusStatCd <> 7 Then %>
		<br><%= oHomeplus.FItemList(i).getHomeplusItemStatCd %>
		<% End If %>
		<% If oHomeplus.FItemList(i).FLimitYn= "Y" Then %><br><%= oHomeplus.FItemList(i).getLimitHtmlStr %></font><% End If %>
	</td>
	<td align="left"><%= oHomeplus.FItemList(i).FMakerid %><%= oHomeplus.FItemList(i).getDeliverytypeName %><br><%= oHomeplus.FItemList(i).FItemName %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FRegdate %><br><%= oHomeplus.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FHomeplusRegdate %><br><%= oHomeplus.FItemList(i).FHomeplusLastUpdate %></td>
	<td align="right">
	<% If oHomeplus.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oHomeplus.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oHomeplus.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oHomeplus.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	    <% If oHomeplus.FItemList(i).Fsellcash <> 0 Then %>
	    <%= CLng(10000-oHomeplus.FItemList(i).Fbuycash/oHomeplus.FItemList(i).Fsellcash*100*100)/100 %> %
	    <% End If %>
	</td>
	<td align="center">
	    <% If oHomeplus.FItemList(i).IsSoldOut Then %>
	        <% If oHomeplus.FItemList(i).FSellyn = "N" Then %>
	        <font color="red">ǰ��</font>
	        <% Else %>
	        <font color="red">�Ͻ�<br>ǰ��</font>
	        <% End If %>
	    <% End If %>
	</td>
	<td align="center">
	<%
		If oHomeplus.FItemList(i).FItemdiv = "06" OR oHomeplus.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<% If (oHomeplus.FItemList(i).FHomeplusStatCd > 0) Then %>
	<% If Not IsNULL(oHomeplus.FItemList(i).FHomeplusPrice) Then %>
	    <% If (oHomeplus.FItemList(i).Fsellcash<>oHomeplus.FItemList(i).FHomeplusPrice) Then %>
	    <strong><%= formatNumber(oHomeplus.FItemList(i).FHomeplusPrice,0) %></strong>
	    <% Else %>
	    <%= formatNumber(oHomeplus.FItemList(i).FHomeplusPrice,0) %>
	    <% End If %>
	    <br>
	    <% If (oHomeplus.FItemList(i).FSellyn<>oHomeplus.FItemList(i).FHomeplusSellYn) Then %>
	    <strong><%= oHomeplus.FItemList(i).FHomeplusSellYn %></strong>
	    <% Else %>
	    <%= oHomeplus.FItemList(i).FHomeplusSellYn %>
	    <% End If %>
	<% End If %>
	<% End If %>
	</td>
	<td align="center">
	<%
		'#�ǻ�ǰ��ȣ
		If Not(IsNULL(oHomeplus.FItemList(i).FHomeplusGoodNo)) Then
	    	Response.Write "<span style='cursor:pointer;' onclick=window.open('http://direct.homeplus.co.kr/app.product.Product.ghs?comm=usr.product.detail&i_style="&oHomeplus.FItemList(i).FHomeplusGoodNo&"')>"&oHomeplus.FItemList(i).FHomeplusGoodNo&"</span>"
		Else
			Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"& CHKIIF(oHomeplus.FItemList(i).FHomeplusStatCd="0","(��Ͽ���)","")
		End If
	%>
	</td>
	<td align="center"><%= oHomeplus.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHomeplus.FItemList(i).FItemID%>','0');"><%= oHomeplus.FItemList(i).FoptionCnt %>:<%= oHomeplus.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHomeplus.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oHomeplus.FItemList(i).FCateMapCnt > 0 Then %>
		<font color='BLUE'>��Ī��</font>
	<% Else %>
		<font color="darkred">��Ī�ȵ�</font>
	<% End If %>

	<% If (oHomeplus.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oHomeplus.FItemList(i).FlastErrStr %>">ERR:<%= oHomeplus.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
	<td align="center">
	<%
		If oHomeplus.FItemList(i).FhDIVISION = "" Then
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		Else
			response.write "<font color='BLUE'>��Ī��</font>"
		End If
	%>
	</td>
	<td align="center"><%= oHomeplus.FItemList(i).FinfoDiv %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oHomeplus.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHomeplus.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHomeplus.StartScrollPage to oHomeplus.FScrollCount + oHomeplus.StartScrollPage - 1 %>
    		<% if i>oHomeplus.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHomeplus.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
