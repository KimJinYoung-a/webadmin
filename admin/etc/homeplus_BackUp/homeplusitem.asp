<%@ language=vbscript %>
<%' option Explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim oHomeplus, research, itemname, showminusmagin, expensive10x10, diffPrc, HomeplusYes10x10No, HomeplusNo10x10Yes, extsellyn, infoDiv
Dim i, page, itemid, sellyn, makerid, HomeplusGoodNo, limityn, sailyn, optAddprcExists, optAddprcExistsExcept, optAddPrcRegTypeNone, failCntExists
Dim HomeplusNotReg, MatchCate, dftMatchCate, onlyValidMargin, optExists, reqEdit, reqExpire, regedOptNull, optnotExists, isMadeHand
Dim bestOrd, bestOrdMall
page    = request("page")
itemid  = request("itemid")

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

makerid						= request("makerid")
itemname					= html2db(request("itemname"))
HomeplusNotReg				= request("HomeplusNotReg")
MatchCate					= request("MatchCate")
dftMatchCate				= request("dftMatchCate")
onlyValidMargin				= request("onlyValidMargin")
sellyn						= request("sellyn")
limityn						= request("limityn")
sailyn						= request("sailyn")
HomeplusGoodNo				= request("HomeplusGoodNo")
research					= request("research")
optExists					= request("optExists")
optnotExists				= request("optnotExists")
bestOrd						= request("bestOrd")
bestOrdMall					= request("bestOrdMall")
optAddprcExists				= request("optAddprcExists")
optAddPrcRegTypeNone		= request("optAddPrcRegTypeNone")
regedOptNull				= request("regedOptNull")
failCntExists				= request("failCntExists")

showminusmagin				= request("showminusmagin")
expensive10x10				= request("expensive10x10")
diffPrc						= request("diffPrc")
HomeplusYes10x10No			= request("HomeplusYes10x10No")
HomeplusNo10x10Yes			= request("HomeplusNo10x10Yes")
reqEdit						= request("reqEdit")
reqExpire					= request("reqExpire")
extsellyn					= request("extsellyn")
infoDiv						= request("infoDiv")
optAddprcExistsExcept		= request("optAddprcExistsExcept")
isMadeHand					= request("isMadeHand")

If page = "" Then page = 1
If sellyn = "" Then sellyn = "Y"

''�⺻���� ��Ͽ����̻�
If (research="") Then
	HomeplusNotReg = "J"
	MatchCate = ""
	dftMatchCate = ""
	onlyValidMargin="on"
	sellyn="Y"
'	optAddprcExistsExcept = "on"
'	limityn="N"
End If

Set oHomeplus = new CHomeplus
	oHomeplus.FPageSize 					= 20
	oHomeplus.FCurrPage						= page
	oHomeplus.FRectCDL						= request("cdl")
	oHomeplus.FRectCDM						= request("cdm")
	oHomeplus.FRectCDS						= request("cds")
	oHomeplus.FRectItemID					= itemid
	oHomeplus.FRectItemName					= itemname
	oHomeplus.FRectSellYn					= sellyn
	oHomeplus.FRectLimitYn					= limityn
	oHomeplus.FRectSailYn					= sailyn
	oHomeplus.FRectonlyValidMargin			= onlyValidMargin
	oHomeplus.FRectMakerid					= makerid
	oHomeplus.FRectHomeplusGoodNo			= HomeplusGoodNo
	oHomeplus.FRectMatchCate				= MatchCate
	oHomeplus.FRectdftMatchCate				= dftMatchCate

	oHomeplus.FRectoptExists				= optExists
	oHomeplus.FRectoptnotExists				= optnotExists
	oHomeplus.FRectHomeplusNotReg			= HomeplusNotReg
	oHomeplus.FRectMinusMigin				= showminusmagin
	oHomeplus.FRectExpensive10x10			= expensive10x10
	oHomeplus.FRectdiffPrc					= diffPrc
	oHomeplus.FRectHomeplusYes10x10No		= HomeplusYes10x10No
	oHomeplus.FRectHomeplusNo10x10Yes		= HomeplusNo10x10Yes
	oHomeplus.FRectExtSellYn				= extsellyn
	oHomeplus.FRectInfoDiv					= infoDiv
	oHomeplus.FRectFailCntOverExcept		= ""
	oHomeplus.FRectoptAddprcExists			= optAddprcExists
	oHomeplus.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oHomeplus.FRectoptAddPrcRegTypeNone		= optAddPrcRegTypeNone                         ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
	oHomeplus.FRectregedOptNull				= regedOptNull
	oHomeplus.FRectFailCntExists			= failCntExists
	oHomeplus.FRectisMadeHand				= isMadeHand
If (bestOrd = "on") Then
    oHomeplus.FRectOrdType					 = "B"
ElseIf (bestOrdMall = "on") Then
    oHomeplus.FRectOrdType					= "BM"
End If

If reqExpire <> "" Then
	oHomeplus.getHomeplusreqExpireItemList
Else
	oHomeplus.getHomeplusRegedItemList
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=homeplus","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin2=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=homeplus','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//ī�װ� ����
function pop_prdDivManager() {
	var pCM1 = window.open("/admin/etc/homeplus/pophomeplusprdDivList.asp","popCatehomeplus","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM1.focus();
}
//ī�װ� ����
function pop_cateManager() {
	var pCM2 = window.open("/admin/etc/homeplus/pophomepluscateList.asp","popCatehomeplusmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
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
//      document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "RegSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
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
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ���� ����
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
        document.frmSvArr.cmdparam.value = "EditItemSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
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
        document.frmSvArr.cmdparam.value = "EditImgSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �̹��� ����
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
        document.frmSvArr.cmdparam.value = "ViewSelect";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
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
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}

function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.HomeplusNotReg.value="M";
            frm.MatchCate.value="Y";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="reqEdit"){
            frm.HomeplusNotReg.value="R";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.HomeplusNotReg.value="D";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="HomeplusNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="HomeplusYes10x10No")&&(frm.HomeplusYes10x10No.checked)){ frm.HomeplusYes10x10No.checked=false }
        if ((comp.name!="HomeplusNo10x10Yes")&&(frm.HomeplusNo10x10Yes.checked)){ frm.HomeplusNo10x10Yes.checked=false }
        if ((comp.name!="reqREG")&&(frm.reqREG.checked)){ frm.reqREG.checked=false }
        if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
        if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
    }
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_homeplus.asp?itemid="+iitemid+'&mallid=homeplus&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

function checkComp(comp){
    if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
        if ((comp.name=="bestOrd")&&(comp.checked)){
            comp.form.bestOrdMall.checked=false;
        }

        if ((comp.name=="bestOrdMall")&&(comp.checked)){
            comp.form.bestOrd.checked=false;
        }
    }else if ((comp.name=="optAddprcExists")||(comp.name=="optAddprcExistsExcept")){
        if ((comp.name=="optAddprcExists")&&(comp.checked)){
            comp.form.optAddprcExistsExcept.checked=false;
        }

        if ((comp.name=="optAddprcExistsExcept")&&(comp.checked)){
            comp.form.optAddprcExists.checked=false;
        }
    }
}

function homeplusCateAPI() {
    if (confirm('Ȩ�÷��� ī�װ�API�� �����Ͻðڽ��ϱ�?\n������ ��ϵ� ī�װ��� ������ �� �ֽ��ϴ�.')){
    	document.getElementById("btncate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CategoryView";
        document.frmSvArr.action = "/admin/etc/homeplus/acthomeplusReq.asp"
        document.frmSvArr.submit();
    }
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		�� �� �� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		Homeplus��ǰ��ȣ: <input type="text" name="HomeplusGoodNo" value="<%= HomeplusGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;&nbsp;
		��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		<a href="https://bos.homeplus.co.kr:446/LoginForm.jsp" target="_blank">Homeplus Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[  292811 | cube1010!! ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ��ȣ: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		�ֹ����ۿ��� :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		��Ͽ��� :
		<select name="HomeplusNotReg" class="select">
			<option value="">��ü
			<option value="M" <%= CHkIIF(HomeplusNotReg="M","selected","") %> >Homeplus �̵��(��ϰ���)
			<option value="Q" <%= CHkIIF(HomeplusNotReg="Q","selected","") %> >Homeplus ��Ͻ���
			<option value="J" <%= CHkIIF(HomeplusNotReg="J","selected","") %> >Homeplus ��Ͽ����̻�
			<option value="A" <%= CHkIIF(HomeplusNotReg="A","selected","") %> >Homeplus ���۽õ��߿���
			<option value="D" <%= CHkIIF(HomeplusNotReg="D","selected","") %> >Homeplus ��ϿϷ�(����)
			<option value="R" <%= CHkIIF(HomeplusNotReg="R","selected","") %> >Homeplus �������
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>
		&nbsp;
		����ī�׸�Ī :
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>
		&nbsp;
		����ī�׸�Ī :
		<select name="dftMatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(dftMatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(dftMatchCate="N","selected","") %> >�̸�Ī
		</select>
		&nbsp;
		�Ǹſ��� :
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
		</select>
		&nbsp;
		�������� :
		<select name="limityn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
		</select>
		&nbsp;
		���Ͽ��� :
		<select name="sailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
		</select>

		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� <%= CMAXMARGIN %>%�̻� ��ǰ�� ����
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ
		&nbsp;
		<input type="checkbox" name="optAddPrcRegTypeNone" <%= ChkIIF(optAddPrcRegTypeNone="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ǸŹ̼�����ǰ
		&nbsp;
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ����
		&nbsp;

		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >�ɼ������ǰ
		&nbsp;
		<input type="checkbox" name="optnotExists" <%= ChkIIF(optnotExists="on","checked","") %> >��ǰ��ǰ(�ɼ�=0)
		&nbsp;
		<input type="checkbox" name="regedOptNull" <%= ChkIIF(regedOptNull="on","checked","") %> >��ǰ��� �̼���
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >��ϼ���������ǰ
		<br><br>
		-- Quick �˻� / ��� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >��ϰ��� ��ǰ
		<br><br>
		-- Quick �˻� / ���� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>������</font>��ǰ���� (MaxMagin : <%= CMAXMARGIN %>%) (Homeplus �Ǹ���)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusYes10x10No" <%= ChkIIF(HomeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusNo10x10Yes" <%= ChkIIF(HomeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplusǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (���޸� �����Ե�)
		&nbsp;&nbsp;�����ǸŻ��� :
		<select name="extsellyn" class="select">
			<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
		</select>
		&nbsp;&nbsp;ǰ������ :
		<% CALL DrawItemInfoDiv("infoDiv", infoDiv, true, "") %>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<br>
<form name="frmReg" method="post" action="homeplusItem.asp" style="margin:0px;">
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
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="���� �� ����ī�װ�" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="���� ī�װ�" onclick="pop_cateManager();">
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
	    		<input class="button" type="button" id="btnRegSel" value="��ǰ ���" onClick="HomeplusSelectRegProcess();">
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
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="18" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oHomeplus.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oHomeplus.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="80">�ٹ�����<br>��ǰ��ȣ</td>
	<td >�귣��<br>��ǰ��</td>
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
<tr bgcolor="#FFFFFF" height="20">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHomeplus.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oHomeplus.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHomeplus.FItemList(i).FItemID %>','Homeplus','')" style="cursor:pointer"></td>
	<td align="center"><%= oHomeplus.FItemList(i).FItemID %><br>
	<% If oHomeplus.FItemList(i).FLimitYn= "Y" Then %><%= oHomeplus.FItemList(i).getLimitHtmlStr %></font><% End If %>
	</td>
	<td><%= oHomeplus.FItemList(i).FMakerid %> <%= oHomeplus.FItemList(i).getDeliverytypeName %><br><%= oHomeplus.FItemList(i).FItemName %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FRegdate %><br><%= oHomeplus.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHomeplus.FItemList(i).FHomeplusRegdate %><br><%= oHomeplus.FItemList(i).FHomeplusLastUpdate %></td>
	<td align="right">
	    <% If oHomeplus.FItemList(i).FSaleYn = "Y" Then %>
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
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="600"></iframe>
<% Set oHomeplus = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->