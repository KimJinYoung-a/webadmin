<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/shoplinkercls.asp"-->
<!-- #include virtual="/admin/etc/shoplinker/incShoplinkerFunction.asp"-->
<%
Dim itemid, itemname,  mode
Dim page, makerid, ShoplinkerNotReg, sellyn, limityn
Dim delitemid, ShoplinkerGoodNo, showminusmagin, expensive10x10, ShoplinkerYes10x10No, ShoplinkerNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists, failCntExists, optAddPrcRegTypeNone
Dim bestOrd, bestOrdMall, extsellyn, infoDiv
Dim i
page    = request("page")
itemid  = request("itemid")
If itemid <> "" Then
	If Right(itemid,1) = "," OR Right(itemid,1) = " " Then
		Response.Write "<script>alert('��ǰ�ڵ尡 �߸� �ԷµǾ����ϴ�.');history.back();</script>"
		Response.End
	End IF
End IF

itemname				= html2db(request("itemname"))
mode					= request("mode")
makerid					= request("makerid")
ShoplinkerNotReg		= request("ShoplinkerNotReg")
sellyn					= request("sellyn")
limityn					= request("limityn")
delitemid				= request("delitemid")
ShoplinkerGoodNo		= request("ShoplinkerGoodNo")
showminusmagin			= request("showminusmagin")
expensive10x10			= request("expensive10x10")
ShoplinkerYes10x10No	= request("ShoplinkerYes10x10No")
ShoplinkerNo10x10Yes	= request("ShoplinkerNo10x10Yes")
onreginotmapping		= request("onreginotmapping")
diffPrc					= request("diffPrc")
onlyValidMargin			= request("onlyValidMargin")
research				= request("research")
reqExpire				= request("reqExpire")
reqEdit					= request("reqEdit")
optAddprcExists			= request("optAddprcExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
optAddprcExistsExcept	= request("optAddprcExistsExcept")
optExists				= request("optExists")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
failCntExists			= request("failCntExists")
extsellyn				= request("extsellyn")
infoDiv					= request("infoDiv")
mall_name				= request("mall_name")

If page = "" Then page = 1
If sellyn = "" Then sellyn = "Y"

''�⺻���� ��Ͽ����̻�
If (research="") Then
    ShoplinkerNotReg = "M"		'J
    onlyValidMargin = "on"
    bestOrd="on"
    sellyn = "Y"
End If

Dim oshoplinker
SET oshoplinker = new CShoplinker
If (ShoplinkerNotReg="F") then                       '''���δ��
	oshoplinker.FPageSize 					= 20
Else
	oshoplinker.FPageSize 					= 20
End If
	oshoplinker.FCurrPage					= page
	oshoplinker.FRectItemID					= itemid
	oshoplinker.FRectItemName				= itemname
	oshoplinker.FRectMakerid				= makerid
	oshoplinker.FRectCDL					= request("cdl")
	oshoplinker.FRectCDM					= request("cdm")
	oshoplinker.FRectCDS					= request("cds")
	oshoplinker.FRectShoplinkerNotReg		= ShoplinkerNotReg
	oshoplinker.FRectSellYn					= sellyn
	oshoplinker.FRectLimitYn				= limityn
	oshoplinker.FRectShoplinkerGoodNo		= ShoplinkerGoodNo
	oshoplinker.FRectMinusMigin				= showminusmagin
	oshoplinker.FRectExpensive10x10			= expensive10x10
	oshoplinker.FRectShoplinkerYes10x10No	= ShoplinkerYes10x10No
	oshoplinker.FRectShoplinkerNo10x10Yes	= ShoplinkerNo10x10Yes
	oshoplinker.FRectOnreginotmapping		= onreginotmapping
	oshoplinker.FRectdiffPrc				= diffPrc
	oshoplinker.FRectonlyValidMargin		= onlyValidMargin
	oshoplinker.FRectoptAddprcExists		= optAddprcExists
	oshoplinker.FRectoptAddPrcRegTypeNone	= optAddPrcRegTypeNone                         ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
	oshoplinker.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oshoplinker.FRectoptExists				= optExists
	oshoplinker.FRectFailCntExists			= failCntExists
	oshoplinker.FRectFailCntOverExcept		= ""
	oshoplinker.FRectExtSellYn				= extsellyn
	oshoplinker.FRectInfoDiv				= infoDiv
	oshoplinker.FRectMall_name				= mall_name

If (bestOrd = "on") Then
    oshoplinker.FRectOrdType				 = "B"
ElseIf (bestOrdMall = "on") Then
    oshoplinker.FRectOrdType				= "BM"
End If

If reqExpire <> "" Then
	oshoplinker.getShoplinkerreqExpireItemList
Else
	oshoplinker.getShoplinkerRegedItemList
End If
%>
<script language="javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.shoplinkerNotReg.value="M";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="reqEdit"){
            frm.shoplinkerNotReg.value="R";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.shoplinkerNotReg.value="D";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="shoplinkerNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="shoplinkerYes10x10No")&&(frm.shoplinkerYes10x10No.checked)){ frm.shoplinkerYes10x10No.checked=false }
        if ((comp.name!="shoplinkerNo10x10Yes")&&(frm.shoplinkerNo10x10Yes.checked)){ frm.shoplinkerNo10x10Yes.checked=false }
        if ((comp.name!="reqREG")&&(frm.reqREG.checked)){ frm.reqREG.checked=false }
        if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
        if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
    }
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

// ���õ� ��ǰ �ϰ� ���
function ShoplinkerSelectRegProcess(isreal) {

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

	if(document.getElementById('categbn').value == ""){
		alert('��ī�װ������� �����ϼ���');
		document.getElementById('categbn').focus();
		return;
	}

    if (isreal){
		if (confirm('Shoplinker�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?\n\Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.getElementById("btnRegSelR").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelect";
			document.frmSvArr.subcmd.value = document.getElementById('categbn').value;
			document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
			document.frmSvArr.submit();
		}
	}else{
		if (confirm('Shoplinker�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
			document.getElementById("btnRegSelR2").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "EDTSelect";
			document.frmSvArr.subcmd.value = document.getElementById('categbn').value;
			document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
			document.frmSvArr.submit();
		}
	}
}

function ShoplinkerSelectRegPoomOKProcess() {
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

	if (confirm('Shoplinker�� �����Ͻ� ' + chkSel + '�� ǰ�������� �ϰ� ��� �Ͻðڽ��ϱ�?\n\Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSelP").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "RegPoom";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}

function OutmallSelectEditProcess() {
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

	if (confirm('Shoplinker�� ������ �����Ͻ� ' + chkSel + '�� �ܺθ� ������ ���� �Ͻðڽ��ϱ�?\n\Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditOut").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditOutMall";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}

function SelectItemCDSearch(){
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

	if (confirm('Shoplinker�� �����Ͻ� ' + chkSel + '�� ��ǰ�ڵ� ��ȸ �Ͻðڽ��ϱ�?\n\Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSelS").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "SearchITEM";
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
	}
}
function ShoplinkerSellYnProcess(chkYn){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��Shoplinker���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/admin/etc/shoplinker/actShoplinkerReq.asp"
		document.frmSvArr.submit();
    }
}
function divch(divid, itemid){
	document.frmSvArr.cmdparam.value = divid;
	document.frmSvArr.subcmd.value = itemid;
	document.frmSvArr.target="xLink";
	document.frmSvArr.action='shoplinker_Outmallsearch.asp';
	document.frmSvArr.submit();
}
function OutmallSetting(){
	var popwin=window.open('/admin/etc/shoplinker/popOutmallsetting.asp','notin','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� �귣��
function NotInMakerid(){
	var popwin=window.open('/admin/etc/shoplinker/JaehyuMall_Not_In_Makerid.asp?mallgubun=shoplinker','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ��ǰ
/*
function NotInItemid(){
	var popwin=window.open('/admin/etc/shoplinker/JaehyuMall_Not_In_Itemid.asp?mallgubun=shoplinker','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
*/
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		�� �� �� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		����Ŀ��ǰ��ȣ: <input type="text" name="shoplinkerGoodNo" value="<%= shoplinkerGoodNo %>" size="15" maxlength="15" class="text"> &nbsp;&nbsp;
		��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		&nbsp;
		<a href="http://ad2.shoplinker.co.kr/" target="_blank">����ĿAdmin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 10x10 | cube1010 ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ��ȣ: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		<br>
		��Ͽ��� :
		<select name="shoplinkerNotReg" class="select">
		<option value="">��ü
		<option value="M" <%= CHkIIF(shoplinkerNotReg="M","selected","") %> >����Ŀ �̵��(��ϰ���)
		<option value="Q" <%= CHkIIF(shoplinkerNotReg="Q","selected","") %> >����Ŀ ��Ͻ���
		<option value="J" <%= CHkIIF(shoplinkerNotReg="J","selected","") %> >����Ŀ ��Ͻõ�+��ϿϷ�
		<option value="A" <%= CHkIIF(shoplinkerNotReg="A","selected","") %> >����Ŀ ��Ͻõ��߿���
		<option value="F" <%= CHkIIF(shoplinkerNotReg="F","selected","") %> >����Ŀ ��ϿϷ�(�ܺθ� �̿���)
		<option value="D" <%= CHkIIF(shoplinkerNotReg="D","selected","") %> >����Ŀ ��ϿϷ�(�ܺθ� ����)
		<option value="R" <%= CHkIIF(shoplinkerNotReg="R","selected","") %> >����Ŀ �������
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>
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

		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� <%= CMAXMARGIN %>%�̻� ��ǰ�� ����
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ
		&nbsp;
<!--	<input type="checkbox" name="optAddPrcRegTypeNone" <%= ChkIIF(optAddPrcRegTypeNone="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ǸŹ̼�����ǰ &nbsp; -->
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ����
		&nbsp;

		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >�ɼ������ǰ
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >��ϼ���������ǰ
		<br><br>
		-- Quick �˻� / ��� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >��ϰ��� ��ǰ
		<br><br>
		-- Quick �˻� / ���� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>������</font>��ǰ���� (MaxMagin : <%= CMAXMARGIN %>%) (����Ŀ �Ǹ���)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>����Ŀ ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="shoplinkerYes10x10No" <%= ChkIIF(shoplinkerYes10x10No="on","checked","") %> ><font color=red>����Ŀ�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="shoplinkerNo10x10Yes" <%= ChkIIF(shoplinkerNo10x10Yes="on","checked","") %> ><font color=red>����Ŀǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����
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
		&nbsp;&nbsp;�ܺθ� :
		<% CALL DrawShoplinkerOutmall("mall_name", mall_name, true, "") %>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a">
   		��ī�װ� ���� : 
   		<select name="categbn" id="categbn" class="select">
   			<option value="">--CHOICE--</option>
   			<option value="10x10">10x10(10x10)</option>
   			<option value="Sourcing">Sourcing(10x10)</option>
   			<option value="ithinkso">ithinkso(ithinkso)</option>
   		</select>&nbsp;
   		���� 5���̸��̳� ��� : 
   		<input type="button" class="button" value="���ܺ귣��" onclick="NotInMakerid();">&nbsp;
   		<!--
   		���� 5���̸��̳� ��� : 
   		<input type="button" class="button" value="���ܻ�ǰ" onclick="NotInItemid();">
   		-->
   		<br><br>
   		����Ŀ ��ǰ ���� :
   		<input class="button" type="button" id="btnRegSelR" value="��ǰ���" onClick="ShoplinkerSelectRegProcess(true);">&nbsp
   		<input class="button" type="button" id="btnRegSelR2" value="��ǰ����" onClick="ShoplinkerSelectRegProcess(false);">&nbsp
   		<input class="button" type="button" id="btnRegSelP" value="ǰ����/����" onClick="ShoplinkerSelectRegPoomOKProcess();">&nbsp
   		<input class="button" type="button" id="btnRegSelS" value="���θ� ��ǰ�ڵ� ��ȸ" onClick="SelectItemCDSearch();">
   		<br><br>
   		�ܺθ� ��ǰ ���� :
   		<input class="button" type="button" id="btnEditOut" value="��ǰ����" onClick="OutmallSelectEditProcess();">&nbsp
	</td>
	<td align="right" valign="top" class="a">
		<font Color= "BLUE">�� ���� ��ư�� Ŭ���Ͽ� �̸� �������ּž� �մϴ�.</font>
		<input class="button" type="button" id="btnOutmallSet" value="�ƿ�������" onClick="OutmallSetting()">
		<br><br>
		���û�ǰ��
		<Select name="chgSellYn" class="select">
			<option value="N">ǰ��</option>
			<option value="Y">�Ǹ���</option>
		</Select>(��)��
		<input class="button" type="button" id="btnSellYn" value="����" onClick="ShoplinkerSellYnProcess(frm.chgSellYn.value);">
		<br><br>
		<a href='/admin/etc/shoplinker/201305.xls' onfocus='this.blur()'><font color="RED">*����Ŀ���޸�����_�����ٿ�</font>
    </td>
</tr>
</table>
</form>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oshoplinker.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oshoplinker.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td >�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">����Ŀ �����<br>����Ŀ ����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">����Ŀ<br>���ݹ��Ǹ�</td>
	<td width="120">����Ŀ ��ǰ��ȣ<br>(�ӽù�ȣ)</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">ǰ��</td>
	<td width="80">����Ŀ<br>ǰ����YN</td>
	<td width="80">�ܺθ�����YN</td>
</tr>
<%
If oshoplinker.FResultCount > 0 Then
	For i=0 to oshoplinker.FResultCount - 1
%>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oshoplinker.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oshoplinker.FItemList(i).Fsmallimage %>" width="50"</td>
    <td align="center"><%= oshoplinker.FItemList(i).FItemID %><br><%= oshoplinker.FItemList(i).getShoplinkerItemStatCd %>
    <% If oshoplinker.FItemList(i).FLimitYn = "Y" Then %><br><%= oshoplinker.FItemList(i).getLimitHtmlStr %></font><% end if %>
    </td>
    <td><%= oshoplinker.FItemList(i).FMakerid %> <%= oshoplinker.FItemList(i).getDeliverytypeName %><br><%= oshoplinker.FItemList(i).FItemName %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FRegdate %><br><%= oshoplinker.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FShoplinkerRegdate %><br><%= oshoplinker.FItemList(i).FShoplinkerLastUpdate %></td>
    <td align="right">
        <% If oshoplinker.FItemList(i).FSaleYn = "Y" Then %>
        <strike><%= FormatNumber(oshoplinker.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oshoplinker.FItemList(i).FSellcash,0) %></font>
        <% Else %>
        <%= FormatNumber(oshoplinker.FItemList(i).FSellcash,0) %>
        <% End If %>
    </td>
    <td align="center">
        <% If oshoplinker.FItemList(i).Fsellcash <> 0 Then %>
        <%= CLng(10000-oshoplinker.FItemList(i).Fbuycash/oshoplinker.FItemList(i).Fsellcash*100*100)/100 %> %
        <% End If %>
    </td>
    <td align="center">
        <% If oshoplinker.FItemList(i).IsSoldOut Then %>
            <% If oshoplinker.FItemList(i).FSellyn = "N" Then %>
            <font color="red">ǰ��</font>
            <% Else %>
            <font color="red">�Ͻ�<br>ǰ��</font>
            <% End if %>
        <% End If %>
    </td>
    <td align="center">
    <% If (oshoplinker.FItemList(i).FshoplinkerStatCd > 0) then %>
    <% If Not IsNULL(oshoplinker.FItemList(i).FshoplinkerPrice) Then %>
        <% If (oshoplinker.FItemList(i).Fsellcash <> oshoplinker.FItemList(i).FshoplinkerPrice) Then %>
        <strong><%= formatNumber(oshoplinker.FItemList(i).FshoplinkerPrice,0) %></strong>
        <% Else %>
        <%= formatNumber(oshoplinker.FItemList(i).FshoplinkerPrice,0) %>
        <% End If %>
        <br>
        <% If (oshoplinker.FItemList(i).FshoplinkerSellYn = "X" or oshoplinker.FItemList(i).FshoplinkerSellYn = "N") Then %><a href="javascript:checkNdelReged('<%=oshoplinker.FItemList(i).FItemID%>');"><% End If %>
        <% If (oshoplinker.FItemList(i).FSellyn<>oshoplinker.FItemList(i).FshoplinkerSellYn) Then %>
        <strong><%= oshoplinker.FItemList(i).FshoplinkerSellYn %></strong>
        <% Else %>
        <%= oshoplinker.FItemList(i).FshoplinkerSellYn %>
        <% End If %>
        <% If (oshoplinker.FItemList(i).FshoplinkerSellYn = "X" or oshoplinker.FItemList(i).FshoplinkerSellYn="N") Then %></a><% End If %>
    <% End if %>
    <% End if %>
    </td>
    <td align="center">
    <%
    	If Not(IsNULL(oshoplinker.FItemList(i).FshoplinkerGoodNo)) then
        	Response.Write oshoplinker.FItemList(i).FshoplinkerGoodNo
		End If
	%>
    </td>
    <td align="center"><%= oshoplinker.FItemList(i).FReguserid %></td>
    <td align="center"><%= oshoplinker.FItemList(i).FoptionCnt %></td>
    <td align="center">
    	<%= oshoplinker.FItemList(i).FrctSellCNT %>
	    <% if (oshoplinker.FItemList(i).FaccFailCNT>0) then %>
	        <br><font color="red" title="<%= oshoplinker.FItemList(i).FlastErrStr %>">ERR:<%= oshoplinker.FItemList(i).FaccFailCNT %></font>
	    <% end if %>
   	</td>
    <td align="center"><%= oshoplinker.FItemList(i).FinfoDiv %>
    <% If (oshoplinker.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oshoplinker.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oshoplinker.FItemList(i).FoptAddPrcRegType <> 0,"gray","red")%>">�ɼǱݾ�</font>
	    <% If oshoplinker.FItemList(i).FoptAddPrcRegType <> 0 Then %>
	    (<%=oshoplinker.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% End If %>
    </td>
    <td align="center"><%= oshoplinker.FItemList(i).FInsert_infoCD %></td>
    <td align="center">
		<% if oshoplinker.FItemList(i).FShoplinkerOutMallConnect = "Y" then %>
			<div name="div<%=i%>" id="div<%=i%>">
				<img src="/images/icon_search.jpg" onclick="javascript:divch('div<%=i%>','<%=oshoplinker.FItemList(i).FItemID%>');" style="cursor:pointer;">
			</div>
		<%
		   Else
				response.write oshoplinker.FItemList(i).FShoplinkerOutMallConnect
    	   End If
    	%>
    </td>
</tr>
<%  Next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% If oshoplinker.HasPreScroll then %>
		<a href="javascript:goPage('<%= oshoplinker.StartScrollPage-1 %>');">[pre]</a>
    	<% Else %>
    		[pre]
    	<% End If %>

    	<% For i = 0 + oshoplinker.StartScrollPage to oshoplinker.FScrollCount + oshoplinker.StartScrollPage - 1 %>
    		<% If i>oshoplinker.FTotalpage Then Exit For %>
    		<% If CStr(page) = CStr(i) Then %>
    		<font color="red">[<%= i %>]</font>
    		<% Else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% End If %>
    	<% Next %>

    	<% If oshoplinker.HasNextScroll Then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% Else %>
    		[next]
    	<% End If %>
    </td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF" height="50" align="center">
    <td colspan="17">��ǰ�� �����ϴ�.</td>
</tr>
<% End If %>
</form>
</table>
<% Set oshoplinker = nothing %>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="600"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->