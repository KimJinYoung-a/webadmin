<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �¶��� �ؿ��ǸŴ���ǰ
' History : 2012.11.01 ���ر� ����
'			2013.05.06 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<%
dim oitem, page, i, vKaffaUseYN, vItemID
Dim makerid, kaffaitemid, itemname, eventid, ExtNotReg, bestOrd, bestOrdMall, sellyn, limityn, TenSailyn, KaffaBaseSailyn,KaffaSailyn
Dim onlyValidMargin, failCntExists, optAddprcExists, optAddprcExistsExcept, optExists, KAFFASell10x10Soldout, expensive10x10, mwdiv
Dim diffPrc, diffMultiPrc, reqExpire, extsellyn, extdispyn, kaffaurl, diffWeight
page					= request("page")
vKaffaUseYN				= request("kaffauseyn")
vItemID					= request("itemid")
makerid					= request("makerid")
itemname				= request("itemname")
kaffaitemid				= request("kaffaitemid")
eventid					= request("eventid")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
sellyn					= request("sellyn")
limityn					= request("limityn")
TenSailyn               = request("TenSailyn")
KaffaBaseSailyn         = request("KaffaBaseSailyn")
KaffaSailyn             = request("KaffaSailyn")
onlyValidMargin			= request("onlyValidMargin")
FailCntExists			= request("FailCntExists")
optAddprcExists 		= request("optAddprcExists")
optAddprcExistsExcept	= request("optAddprcExistsExcept")
optExists				= request("optExists")
KAFFASell10x10Soldout	= request("KAFFASell10x10Soldout")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
diffMultiPrc            = request("diffMultiPrc")
diffWeight				= request("diffWeight")
reqExpire				= request("reqExpire")
extsellyn				= request("extsellyn")
extdispyn               = request("extdispyn")
mwdiv                   = request("mwdiv")
If (page = "") then page = 1
If sellyn="" Then sellyn = "Y"

IF application("Svr_Info") = "Dev" THEN
	kaffaurl = "http://test.kaffa.com"
Else
	kaffaurl = "http://10x10shop.com"
End If

Set oitem = new cKaffaItem
	oitem.FPageSize						= 10
	oitem.FCurrPage         			= page
	oitem.FRectMakerid					= makerid
	oitem.FRectItemName					= itemname
	oitem.FRectCDL 						= request("cdl")
	oitem.FRectCDM 						= request("cdm")
	oitem.FRectCDS 						= request("cds")
	oitem.FRectKAFFAPrdNo				= kaffaitemid
	oitem.FRectEventid					= eventid
	oitem.FRectKaffaUseYN				= vKaffaUseYN
	oitem.FRectItemID					= vItemID
	oitem.FRectSellYn					= sellyn
	oitem.FRectLimitYn					= limityn
	oitem.FRectTenSailyn                = TenSailyn
	oitem.FRectKaffaBaseSailyn          = KaffaBaseSailyn
	oitem.FRectKaffaSailyn              = KaffaSailyn
	oitem.FRectonlyValidMargin 			= onlyValidMargin
	oitem.FRectFailCntExists			= failCntExists
	oitem.FRectoptAddprcExists			= optAddprcExists
	oitem.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	oitem.FRectoptExists				= optExists
	oitem.FRectKAFFASell10x10Soldout   	= KAFFASell10x10Soldout
	oitem.FRectExpensive10x10          	= expensive10x10
	oitem.FRectdiffPrc 					= diffPrc
	oitem.FRectdiffMultiPrc             = diffMultiPrc
	oitem.FRectdiffWeight				= diffWeight
	oitem.FRectExtSellYn  				= extsellyn
	oitem.FRectExtDispYn  				= extdispyn
	oitem.FRectMWDiv                    = mwdiv

	If (bestOrd="on") Then
	    oitem.FRectOrdType = "B"
	ElseIf (bestOrdMall="on") Then
	    oitem.FRectOrdType = "BM"
	End If

	If reqExpire <> "" Then
	    oitem.getKaffaReqExpireItemList
	Else
	    oitem.GetAllKaffaItemList_USESCM
    end if
%>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

// ������� �귣��
function NotInMakerid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=kaffa','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=kaffa','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ���õ� ��ǰ �ǸŻ��� Ȯ��
function checkkaffaItemConfirm(){
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

        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CheckItemStat";
        document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
        document.frmSvArr.submit();
}

//�ǸŻ���Ȯ�� ��ġ
function batchStatCheck(){
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "CheckItemStatAuto";
    document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
    document.frmSvArr.submit();
}

// ���õ� ��ǰ �Ǹſ��� ����
function kaffamallSellYnProcess(chkYn) {
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
		case "N": strSell="�Ͻ��ߴ�";break;
		case "X": strSell="�Ǹ�����(����)";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n����Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� ����mall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

        //document.getElementById("btnSellYn").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "product_sale";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
		document.frmSvArr.submit();
    }
}
function kaffaSelectEditProcess(){
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n����Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "set_product";
        document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
        document.frmSvArr.submit();
    }
}

function kaffaSelectSaleStatEditProcess(){
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

	if (confirm('����all�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n������Mall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditDanpum").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "stock_fix";
		document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
		document.frmSvArr.submit();
	}
}

function kaffaSelectDateEdit2Process(){
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

    if (confirm('����all�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n������Mall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditDate").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "productstock";
        document.frmSvArr.action = "/admin/etc/kaffa/actKaffaReq.asp"
        document.frmSvArr.submit();
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

    if ((comp.name=="KAFFASell10x10Soldout")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }

        //comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "A";
        comp.form.limityn.value = "";
        //comp.form.infodiv.value = "";
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.KAFFASell10x10Soldout.checked){
            comp.form.KAFFASell10x10Soldout.checked = false;
        }

//		comp.form.ExtNotReg.value = "D";
        comp.form.sellyn.value = "Y";
        comp.form.limityn.value = "";
        comp.form.onlyValidMargin.checked=false;
    }

	if ((comp.name=="reqExpire")&&(comp.checked)){
//	    comp.form.ExtNotReg.value="D";
	    comp.form.sellyn.value="A";
	    comp.form.limityn.value="";
	    comp.form.onlyValidMargin.checked=false;
	}
	if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
	if ((comp.name=="diffPrc")){frm.onlyValidMargin.checked=true;}
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td class="a">
		�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		kaffa��ǰ��ȣ: <input type="text" name="kaffaitemid" value="<%= kaffaitemid %>" size="9" maxlength="9" class="input">
		��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		10x10��ǰ��ȣ : <input type="text" name="itemid" value="<%=vItemID%>" size="40">&nbsp;
   		�̺�Ʈ��ȣ: <input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
		<br>
    	��Ͽ��� :
		<select name="kaffauseyn" class="select">
			<option value="">kaffa��Ͽ���</option>
			<option value="y" <%=CHKIIF(vKaffaUseYN="y","selected","")%>>��ϵ�</option>
			<option value="n" <%=CHKIIF(vKaffaUseYN="n","selected","")%>>�̵��</option>
			<option value="m" <%=CHKIIF(vKaffaUseYN="m","selected","")%>>���������ǰ</option>
			<option value="w" <%=CHKIIF(vKaffaUseYN="w","selected","")%>>���δ��</option>
		</select>&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>&nbsp;
		�Ǹſ��� :
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
		</select>&nbsp;
		�������� :
		<select name="limityn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
		TEN ���ο��� :
		<select name="TenSailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(TenSailyn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(TenSailyn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
		�ؿܱ��� ���ο��� :
		<select name="KaffaBaseSailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(KaffaBaseSailyn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(KaffaBaseSailyn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
        Kaffa ���ο��� :
		<select name="KaffaSailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(KaffaSailyn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(KaffaSailyn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
	    �ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	    &nbsp;
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� <%= CMAXMARGIN %>%�̻� ��ǰ�� ����
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >����������ǰ
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ
		&nbsp;
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ����
		&nbsp;
		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >�ɼ������ǰ
		<br>
		<input type="checkbox" name="KAFFASell10x10Soldout" <%= ChkIIF(KAFFASell10x10Soldout="on","checked","") %> onClick="checkComp(this)"><font color=red>KAFFA�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> onClick="checkComp(this)"><font color=red>KAFFA ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkComp(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>(on�ǸŰ�<>kaffa)��ü����
		&nbsp;
		<input onClick="checkComp(this)" type="checkbox" name="diffMultiPrc" <%= ChkIIF(diffMultiPrc="on","checked","") %> ><font color=red>���ݻ���</font>(�ؿܱ��ذ�<>kaffa)��ü����
		&nbsp;
		<input onClick="checkComp(this)" type="checkbox" name="diffWeight" <%= ChkIIF(diffWeight="on","checked","") %> ><font color=red>���Ի���</font>��ü����
		<br>
		<input onClick="checkComp(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (��ü���, ��������100g�̻�, ����0)
		&nbsp;&nbsp;Kaffa�ǸŻ��� :
		<select name="extsellyn" class="select">
		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
		</select>

		&nbsp;&nbsp;Kaffa���û��� :
		<select name="extdispyn" class="select">
		<option value="" <%= CHkIIF(extdispyn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(extdispyn="Y","selected","") %> >����
		<option value="N" <%= CHkIIF(extdispyn="N","selected","") %> >���þ���
		</select>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<br>
<form name="frmReg" method="post" action="index.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
			<!--
				<input class="button" type="button" value="��� ���� �귣��" onclick="NotInMakerid();"> &nbsp;
			-->
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">
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
	    		������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="���û�ǰ ����/���� ����" onClick="kaffaSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDanpum" value="���û�ǰ��ǰ/�ǸŻ��� ����" onClick="kaffaSelectSaleStatEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDate" value="���û�ǰ ����+��ǰ ����" onClick="kaffaSelectDateEdit2Process();">
				<br><br>
				���ο��� �˻� :
				<input class="button" type="button"  value="���û�ǰ (�ǸŻ���) Ȯ��" onClick="checkkaffaItemConfirm(this);" >
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ͻ��ߴ�</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="kaffamallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input type="button" value="�ǸŻ���Check(������)" onClick="batchStatCheck();">
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
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="#FFFFFF">
	<td colspan="23" align="right"><strong>Total : <%=oitem.FTotalCount%></strong></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width=50> �̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td align="center">�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">kaffa�����<br>kaffa����������</td>
	<td width="60">On<br>�ǸŰ�</td>
	<td width="60">On<br>���԰�</td>
	<td width="50">On<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ؿ�<br>���ذ�</td>
	<td width="40">���</td>
	<td width="70">kaffa<br>����</td>
	<td width="50">kaffa<br>�Ǹ�</td>
	<td width="50">kaffa<br>����</td>
	<td width="70">kaffa<br>��ǰ��ȣ</td>
	<td width="100">�����</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="30">���<br>����</td>
	<td width="30">����<br>����</td>
	<td width="40">�ؿ�<br>����</td>
	<td width="60">��ǰ<br>����</td>
</tr>
<%
If oitem.FresultCount < 1 Then
%>
<tr bgcolor="#FFFFFF">
	<td colspan="23" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
End If

If oitem.FresultCount > 0 Then
	For i = 0 to oitem.FresultCount - 1
%>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oitem.FItemList(i).Fitemid %>"></td>
	<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
	<td align="center">
		<a href="<%=wwwURL%>/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="�̸�����">
		<%= oitem.FItemList(i).Fitemid %></a>
		<%= oitem.FItemList(i).getLimitDispStr %>
	</td>
	<td><%= oitem.FItemList(i).FMakerid %><%= oitem.FItemList(i).getDeliverytypeName %><br><%= oitem.FItemList(i).FItemName %></td>
	<td align="center"><%= oitem.FItemList(i).FRegdate %><br><%= oitem.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oitem.FItemList(i).FKaffaregdate %><br><%= oitem.FItemList(i).FKaffalastupdate %></td>
	<td align="right">
	<% If oitem.FItemList(i).Fsailyn="Y" Then %>
		<strike><%= FormatNumber(oitem.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oitem.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oitem.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
	<td align="center">
	<%
		If oitem.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oitem.FItemList(i).Fbuycash/oitem.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oitem.FItemList(i).IsSoldOut Then
			If oitem.FItemList(i).FSellyn = "N" Then
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
	    <% if oitem.FItemList(i).FmaydiscountPrice<>0 then %>
	    <strike><%= formatNumber(oitem.FItemList(i).FmultiPrice,0)%></strike>
	    <br><font color="red"><%= formatNumber(oitem.FItemList(i).FmaydiscountPrice,0)%></font>
	    <% else %>
	    <%= formatNumber(oitem.FItemList(i).FmultiPrice,0)%>
	    <% end if %>
	</td>
	<td align="center">
	<%= oitem.FItemList(i).getForeignMultipleStr() %>
	</td>
	<td align="center">
	<%
		If Not IsNULL(oitem.FItemList(i).FKaffaprice) Then
			If (oitem.FItemList(i).FmultiPrice <> oitem.FItemList(i).FKaffaprice) Then

			    if oitem.FItemList(i).IsKaffaSiteDiscountSale then
			        response.write "<strike><strong>"&formatNumber(oitem.FItemList(i).FKaffaprice,0)&"</strong></strike><br><font color=red title='"&oitem.FItemList(i).getDiscountDateStr&"'>"&formatNumber(oitem.FItemList(i).FKaffaDiscountPrice,0)&"</font>"
			    else
			        response.write "<strong>"&formatNumber(oitem.FItemList(i).FKaffaprice,0)&"</strong>"
			    end if
			Else
			    if oitem.FItemList(i).IsKaffaSiteDiscountSale then
				    response.write "<strike>"&formatNumber(oitem.FItemList(i).FKaffaprice,0)&"</strike><br><font color=red title='"&oitem.FItemList(i).getDiscountDateStr&"'>"&formatNumber(oitem.FItemList(i).FKaffaDiscountPrice,0)&"</font>"
				else
				    response.write formatNumber(oitem.FItemList(i).FKaffaprice,0)
			    end if
			End If
		End If
	%>
	</td>
	<td align="center">
	<% If (oitem.FItemList(i).FSellyn="Y" and oitem.FItemList(i).FKaffasellyn<>"Y") or (oitem.FItemList(i).FSellyn<>"Y" and oitem.FItemList(i).FKaffasellyn="Y") Then %>
	    <strong><%= oitem.FItemList(i).FKaffasellyn %></strong>
	<% else %>
	    <%= oitem.FItemList(i).FKaffasellyn %>
	<% end if %>
	</td>
	<td align="center"><%= CHKIIF(oitem.FItemList(i).FkaffaIsDisplay=1,"Y","<font color='red'>N</font>") %></td>
	<td align="center">
	<%
		If Not(IsNULL(oitem.FItemList(i).FKaffagoodno)) Then
			Response.Write "<a target='_blank' href='"&kaffaurl & "/front/product?product_id="&oitem.FItemList(i).FKaffagoodno&"'>"&oitem.FItemList(i).FKaffagoodno&"</a>"
		End If
	%>
	</td>
	<td align="center"><%=oitem.FItemList(i).FRegUser%></td>
	<td align="center"><%= oitem.FItemList(i).FoptionCnt %>:<%= oitem.FItemList(i).FregedOptCnt %></td>
	<td align="center"><%= oitem.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
	<td align="center"><%= fnColor(oitem.FItemList(i).FdeliverOverseas,"yn") %></td>
	<td align="center"><%= FormatNumber(oitem.FItemList(i).FitemWeight,0) %>g
	<% if oitem.FItemList(i).FitemWeight<>oitem.FItemList(i).FkaffaWeight then %>
	    <br><font color="red"><%=oitem.FItemList(i).FkaffaWeight%>g</font>
	<% end if %>
	<%
		If (oitem.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& oitem.FItemList(i).FlastErrStr &"'>ERR:"& oitem.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="23" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% set oitem = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->