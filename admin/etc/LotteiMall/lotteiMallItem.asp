<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->

<%
dim itemid, itemname, eventid, mode
dim page, makerid, LotteNotReg, MatchCate, sellyn, limityn
dim delitemid, lotteGoodNo, showminusmagin, expensive10x10, LotteYes10x10No, LotteNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists, regedOptNull,failCntExists, optAddPrcRegTypeNone
Dim bestOrd, bestOrdMall, extsellyn, infoDiv
dim i
page    = request("page")
itemid  = request("itemid")
If itemid <> "" Then
	If Right(itemid,1) = "," OR Right(itemid,1) = " " Then
		Response.Write "<script>alert('��ǰ�ڵ尡 �߸� �ԷµǾ����ϴ�.');history.back();</script>"
		Response.End
	End IF
End IF


eventid  = request("eventid")
itemname = html2db(request("itemname"))
mode     = request("mode")
makerid  = request("makerid")
LotteNotReg = request("LotteNotReg")
MatchCate = request("MatchCate")
sellyn    = request("sellyn")
limityn   = request("limityn")
delitemid = requestCheckvar(request("delitemid"),9)
lotteGoodNo = requestCheckvar(request("lotteGoodNo"),9)
showminusmagin = request("showminusmagin")
expensive10x10 = request("expensive10x10")
LotteYes10x10No = request("LotteYes10x10No")
LotteNo10x10Yes = request("LotteNo10x10Yes")
onreginotmapping = request("onreginotmapping")
diffPrc = request("diffPrc")
onlyValidMargin = request("onlyValidMargin")
research = request("research")
reqExpire = request("reqExpire")
reqEdit   = request("reqEdit")
optAddprcExists = request("optAddprcExists")
optAddPrcRegTypeNone  = request("optAddPrcRegTypeNone")
optAddprcExistsExcept = request("optAddprcExistsExcept")
optExists   = request("optExists")
regedOptNull= request("regedOptNull")
bestOrd  = request("bestOrd")
bestOrdMall= request("bestOrdMall")
failCntExists= request("failCntExists")
extsellyn   = request("extsellyn")
infoDiv     = request("infoDiv")

if page="" then page=1
if sellyn="" then sellyn="Y"

''�⺻���� ��Ͽ����̻�
if (research="") then
    LotteNotReg = "J"
    MatchCate = "" ''Y
    onlyValidMargin="on"    ''2013/05/23����
    ''bestOrd="on"
    sellyn="Y"              ''2013/05/23����
end if

dim oiMall
set oiMall = new CLotteiMall
if (LotteNotReg="F") then                       '''���δ��
oiMall.FPageSize       = 50
else
oiMall.FPageSize       = 20
end if
oiMall.FCurrPage       = page
oiMall.FRectItemID     = itemid
oiMall.FRectEventid    = eventid
oiMall.FRectItemName   = itemname
oiMall.FRectMakerid    = makerid
oiMall.FRectCDL = request("cdl")
oiMall.FRectCDM = request("cdm")
oiMall.FRectCDS = request("cds")
oiMall.FRectLotteNotReg  = LotteNotReg
oiMall.FRectMatchCate  = MatchCate
oiMall.FRectSellYn  = sellyn
oiMall.FRectLimitYn  = limityn
oiMall.FRectLTiMallGoodNo  = lotteGoodNo
oiMall.FRectMinusMigin = showminusmagin
oiMall.FRectExpensive10x10 = expensive10x10
oiMall.FRectLotteYes10x10No = LotteYes10x10No
oiMall.FRectLotteNo10x10Yes = LotteNo10x10Yes
oiMall.FRectOnreginotmapping = onreginotmapping
oiMall.FRectdiffPrc = diffPrc
oiMall.FRectonlyValidMargin = onlyValidMargin
oiMall.FRectoptAddprcExists= optAddprcExists
oiMall.FRectoptAddPrcRegTypeNone = optAddPrcRegTypeNone                         ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
oiMall.FRectoptAddprcExistsExcept= optAddprcExistsExcept
oiMall.FRectoptExists= optExists
oiMall.FRectregedOptNull= regedOptNull
oiMall.FRectFailCntExists = failCntExists
oiMall.FRectFailCntOverExcept = ""
oiMall.FRectExtSellYn  = extsellyn
oiMall.FRectInfoDiv = infoDiv

IF (bestOrd="on") then
    oiMall.FRectOrdType = "B"
ELSEIF (bestOrdMall="on") then
    oiMall.FRectOrdType = "BM"
end if

IF reqExpire<>"" then
    oiMall.getLottereqExpireItemList
ELSE
    oiMall.getLTiMallRegedItemList
ENd IF
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function popManageOptAddPrc(iitemid){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteimall',"popOptionAddPrc","width=600,height=300,scrollbars=yes,resizable=yes");
	pwin.focus();
}

// �Ե�iMall ���MD ���
function pop_MDList() {
	var pMD = window.open("/admin/etc/LotteiMall/popLTiMallMDList.asp","popMDListIMall","width=600,height=300,scrollbars=yes,resizable=yes");
	pMD.focus();
}

// �Ե�iMall ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/LotteiMall/popLTiMallCateList.asp","popCateManIMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}

// �Ե�iMall �귣�� ����
function pop_BrandList() {
	alert("�귣��� �ٹ�����()���� �����˴ϴ�.");
	//var pBM = window.open("/admin/etc/LotteiMall/popLotteBrandMap.asp","popBrandManIMall","width=500,height=500,scrollbars=yes,resizable=yes");
	//pBM.focus();
}

// �̵�� ��ǰ �ϰ����
function LotteRegProcess(){
    if (confirm('�Ե�iMall�� �̵�ϵ� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnRegAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegAll";
        document.frmSvArr.action = "/admin/etc/LotteiMall/actRegLtiMallItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function LotteSelectRegProcess(isreal) {
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

    if (isreal){
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelect";
            document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?\n\n��30�д����� ��ġ ��ϵ˴ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }
}

// ���õ� ��ǰ �ϰ� ���
function LotteregIMSI(isreg) {
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

    if (isreg){
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?\n\n��30�д����� ��ġ ��ϵ˴ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� ���� �Ͻðڽ��ϱ�?')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "DelSelectWait";
            document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }
}

// ������ ��ǰ �ϰ�����
function LotteEditProcess(){
    if (confirm('������ ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "/admin/etc/LotteiMall/actRegLtiMallItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function LotteSelectEditProcess() {
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

    if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ��ǰ ����
function LotteSelectSaleStatEditProcess() {
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

    if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditDanpum").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EdSaleDTSel";
        document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}


// ���޸� �ƴѰ� ����
function LotteDelJaeHyuProcess(){
    //return;
    if (confirm('���޸��ǸŰ� �ƴѰ��� �ϰ� ���� �Ͻðڽ��ϱ�?')){
        document.getElementById("btnDelJehyu").disabled=true;

        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "DelJaeHyu";

        document.frmSvArr.action = "/admin/etc/LotteiMall/actRegLtiMallItem.asp"
        document.frmSvArr.submit();
    }

}

// ������� �귣��
function NotInMakerid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=lotteiMall','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteiMall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ӽû�ǰ �ϰ� ����
function LotteRealItemMapping() {
	if(confirm("�ӽõ�� ��ǰ�� ���û�ǰ���� ��ϵǾ����� �ϰ� Ȯ���Ͻðڽ��ϱ�?\n\n�� ��Ż��¿����� �ټ� �ð��� �ɸ� �� �ֽ��ϴ�.")) {
		document.getElementById("btnChkReal").disabled=true;
		xLink.location.href="/admin/etc/LotteiMall/actLotteiMallReq.asp?cmdparam=getconfirmList"
	}
}

// ���õ� ��ǰ �ϰ� ����
function LotteRealItemMappingChecked(chkYn) {
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


    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "getconfirmList";
    document.frmSvArr.subcmd.value = "arrconfirm";
    document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
    document.frmSvArr.submit();
}


// ���õ� ��ǰ �Ǹſ��� ����
function LTiMallSellYnProcess(chkYn) {
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
		case "X": strSell="�Ǹ�����(����)";break;
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� �Ե�iMall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

        //document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.subcmd.value = chkYn;
        document.frmSvArr.action = "/admin/etc/LotteiMall/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

function checkQuickClick(comp){
    var frm = comp.form;

    if (comp.checked){
        if (comp.name=="reqREG") {
            frm.LotteNotReg.value="M";
            frm.MatchCate.value="Y";
            frm.sellyn.value="Y";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="mapRealPrdCode"){
            frm.LotteNotReg.value="F";
            frm.MatchCate.value="Y";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }else if (comp.name=="reqEdit"){
            frm.LotteNotReg.value="R";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else{
            frm.LotteNotReg.value="D";
            frm.MatchCate.value="";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=false;
        }

        if ((comp.name=="LotteNo10x10Yes")||(comp.name=="diffPrc")){
            frm.onlyValidMargin.checked=true;
        }

        if ((comp.name!="showminusmagin")&&(frm.showminusmagin.checked)){ frm.showminusmagin.checked=false }
        if ((comp.name!="expensive10x10")&&(frm.expensive10x10.checked)){ frm.expensive10x10.checked=false }
        if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
        if ((comp.name!="LotteYes10x10No")&&(frm.LotteYes10x10No.checked)){ frm.LotteYes10x10No.checked=false }
        if ((comp.name!="LotteNo10x10Yes")&&(frm.LotteNo10x10Yes.checked)){ frm.LotteNo10x10Yes.checked=false }
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

function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//��ǰ�ǸŻ���,����Check
function batchStatCheck(){
    xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckItemStatAuto";
}

function checkNdel(iitemid){
    xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckNDel&cksel="+iitemid;
}

function checkNdelReged(iitemid){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        xLink.location.href="actLotteiMallReq.asp?cmdparam=CheckNDelReged&cksel="+iitemid;
    }
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		�� �� �� :
		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		�Ե�iMall��ǰ��ȣ:
		<input type="text" name="lotteGoodNo" value="<%= lotteGoodNo %>" size="9" maxlength="9" class="text"> &nbsp;&nbsp;
		��ǰ��:
		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ��ȣ:
		<input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		�̺�Ʈ��ȣ:
		<input type="text" name="eventid" value="<%= eventid %>" size="8" maxlength="6" class="text">
		<br>
		��Ͽ��� :
		<select name="LotteNotReg" class="select">
		<option value="">��ü
		<option value="M" <%= CHkIIF(LotteNotReg="M","selected","") %> >�Ե�iMall �̵��(��ϰ���)
		<option value="Q" <%= CHkIIF(LotteNotReg="Q","selected","") %> >�Ե�iMall ��Ͻ���
		<option value="J" <%= CHkIIF(LotteNotReg="J","selected","") %> >�Ե�iMall ��Ͽ����̻�
		<option value="W" <%= CHkIIF(LotteNotReg="W","selected","") %> >�Ե�iMall ��Ͽ���
		<option value="V" <%= CHkIIF(LotteNotReg="V","selected","") %> >�Ե�iMall ��Ͽ���/��ϰ���
		<option value="A" <%= CHkIIF(LotteNotReg="A","selected","") %> >�Ե�iMall ���۽õ��߿���
		<option value="F" <%= CHkIIF(LotteNotReg="F","selected","") %> >�Ե�iMall ����� ���δ��(�ӽ�)
		<option value="D" <%= CHkIIF(LotteNotReg="D","selected","") %> >�Ե�iMall ��ϿϷ�(����)
		<option value="R" <%= CHkIIF(LotteNotReg="R","selected","") %> >�Ե�iMall �������
		</select>
		&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>
		&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>
		&nbsp;
		ī�׸�Ī :
		<select name="MatchCate" class="select">
		<option value="">��ü
		<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
		<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
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
		<input type="checkbox" name="regedOptNull" <%= ChkIIF(regedOptNull="on","checked","") %> >��ǰ��� �̼���
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >��ϼ���������ǰ
		<br><br>
		-- Quick �˻� / ��� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >��ϰ��� ��ǰ
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="mapRealPrdCode"  >�ӽû�ǰ �ϰ�Ȯ��
		<br><br>
		-- Quick �˻� / ���� / --
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>������</font>��ǰ���� (MaxMagin : <%= CMAXMARGIN %>%) (�Ե� �Ǹ���)
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>�Ե�iMall ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteYes10x10No" <%= ChkIIF(LotteYes10x10No="on","checked","") %> ><font color=red>�Ե�iMall�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteNo10x10Yes" <%= ChkIIF(LotteNo10x10Yes="on","checked","") %> ><font color=red>�Ե�iMallǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (���޸� �����Ե�)
		&nbsp;&nbsp;�����ǸŻ��� :
		<select name="extsellyn" class="select">
		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
		</select>

		&nbsp;&nbsp;ǰ������ :
		<% CALL DrawItemInfoDiv("infoDiv", infoDiv, true, "") %>

		<!--
		&nbsp;&nbsp;ǰ�������Է¿��� :
		<select name="infoDivYn" class="select">
		<option value="" <%= CHkIIF(infoDivYn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(infoDivYn="Y","selected","") %> >�Է�
		<option value="N" <%= CHkIIF(infoDivYn="N","selected","") %> >���Է�
		<option value="21" <%= CHkIIF(infoDivYn="21","selected","") %> >21
		<option value="23" <%= CHkIIF(infoDivYn="23","selected","") %> >23
		<option value="35" <%= CHkIIF(infoDivYn="35","selected","") %> >35
		</select>
		-->
		<!--
		<input type="checkbox" name="onreginotmapping" <%= ChkIIF(onreginotmapping="on","checked","") %> ><font color=red>�Ե�iMall ���&����Ī ī�װ�</font>��ǰ����
		&nbsp;
		-->
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<br>
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
			<!--
				<input class="button" type="button" value="Lotte���MD" onclick="pop_MDList();"> &nbsp;
				<input class="button" type="button" value="Lotte�귣���Ī" onclick="pop_BrandList();"> &nbsp;
			-->
				<input class="button" type="button" value="LotteiMallī�װ���Ī" onclick="pop_CateManager();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal" value="���û�ǰ [����] Ȯ��" onClick="LotteRealItemMappingChecked();" <%= CHKIIF((TRUE) or (LotteNotReg="F"),"","disabled") %> >

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
	    	    <!--
			    <input class="button" type="button" id="btnEditAll" value="������ ��ǰ �Ե�iMall���� �ϰ����� [10�Ǿ�]" onClick="LotteEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegAll" value="�̵�� ��ǰ �Ե�iMall���� �ϰ���� [10�Ǿ�]" onClick="LotteRegProcess();">
			    -->
			    <br><br>
			    <input class="button" type="button" id="btnEditSel" value="���û�ǰ����/���� ����" onClick="LotteSelectEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditDanpum" value="���û�ǰ��ǰ/�ǸŻ��� ����" onClick="LotteSelectSaleStatEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="���û�ǰ �� ���" onClick="LotteSelectRegProcess(true);">
			    <br><br>
			    <input class="button" type="button" id="btnRegSel" value="���û�ǰ ���� ���" onClick="LotteregIMSI(true);">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="���û�ǰ ���� ����" onClick="LotteregIMSI(false);" >
			    &nbsp;&nbsp;
			    <!--
			    <input class="button" type="button" id="btnDelJehyu" value="���޸��ƴѰ� �ϰ����� [20�Ǿ�]" onClick="LotteDelJaeHyuProcess();">
			    -->

			</td>
			<td align="right" valign="top">
			    <!--
			    &nbsp;
				<input class="button" type="button" id="btnChkReal" value="�ӽû�ǰ �ϰ�Ȯ�� [200�Ǿ�]" onClick="LotteRealItemMapping();">
				-->
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
				<!-- <option value="Y">�Ǹ���</option> -->
				<option value="N">ǰ��</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="LTiMallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input type="button" value="�ǸŻ���Check(������)" onClick="batchStatCheck();">
				<!-- &nbsp;&nbsp;<input type="button" value="��ǰ�����(������)" onClick="batchItemNmCheck();"> -->

			</td>
		</tr>
		</table>
    </td>
</tr>
<tr>
    <td>
    ����ó����ǰ(���� ����) : 210499,724724,692489
    </td>
</tr>
</table>
</form>
<br>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="subcmd" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oiMall.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oiMall.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td >�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">�Ե�iMall�����<br>�Ե�iMall����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�Ե�iMall<br>���ݹ��Ǹ�</td>
	<td width="70">�Ե�iMall<br>��ǰ��ȣ<br>(�ӽù�ȣ)</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% for i=0 to oiMall.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oiMall.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oiMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oiMall.FItemList(i).FItemID %>','lotteimall','')" style="cursor:pointer"></td>
    <td align="center"><%= oiMall.FItemList(i).FItemID %><br><%= oiMall.FItemList(i).getLtiMallStatName %>
    <% if oiMall.FItemList(i).FLimitYn="Y" then %><br><%= oiMall.FItemList(i).getLimitHtmlStr %></font><% end if %>
    </td>
    <td><%= oiMall.FItemList(i).FMakerid %> <%= oiMall.FItemList(i).getDeliverytypeName %><br><%= oiMall.FItemList(i).FItemName %></td>
    <td align="center"><%= oiMall.FItemList(i).FRegdate %><br><%= oiMall.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oiMall.FItemList(i).FLTiMallRegdate %><br><%= oiMall.FItemList(i).FLTiMallLastUpdate %></td>
    <td align="right">
        <% if oiMall.FItemList(i).FSaleYn="Y" then %>
        <strike><%= FormatNumber(oiMall.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %></font>
        <% else %>
        <%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %>
        <% end if %>
    </td>
    <td align="center">
        <% if oiMall.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oiMall.FItemList(i).IsSoldOut then %>
            <% if oiMall.FItemList(i).FSellyn="N" then %>
            <font color="red">ǰ��</font>
            <% else %>
            <font color="red">�Ͻ�<br>ǰ��</font>
            <% end if %>
        <% end if %>
    </td>
    <td align="center">
    <% if (oiMall.FItemList(i).FLTiMallStatCd>0) then %>
    <% if Not IsNULL(oiMall.FItemList(i).FLTiMallPrice) then %>
        <% if (oiMall.FItemList(i).Fsellcash<>oiMall.FItemList(i).FLTiMallPrice) then %>
        <strong><%= formatNumber(oiMall.FItemList(i).FLTiMallPrice,0) %></strong>
        <% else %>
        <%= formatNumber(oiMall.FItemList(i).FLTiMallPrice,0) %>
        <% end if %>
        <br>
        <% if (oiMall.FItemList(i).FLTiMallSellYn="X" or oiMall.FItemList(i).FLTiMallSellYn="N") then %><a href="javascript:checkNdelReged('<%=oiMall.FItemList(i).FItemID%>');"><% end if %>
        <% if (oiMall.FItemList(i).FSellyn<>oiMall.FItemList(i).FLTiMallSellYn) then %>
        <strong><%= oiMall.FItemList(i).FLTiMallSellYn %></strong>
        <% else %>
        <%= oiMall.FItemList(i).FLTiMallSellYn %>
        <% end if %>
        <% if (oiMall.FItemList(i).FLTiMallSellYn="X" or oiMall.FItemList(i).FLTiMallSellYn="N") then %></a><% end if %>
    <% end if %>
    <% end if %>
    </td>
    <td align="center">
    <%
    	'#�ǻ�ǰ��ȣ
    	if Not(IsNULL(oiMall.FItemList(i).FLtiMallGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotteimall.com/product/Product.jsp?i_code="&oiMall.FItemList(i).FLtiMallGoodNo&"'>"&oiMall.FItemList(i).FLtiMallGoodNo&"</a>"
		else
			'#�ӽû�ǰ��ȣ
			if Not(IsNULL(oiMall.FItemList(i).FLtiMallTmpGoodNo)) then
				if oiMall.FItemList(i).FLTiMallStatCd<>"30" then
					Response.Write oiMall.FItemList(i).getLotteItemStatCd & "<br>(" & oiMall.FItemList(i).FLtiMallTmpGoodNo & ")"
				end if
			else
				Response.Write "<a href='javascript:checkNdel("&oiMall.FItemList(i).FItemID&");'><img src='/images/i_delete.gif' width='8' height='9' border='0'></a>"
			end if
		end if
	%>
    </td>
    <td align="center"><%= oiMall.FItemList(i).Freguserid %></td>
    <td align="center"><%= oiMall.FItemList(i).FoptionCnt %>:<%= oiMall.FItemList(i).FregedOptCnt %></td>
    <td align="center"><%= oiMall.FItemList(i).FrctSellCNT %></td>
    <td align="center">
    <% if oiMall.FItemList(i).FCateMapCnt>0 then %>
	    ��Ī��
    <% else %>
    	<font color="darkred">��Ī�ȵ�</font>
    <% end if %>

    <% if (oiMall.FItemList(i).FaccFailCNT>0) then %>
        <br><font color="red" title="<%= oiMall.FItemList(i).FlastErrStr %>">ERR:<%= oiMall.FItemList(i).FaccFailCNT %></font>
    <% end if %>
    </td>
    <td align="center"><%= oiMall.FItemList(i).FinfoDiv %>
    <% if (oiMall.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oiMall.FItemList(i).FItemID%>');"><font color="red">�ɼǱݾ�</font>(<%=oiMall.FItemList(i).FoptAddPrcRegType%>)</a>
    <% end if %>
    </td>

</tr>
<% next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if oiMall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oiMall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oiMall.StartScrollPage to oiMall.FScrollCount + oiMall.StartScrollPage - 1 %>
    		<% if i>oiMall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oiMall.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<form name="frmDel" method="post" action="Lotteitem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
<% set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
