<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
dim itemid, itemname, eventid, mode
dim page, makerid, LotteNotReg, MatchCate, sellyn, limityn, sailyn
dim delitemid, lotteGoodNo, showminusmagin, expensive10x10, LotteYes10x10No, LotteNo10x10Yes, onreginotmapping, diffPrc, onlyValidMargin, research, lotteTmpGoodNo
Dim reqExpire, reqEdit, optAddprcExists,optAddprcExistsExcept,optExists, regedOptNull,failCntExists, optAddPrcRegTypeNone, optnotExists, isMadeHand
dim bestOrd,bestOrdMall, ckLimitOver, extsellyn, infoDivYn

dim i
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

eventid  = request("eventid")
itemname= html2db(request("itemname"))
mode    = request("mode")
makerid= request("makerid")
LotteNotReg = request("LotteNotReg")
MatchCate = request("MatchCate")
sellyn = request("sellyn")
limityn = request("limityn")
sailyn = request("sailyn")
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
lotteTmpGoodNo	= request("lotteTmpGoodNo")
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
ckLimitOver= request("ckLimitOver")
extsellyn   = request("extsellyn")
infoDivYn   = request("infoDivYn")
optnotExists   			= request("optnotExists")
isMadeHand					= request("isMadeHand")
if page="" then page=1



if sellyn="" then sellyn="Y"
if research="" then
    onlyValidMargin="on"
    LotteNotReg = "D"
end if


dim oLotteitem
set oLotteitem = new CLotte
oLotteitem.FPageSize       = 30
oLotteitem.FCurrPage       = page
oLotteitem.FRectItemID     = itemid
oLotteitem.FRectEventid    = eventid
oLotteitem.FRectItemName   = itemname
oLotteitem.FRectMakerid    = makerid
oLotteitem.FRectCDL = request("cdl")
oLotteitem.FRectCDM = request("cdm")
oLotteitem.FRectCDS = request("cds")
oLotteitem.FRectLotteNotReg  = LotteNotReg
oLotteitem.FRectMatchCate  = MatchCate
oLotteitem.FRectSellYn  = sellyn
oLotteitem.FRectLimitYn  = limityn
oLotteitem.FRectSailYn  = sailyn
oLotteitem.FRectLotteGoodNo  = lotteGoodNo
oLotteitem.FRectMinusMigin = showminusmagin
oLotteitem.FRectExpensive10x10 = expensive10x10
oLotteitem.FRectLotteYes10x10No = LotteYes10x10No
oLotteitem.FRectLotteNo10x10Yes = LotteNo10x10Yes
oLotteitem.FRectOnreginotmapping = onreginotmapping
oLotteitem.FRectdiffPrc = diffPrc
oLotteitem.FRectonlyValidMargin = onlyValidMargin
oLotteitem.FRectoptAddprcExists= optAddprcExists
oLotteitem.FRectLotteTmpGoodNo		= lotteTmpGoodNo
oLotteitem.FRectoptAddPrcRegTypeNone = optAddPrcRegTypeNone                         ''�ɼ��߰��ݾ׻�ǰ �̼��� ��ǰ.
oLotteitem.FRectoptAddprcExistsExcept= optAddprcExistsExcept
oLotteitem.FRectoptExists= optExists
oLotteitem.FRectregedOptNull= regedOptNull
oLotteitem.FRectFailCntExists = failCntExists
oLotteitem.FRectFailCntOverExcept = ""
oLotteitem.FRectExtSellYn  = extsellyn
oLotteitem.FRectInfoDivYn = infoDivYn
oLotteitem.FRectoptnotExists			= optnotExists
oLotteitem.FRectisMadeHand				= isMadeHand
if (ckLimitOver="on") then
    oLotteitem.FRectLimitOver=CStr(CMAXLIMITSELL)
end if

IF (bestOrd="on") then
    oLotteitem.FRectOrdType = "B"
ELSEIF (bestOrdMall="on") then
    oLotteitem.FRectOrdType = "BM"
end if

IF reqExpire<>"" then
    oLotteitem.getLottereqExpireItemList
ELSE
    oLotteitem.GetLotteRegedItemList
ENd IF

Dim outMallItemArr
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteCom&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

// �Ե����� ���MD ���
function pop_MDList() {
	var pMD = window.open("./lotte/popLotteMDList.asp","popMDList","width=600,height=300,scrollbars=yes,resizable=yes");
	pMD.focus();
}

// �Ե����� ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("./lotte/popLotteCateList.asp","popCateMan","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}

// �Ե����� �귣�� ����
function pop_BrandList() {
	alert("�귣��� �ٹ�����(155112)���� �����˴ϴ�.");
	//var pBM = window.open("./lotte/popLotteBrandMap.asp","popBrandMan","width=500,height=500,scrollbars=yes,resizable=yes");
	//pBM.focus();
}

// �̵�� ��ǰ �ϰ����
function LotteRegProcess(){
    if (confirm('�Ե����Ŀ� �̵�ϵ� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnRegAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegAll";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function LotteSelectRegProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "RegSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���û�ǰ �����ȸ
function LotteSelectcheckStock() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� ��ȸ �Ͻðڽ��ϱ�?')){
       // document.getElementById("btnRegSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "ChkStockSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ��Ͽ��� ���
function LotteregIMSI(isreg){
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
        if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?\n\n��30�д����� ��ġ ��ϵ˴ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.mode.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
            document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� ���� �Ͻðڽ��ϱ�?')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.mode.value = "DelSelectWait";
            document.frmSvArr.action = "/admin/etc/lotte/actRegLotteItem.asp"
            document.frmSvArr.submit();
        }
    }

}

// ������ ��ǰ �ϰ�����
function LotteEditProcess(){
    if (confirm('������ ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditAll").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditAll";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���� ��ǰ �ǸŻ��� Ȯ�� //20130305
function LotteSelectStatCheck(){
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹ� ���¸� Ȯ���Ͻðڽ��ϱ�?')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "CheckItemStat";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���� ��ǰ ���û�ǰ����ȸ //20130612
function LotteSelectStatWithOption(){
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� ����ȸ�� �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnStatWithOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "StatWithOption";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function LotteSelectEditProcess(v) {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditSelect" + v;
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function LotteSelectPriceEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditPSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditPriceSelect";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function LotteRealItemMappingSel() {
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
    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ ���Ȯ�� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnChkReal2").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "getconfirmList";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}


// ���õ� ��ǰ ���� ����
function LotteSelectItemNmEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ ���� �ϰ� ���� ��û �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditItemNm";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ǰ�� ����
function LotteSelectPoomOkEditProcess() {
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

    if (confirm('�Ե����Ŀ� �����Ͻ� ' + chkSel + '�� ��ǰ ǰ���� �ϰ� ���� ��û �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnPOEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditItemPO";
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
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

        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
        document.frmSvArr.submit();
    }

}

// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteCom","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// ������� ��ǰ
function NotInItemid()
{
	var popwin = window.open('JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteCom','notinItem','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ������� �귣��(��ü ���� �Ұ�)
function LotteNotInMakerid()
{
	var popwin = window.open('onlyLotte_Not_In_Makerid.asp','lottenotin','width=900,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ��ǰ �ּ��� ����
function ReturnCodeMappid()
{
	var popwin = window.open('JaehyuMall_ReturnCode_Mappid.asp?mallgubun=lotteCom&lotteSellyn=Y','notin','width=900,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// �ӽû�ǰ �ϰ� ����
function LotteRealItemMapping() {
	if(confirm("�ӽõ�� ��ǰ�� ���û�ǰ���� ��ϵǾ����� �ϰ� Ȯ���Ͻðڽ��ϱ�?\n\n�� ��Ż��¿����� �ټ� �ð��� �ɸ� �� �ֽ��ϴ�.")) {
		document.getElementById("btnChkReal").disabled=true;
		xLink.location.href="./lotte/actLotteCheckRDItem.asp";
	}
}

//��ǰ�ǸŻ���,����Check
function batchStatCheck(){
    xLink.location.href="./lotte/actRegLotteItem.asp?mode=CheckItemStatAuto";
}

//��ǰ�� �ٸ����� ����.
function batchItemNmCheck(){
    xLink.location.href="./lotte/actRegLotteItem.asp?mode=CheckItemNmAuto";
}

// ���õ� ��ǰ �Ǹſ��� ����
function LotteSellYnProcess(chkYn) {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n�طԵ����İ��� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� �Ե����Ŀ��� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

        document.getElementById("btnSellYn").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.mode.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "./lotte/actRegLotteItem.asp"
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
            frm.MatchCate.value="Y";
            frm.sellyn.value="A";
            frm.limityn.value="";
            frm.onlyValidMargin.checked=true;
        }else if (comp.name=="expensive10x10"){
            frm.sellyn.value="Y";                   //�Ǹ����ΰŸ�����
        }else{
            frm.LotteNotReg.value="D";
            frm.MatchCate.value="Y";
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

        if ((comp.name=="LotteYes10x10No")||(comp.name=="expensive10x10")){
            frm.MatchCate.value="";
        }

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

function popApiTest(){
    var popwin = window.open('/admin/etc/lotte/lotteApiTest.asp','lotteApiTest','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
/*
function popPoomOk(mallid,infodiv,itemid){
    var popwin = window.open('/admin/etc/lotte/popPoomOk.asp?mallid='+mallid+'&infoDiv='+infodiv+'&itemid='+itemid+'','popPoomOk','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}
*/
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//���� ���� : �Ǹ�����, ��ȸ������Ȯ���� ����.
function delexpiredItem(iitemid){
    if (confirm('�Ǹ� ����� ��ǰ�� ���� �ϼž� �մϴ�.\n\n����Ͻðڽ��ϱ�?')){
        document.frmDel.target = "xLink";
        document.frmDel.mode.value = "DelSelectExpireItem";
        document.frmDel.delitemid.value =iitemid;
        document.frmDel.action = "/admin/etc/lotte/actRegLotteItem.asp"
        document.frmDel.submit();
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
		�Ե����Ļ�ǰ��ȣ:
		<input type="text" name="lotteGoodNo" value="<%= lotteGoodNo %>" size="9" maxlength="9" class="text"> &nbsp;&nbsp;
		��ǰ��:
		<input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">&nbsp;&nbsp;
		�Ե������ӽû�ǰ��ȣ: <input type="text" name="lotteTmpGoodNo" value="<%= lotteTmpGoodNo %>" size="15" maxlength="15" class="text">
		&nbsp;
		<a href="https://partner.lotte.com/main/Login.lotte" target="_blank">�Ե�����Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 124072 | cube101010 ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ��ȣ: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		�̺�Ʈ��ȣ:
		<input type="text" name="eventid" value="<%= eventid %>" size="8" maxlength="6" class="text">
		&nbsp;
		�ֹ����ۿ��� :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		��Ͽ��� :
		<select name="LotteNotReg" class="select">
		<option value="">��ü
		<option value="M" <%= CHkIIF(LotteNotReg="M","selected","") %> >�Ե����� �̵��(��ϰ���)
		<option value="J" <%= CHkIIF(LotteNotReg="J","selected","") %> >�Ե����� �ݷ�
		<option value="W" <%= CHkIIF(LotteNotReg="W","selected","") %> >�Ե����� ��Ͽ���
		<option value="F" <%= CHkIIF(LotteNotReg="F","selected","") %> >�Ե����� ��ϿϷ�(�ӽ�)
		<option value="D" <%= CHkIIF(LotteNotReg="D","selected","") %> >�Ե����� ��ϿϷ�(����)
		<option value="R" <%= CHkIIF(LotteNotReg="R","selected","") %> >�Ե����� �������
		</select>
		&nbsp;
		    <input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>
		&nbsp;
		    <input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>
		&nbsp;
		ī�װ���Ī���� :
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
		&nbsp;
		<input type="checkbox" name="ckLimitOver" <%= ChkIIF(ckLimitOver="on","checked","") %> >���� <%=CMAXLIMITSELL%>���̻�


		<br><br>
		�ɼ� �߰� �ݾ� ���� ��ǰ ��� �Ұ�(2012-07-23)
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
		<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>�Ե����� ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
		<br>
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteYes10x10No" <%= ChkIIF(LotteYes10x10No="on","checked","") %> ><font color=red>�Ե������Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input onClick="checkQuickClick(this)" type="checkbox" name="LotteNo10x10Yes" <%= ChkIIF(LotteNo10x10Yes="on","checked","") %> ><font color=red>�Ե�����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=<%=CMAXLIMITSELL%>) ��ǰ����
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
		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
		</select>
		&nbsp;&nbsp;ǰ�������Է¿��� :
		<select name="infoDivYn" class="select">
		<option value="" <%= CHkIIF(infoDivYn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(infoDivYn="Y","selected","") %> >�Է�
		<option value="N" <%= CHkIIF(infoDivYn="N","selected","") %> >���Է�
		<option value="15" <%= CHkIIF(infoDivYn="15","selected","") %> >15
		<option value="21" <%= CHkIIF(infoDivYn="21","selected","") %> >21
		<option value="22" <%= CHkIIF(infoDivYn="22","selected","") %> >22
		<option value="23" <%= CHkIIF(infoDivYn="23","selected","") %> >23
		<option value="35" <%= CHkIIF(infoDivYn="35","selected","") %> >35
		</select>

		<!--
		<input type="checkbox" name="onreginotmapping" <%= ChkIIF(onreginotmapping="on","checked","") %> ><font color=red>�Ե����� ���&����Ī ī�װ�</font>��ǰ����
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
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();"> &nbsp;
				<input class="button" type="button" value="��� ���� �귣��(��ü�����Ұ�)" onclick="LotteNotInMakerid();"> &nbsp;
				<input class="button" type="button" id="button" value="��ǰ �ּ��� ����" onClick="ReturnCodeMappid();" >
			</td>
			<td align="right">

				<input class="button" type="button" value="Lotte���MD" onclick="pop_MDList();"> &nbsp;
				<!--
				<input class="button" type="button" value="Lotte�귣���Ī" onclick="pop_BrandList();"> &nbsp;
			-->
				<input class="button" type="button" value="Lotteī�װ���Ī" onclick="pop_CateManager();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal2" value="���û�ǰ ���Ȯ��" onClick="LotteRealItemMappingSel();">

				&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkReal" value="�ӽû�ǰ �ϰ�Ȯ�� [200�Ǿ�]" onClick="LotteRealItemMapping();">

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
			    <input class="button" type="button" id="btnRegSel" value="��ǰ ���" onClick="LotteSelectRegProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnEditSel" value="��ǰ����/���� ����" onClick="LotteSelectEditProcess('');">
			    &nbsp;&nbsp;
	    	    <!--
			    <input class="button" type="button" id="btnEditAll" value="������ ��ǰ �Ե��������� �ϰ����� [10�Ǿ�]" onClick="LotteEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegAll" value="�̵�� ��ǰ �Ե��������� �ϰ���� [10�Ǿ�]" onClick="LotteRegProcess();">
			    -->
			    <input class="button" type="button" id="btnEditPSel" value="���� ����" onClick="LotteSelectPriceEditProcess();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="��ǰ �����ȸ" onClick="LotteSelectcheckStock();" >
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="��ǰ�� ������û" onClick="LotteSelectItemNmEditProcess();" >
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnPOEditSel" value="��ǰ ǰ�� ����" onClick="LotteSelectPoomOkEditProcess();" >
			    <br><br>
				�������� ���� :
			    <!--
			    <input class="button" type="button" id="btnDelJehyu" value="���޸��ƴѰ� �ϰ����� [20�Ǿ�]" onClick="LotteDelJaeHyuProcess();">
			    -->
			    <input class="button" type="button" id="btnRegSel" value="��ǰ ���" onClick="LotteregIMSI(true);">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnRegSel" value="��ǰ ����" onClick="LotteregIMSI(false);" >
			    &nbsp;&nbsp;
			    <br><br>
			    <input class="button" type="button" id="btnEditSel" value="���û�ǰ���� ����(��ϴ��)" onClick="LotteSelectEditProcess('2');">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnStatConfirm" value="���û�ǰ �ǸŻ���Ȯ��" onClick="LotteSelectStatCheck();">
			    &nbsp;&nbsp;
			    <input class="button" type="button" id="btnStatWithOption" value="���û�ǰ ���û�ǰ����ȸ" onClick="LotteSelectStatWithOption();">
			</td>
			<td align="right" valign="top">

				���û�ǰ��
				<Select name="chgSellYn" class="select">
				<option value="N"  >ǰ��</option>
				<option value="Y"  >�Ǹ���</option>
				<% if (True) or (reqExpire="on") then %>
				<option value="X" >�Ǹ�����(����)</option><!-- �����ϸ� ���� ���� �� �� ���� -->
				<% end if %>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="LotteSellYnProcess(frmReg.chgSellYn.value);">

				<% if (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") then %>
				<br><input type="button" value="API_TEST(������)" onClick="popApiTest();">
				<% end if %>
				<br><br><input type="button" value="�ǸŻ���Check(������)" onClick="batchStatCheck();">
				&nbsp;&nbsp;<input type="button" value="��ǰ�����(������)" onClick="batchItemNmCheck();">

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
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="brandid" value="">
<input type="hidden" name="chgSellYn" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oLotteitem.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(oLotteitem.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td >�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">�Ե����ĵ����<br>�Ե���������������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�Ե�����<br>���ݹ��Ǹ�</td>
	<td width="70">�Ե�����<br>��ǰ��ȣ<br>(�ӽù�ȣ)</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">ǰ��</td>
</tr>
<% for i=0 to oLotteitem.FResultCount - 1 %>
<%
    outMallItemArr = outMallItemArr & null2blank(oLotteitem.FItemList(i).FLotteGoodNo) & ","
    outMallItemArr = replace(outMallItemArr,",,",",")
%>
<tr bgcolor="#FFFFFF" height="20">
    <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oLotteitem.FItemList(i).FItemID %>"></td>
    <td><img src="<%= oLotteitem.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oLotteitem.FItemList(i).FItemID %>','lotteCom','<%=oLotteitem.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
    <td align="center"><%= oLotteitem.FItemList(i).FItemID %>
    <% if oLotteitem.FItemList(i).FLimitYn="Y" then %><br><font color=blue>����:<%= oLotteitem.FItemList(i).getLimitEa %></font><% end if %>
    </td>
    <td><%= oLotteitem.FItemList(i).FMakerid %> <%= oLotteitem.FItemList(i).getDeliverytypeName %><br><%= oLotteitem.FItemList(i).FItemName %></td>
    <td align="center"><%= oLotteitem.FItemList(i).FRegdate %><br><%= oLotteitem.FItemList(i).FLastupdate %></td>
    <td align="center"><%= oLotteitem.FItemList(i).FLotteRegdate %><br><%= oLotteitem.FItemList(i).FLotteLastUpdate %></td>
    <td align="right">
        <% if oLotteitem.FItemList(i).FSaleYn="Y" then %>
        <strike><%= FormatNumber(oLotteitem.FItemList(i).FOrgPrice,0) %></strike><br>
        <font color="#CC3333"><%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %></font>
        <% else %>
        <%= FormatNumber(oLotteitem.FItemList(i).FSellcash,0) %>
        <% end if %>
    </td>
    <td align="center">
        <% if oLotteitem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oLotteitem.FItemList(i).Fbuycash/oLotteitem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
    </td>
    <td align="center">
        <% if oLotteitem.FItemList(i).IsSoldOut then %>
            <% if oLotteitem.FItemList(i).FSellyn="N" then %>
            <font color="red">ǰ��</font>
            <% else %>
            <font color="red">�Ͻ�<br>ǰ��</font>
            <% end if %>
        <% end if %>
    </td>
	<td align="center">
	<%
		If oLotteitem.FItemList(i).FItemdiv = "06" OR oLotteitem.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
    <td align="center">
    <% if Not IsNULL(oLotteitem.FItemList(i).FLottePrice) then %>
        <% if (oLotteitem.FItemList(i).Fsellcash<>oLotteitem.FItemList(i).FLottePrice) then %>
        <strong><%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %></strong>
        <% else %>
        <%= formatNumber(oLotteitem.FItemList(i).FLottePrice,0) %>
        <% end if %>
        <br>
        <% if (oLotteitem.FItemList(i).FLotteSellYn="X") then %>
        <a href="javascript:delexpiredItem('<%= oLotteitem.FItemList(i).FItemID %>');">
        <% end if %>
            <% if (oLotteitem.FItemList(i).FSellyn<>oLotteitem.FItemList(i).FLotteSellYn) then %>
            <strong><%= oLotteitem.FItemList(i).FLotteSellYn %></strong>
            <% else %>
            <%= oLotteitem.FItemList(i).FLotteSellYn %>
            <% end if %>
        <% if (oLotteitem.FItemList(i).FLotteSellYn="X") then %>
        </a>
        <% end if %>
    <% end if %>
    </td>
    <td align="center">
    <%
    	'#�ǻ�ǰ��ȣ
    	if Not(IsNULL(oLotteitem.FItemList(i).FLotteGoodNo)) then
        	Response.Write "<a target='_blank' href='http://www.lotte.com/goods/viewGoodsDetail.lotte?goods_no="&oLotteitem.FItemList(i).FLotteGoodNo&"'>"&oLotteitem.FItemList(i).FLotteGoodNo&"</a>"
		else
			'#�ӽû�ǰ��ȣ
			if Not(IsNULL(oLotteitem.FItemList(i).FLotteTmpGoodNo)) then
				if oLotteitem.FItemList(i).FLotteStatCd<>"30" then
					Response.Write oLotteitem.FItemList(i).getLotteItemStatCd & "<br>(" & oLotteitem.FItemList(i).FLotteTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if
	%>
    </td>
    <td align="center"><%= oLotteitem.FItemList(i).Freguserid %></td>
    <td align="center"><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','0');"><%= oLotteitem.FItemList(i).FoptionCnt %>:<%= oLotteitem.FItemList(i).FregedOptCnt %></a></td>
    <td align="center"><%= oLotteitem.FItemList(i).FrctSellCNT %></td>
    <td align="center">
    <% if oLotteitem.FItemList(i).FCateMapCnt>0 then %>
	    ��Ī��
    <% else %>
    	<font color="darkred">��Ī�ȵ�</font>
    <% end if %>

    <% if (oLotteitem.FItemList(i).FaccFailCNT>0) then %>
        <br><font color="red" title="<%= oLotteitem.FItemList(i).FlastErrStr %>">ERR:<%= oLotteitem.FItemList(i).FaccFailCNT %></font>
    <% end if %>
    </td>
    <td align="center"><%=oLotteitem.FItemList(i).FinfoDiv%>
    <% if (oLotteitem.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oLotteitem.FItemList(i).FItemID%>','1');">
    <font color="<%=CHKIIF(oLotteitem.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">�ɼǱݾ�</font>
    <% if oLotteitem.FItemList(i).FoptAddPrcRegType<>0 then %>
    (<%=oLotteitem.FItemList(i).FoptAddPrcRegType%>)
    <% end if %>
    </a>
    <% end if %>
    </td>
</tr>
<% next %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if oLotteitem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotteitem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oLotteitem.StartScrollPage to oLotteitem.FScrollCount + oLotteitem.StartScrollPage - 1 %>
    		<% if i>oLotteitem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oLotteitem.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
<tr height="20">
    <td colspan="16" align="left" bgcolor="#FFFFFF">
    <% if Right(outMallItemArr,1)="," then outMallItemArr=Left(outMallItemArr,Len(outMallItemArr)-1) %>
    <%= outMallItemArr %>

    <% if C_ADMIN_AUTH then %>
    <br><%=lotteAuthNo%>
    <% end if %>
    </td>
</tr>

</table>
</form>
<form name="frmDel" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="delitemid" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
<% set oLotteitem = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
