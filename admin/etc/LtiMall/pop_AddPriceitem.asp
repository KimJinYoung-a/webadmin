<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim mallid, infoLoop, infoDivValue
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, ltimallgoodno, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, priceOption, ltimalltmpgoodno, deliverytype, mwdiv
Dim expensive10x10, diffPrc, ltimallYes10x10No, ltimallNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research
Dim oiMall

mallid					= CMALLNAME
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
ltimallgoodno			= request("ltimallgoodno")
ltimalltmpgoodno		= request("ltimalltmpgoodno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
ltimallYes10x10No		= request("ltimallYes10x10No")
ltimallNo10x10Yes		= request("ltimallNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "J"
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
'�Ե�iMall ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If ltimallgoodno <> "" then
	Dim iA2, arrTemp2, arrltimallgoodno
	ltimallgoodno = replace(ltimallgoodno,",",chr(10))
	ltimallgoodno = replace(ltimallgoodno,chr(13),"")
	arrTemp2 = Split(ltimallgoodno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrltimallgoodno = arrltimallgoodno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	ltimallgoodno = left(arrltimallgoodno,len(arrltimallgoodno)-1)
End If

'�Ե�iMall ������ ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If ltimalltmpgoodno <> "" then
	Dim iA3, arrTemp3, arrltimalltmpgoodno
	ltimalltmpgoodno = replace(ltimalltmpgoodno,",",chr(10))
	ltimalltmpgoodno = replace(ltimalltmpgoodno,chr(13),"")
	arrTemp3 = Split(ltimalltmpgoodno,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			If Not(isNumeric(trim(arrTemp3(iA3)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp3(iA3) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrltimalltmpgoodno = arrltimalltmpgoodno & trim(arrTemp3(iA3)) & ","
			End If
		End If
		iA3 = iA3 + 1
	Loop
	ltimalltmpgoodno = left(arrltimalltmpgoodno,len(arrltimalltmpgoodno)-1)
End If

Set oiMall = new CLotteiMall
If (session("ssBctID")="kjy8517") Then
	oiMall.FPageSize					= 50
Else
	oiMall.FPageSize					= 20
End If
	oiMall.FCurrPage					= page
	oiMall.FRectMakerid					= makerid
	oiMall.FRectItemID					= itemid
	oiMall.FRectItemName				= itemname
	oiMall.FRectLTiMallGoodNo			= ltimallgoodno
	oiMall.FRectLTimallTmpGoodNo		= ltimalltmpgoodno
	oiMall.FRectCDL						= request("cdl")
	oiMall.FRectCDM						= request("cdm")
	oiMall.FRectCDS						= request("cds")
	oiMall.FRectExtNotReg				= ExtNotReg
	oiMall.FRectIsReged					= isReged
	oiMall.FRectNotinmakerid			= notinmakerid
	oiMall.FRectPriceOption				= priceOption
	oiMall.FRectDeliverytype			= deliverytype
	oiMall.FRectMwdiv					= mwdiv

	oiMall.FRectSellYn					= sellyn
	oiMall.FRectLimitYn					= limityn
	oiMall.FRectSailYn					= sailyn
	oiMall.FRectonlyValidMargin			= onlyValidMargin
	oiMall.FRectIsMadeHand				= isMadeHand
	oiMall.FRectIsOption				= isOption
	oiMall.FRectInfoDiv					= infoDiv
	oiMall.FRectExtSellYn				= extsellyn
	oiMall.FRectFailCntExists			= failCntExists
	oiMall.FRectMatchCate				= MatchCate
	oiMall.FRectExpensive10x10			= expensive10x10
	oiMall.FRectdiffPrc					= diffPrc
	oiMall.FRectLtimallYes10x10No		= ltimallYes10x10No
	oiMall.FRectLtimallNo10x10Yes		= ltimallNo10x10Yes
	oiMall.FRectReqEdit					= reqEdit
If (bestOrd = "on") Then
    oiMall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oiMall.FRectOrdType = "BM"
End If
	oiMall.getLTiMallAddOptionRegedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=lotteimall","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=lotteimall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=lotteimall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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
	if ((comp.name!="ltimallYes10x10No")&&(frm.ltimallYes10x10No.checked)){ frm.ltimallYes10x10No.checked=false }
	if ((comp.name!="ltimallNo10x10Yes")&&(frm.ltimallNo10x10Yes.checked)){ frm.ltimallNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
function checkisReged(comp){
    if (comp.name=="isReged"){
    	if (document.getElementById("AR").checked == true){
    		comp.form.ExtNotReg.value = "J"
   			comp.form.ExtNotReg.disabled = true;
   		}else if(document.getElementById("QR").checked == true){
    		comp.form.ExtNotReg.value = "J"
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

    if ((comp.name=="ltimallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="ltimallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.ltimallYes10x10No.checked){
            comp.form.ltimallYes10x10No.checked = false;
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
	if ((comp.name!="ltimallYes10x10No")&&(frm.ltimallYes10x10No.checked)){ frm.ltimallYes10x10No.checked=false }
	if ((comp.name!="ltimallNo10x10Yes")&&(frm.ltimallNo10x10Yes.checked)){ frm.ltimallNo10x10Yes.checked=false }
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
	var pCM6 = window.open("/admin/etc/que/popQueOptionLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM6.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function LtimallPriceEditProcess() {
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
        document.getElementById("btnEditPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ�� ����
function LtimallItemnameEditProcess() {
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

    if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ���� ���� �Ͻðڽ��ϱ�?\n\n�طԵ�iMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditName").disabled=true;
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "EditItemNm";
        document.frmSvArr.cmdparam.value = "ITEMNAME";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function LtimallSelectRegProcess(isreal) {
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
			document.frmSvArr.cmdparam.value = "REG";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?\n\n��30�д����� ��ġ ��ϵ˴ϴ�.')){
            document.getElementById("btnRegSel").disabled=true;
            document.frmSvArr.target = "xLink";
            document.frmSvArr.cmdparam.value = "RegSelectWait";
            document.frmSvArr.action = "/admin/etc/Ltimall/actLotteiMallReq.asp"
            document.frmSvArr.submit();
        }
    }
}

//���� �űԻ�ǰ��ȸ
function LtimallStatCheckProcess(){
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

    if (confirm('�Ե����̸��� �����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹ� ���¸� Ȯ�� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        //document.frmSvArr.cmdparam.value = "CheckItemStat";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}

//���û�ǰ ���û�ǰ��ȸ
function LtimallDispViewProcess(){
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

    if (confirm('�Ե����̸��� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���û�ǰ ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
         document.frmSvArr.cmdparam.value = "DISPVIEW";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function LtimallEditProcess(v) {
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
		if(v == ""){
			document.getElementById("btnEditSel").disabled=true;
		}else{
			document.getElementById("btnEditSel2").disabled=true;
		}
		document.frmSvArr.target = "xLink";
		//document.frmSvArr.cmdparam.value = "EditSelect" + v;
		document.frmSvArr.cmdparam.value = "EDIT" + v;
		document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ���
function LtimallregIMSI(isreg) {
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
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('�Ե�iMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/ltimallAddOpt/actLotteiMallReq.asp"
			document.frmSvArr.submit();
        }
    }
}
function chgItemname(idx, iname){
	document.frmUp.target = "xLink";
    document.frmUp.idx.value = idx
    document.frmUp.cName.value = document.getElementById(iname).value;
    document.frmUp.mode.value = "chgName"
    document.frmUp.action = "/admin/etc/optManager/optProc.asp"
    document.frmUp.submit();
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
		<a href="https://partner.lotteimall.com/main/Login.lotte" target="_blank">�Ե����̸�Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ 011799LT | store101010* | 01063242821 ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		�Ե�iMall ��ǰ�ڵ� : <textarea rows="2" cols="20" name="ltimallgoodno" id="itemid"><%=replace(ltimallgoodno,",",chr(10))%></textarea>
		&nbsp;
		������ ��ǰ�ڵ� : <textarea rows="2" cols="20" name="ltimalltmpgoodno" id="itemid"><%=replace(ltimalltmpgoodno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >�Ե�iMall ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >�Ե�iMall ��Ͽ����̻�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >�Ե�iMall ��Ͽ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >�Ե�iMall ���۽õ��߿���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >�Ե�iMall �ݷ�
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >�Ե�iMall ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >�Ե�iMall ��ϿϷ�(����)
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
		�ٹ����� :
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		�Ǹ�
		<select name="sellyn" class="select">
			<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
		</select>&nbsp;
		����
		<select name="limityn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
			<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
		</select>&nbsp;
		����
		<select name="sailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
		</select>&nbsp;
		���ظ���(<%= Chkiif(mallid="lotteimall", "14.9", "") %>%)
		<select name="onlyValidMargin" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(onlyValidMargin="Y","selected","") %> >�����̻�
			<option value="N" <%= CHkIIF(onlyValidMargin="N","selected","") %> >��������
		</select>&nbsp;
		�ֹ�����
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>&nbsp;
		ǰ��
		<select name="infodiv" class="select">
			<option value="" <%= CHkIIF(infoDiv="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >�Է�
			<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >���Է�
		<%
			For infoLoop = 1 To 35
				If infoLoop < 10 Then
					infoDivValue = "0"&infoLoop
				Else
					infoDivValue = infoLoop
				End If
		%>
			<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
		<% Next %>
		</select>
		<br>
		���޸� &nbsp;&nbsp; :
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		�Ǹ�
		<select name="extsellyn" class="select">
			<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
			<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
			<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
			<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
		</select>&nbsp;
		����
		<select name="failCntExists" class="select">
			<option value="" <%= CHkIIF(failCntExists="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >��ϼ�������1ȸ�̻�
			<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >��ϼ�������0ȸ
		</select>&nbsp;
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>�Ե�iMall ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="ltimallYes10x10No" <%= ChkIIF(ltimallYes10x10No="on","checked","") %> ><font color=red>�Ե�iMall�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="ltimallNo10x10Yes" <%= ChkIIF(ltimallNo10x10Yes="on","checked","") %> ><font color=red>�Ե�iMallǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p>
<!-- �׼� ���� -->
<form name="frmUp" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="idx" value="">
<input type="hidden" name="cName" value="">
<input type="hidden" name="mode" value="">
</form>

<form name="frmReg" method="post" action="lotteimallItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('lotteimall');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="LtimallSelectRegProcess(true);">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="LtimallPriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="����&����&�ɼ�&���&����" onClick="LtimallEditProcess('');">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditName" value="��ǰ��" onClick="LtimallItemnameEditProcess();">
				<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="���û�ǰ��ȸ" onClick="LtimallDispViewProcess();">
				<% End If %>
				<br><br>
				���ο��� ��ǰ :
				<!--
				<input class="button" type="button" id="btnEditSel2" value="����" onClick="LtimallEditProcess('2');">
				&nbsp;&nbsp;
				-->
				<input class="button" type="button" id="btnEditInfoDiv" value="�űԻ�ǰ��ȸ" onClick="LtimallStatCheckProcess();">
				<br><br>
				��Ͽ��� ��ǰ :
				<input class="button" type="button" id="btnRegImsi" value="���" onClick="LtimallregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="����" onClick="LtimallregIMSI(false);" >
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
					<option value="X">�Ǹ�����(����)</option><!-- �����ϸ� ���� ���� �� �� ���� -->
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="LTiMallSellYnProcess(frmReg.chgSellYn.value);">
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
		�˻���� : <b><%= FormatNumber(oiMall.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oiMall.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">��ǰ�ڵ�<br>�ɼ��ڵ�</td>
	<td>��ǰ����</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">�Ե�iMall�����<br>�Ե�iMall����������</td>
	<td width="70">�ǸŰ�<br><font color="purple">�ɼǰ�</font></td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�Ե�iMall<br>���ݹ��Ǹ�</td>
	<td width="70">�Ե�iMall<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i = 0 To oiMall.FResultCount - 1 %>
<% If (oiMall.FItemList(i).FItemName <> oiMall.FItemList(i).FRegedItemname) OR (oiMall.FItemList(i).FOptionname <> oiMall.FItemList(i).FRegedOptionname) Then %>
<tr align="center" bgcolor="GOLD">
<% Else %>
<tr align="center" bgcolor="#FFFFFF">
<% End If %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oiMall.FItemList(i).FIdx %>"></td>
	<td><img src="<%= oiMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oiMall.FItemList(i).FItemID %>','lotteimall','<%=oiMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oiMall.FItemList(i).FItemID%>" target="_blank"><%= oiMall.FItemList(i).FItemID %></a>
		<br><font color="gray"><%=oiMall.FItemList(i).FItemOption%></font><br>
		<% if oiMall.FItemList(i).FLimitYn="Y" then %><br><%= oiMall.FItemList(i).getLimitHtmlStr %></font><% end if %>
	</td>
	<td align="left">
		<%= oiMall.FItemList(i).FMakerid %><%= oiMall.FItemList(i).getDeliverytypeName %>
		<br>(��) : <%= oiMall.FItemList(i).FItemName %>
		<br>(��) : <%= oiMall.FItemList(i).FRegedItemname %>
		<br>(��) : <%= oiMall.FItemList(i).FOptionname %>
		<br>(��) : <%= oiMall.FItemList(i).FRegedOptionname %>
		<br><input type="text" style="color:red" id="newitemname<%=oiMall.FItemList(i).FIdx%>" size="50" value="<%= oiMall.FItemList(i).getRealItemname %>">
		<input type="button" class="button" value="����" onclick="chgItemname('<%= oiMall.FItemList(i).FIdx %>', 'newitemname<%=oiMall.FItemList(i).FIdx%>')">
	</td>
	<td align="center"><%= oiMall.FItemList(i).FRegdate %><br><%= oiMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oiMall.FItemList(i).FLtimallRegdate %><br><%= oiMall.FItemList(i).FLtimallLastUpdate %></td>

	<td align="right">
	<% If oiMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oiMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oiMall.FItemList(i).FSellcash,0) %>
	<% End If %>
		<br><font color="purple">+<%= oiMall.FItemList(i).FOptaddprice %></font>
	</td>
	<td align="center">
	<%
		If oiMall.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-oiMall.FItemList(i).Fbuycash/oiMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oiMall.FItemList(i).IsSoldOut Then
			If oiMall.FItemList(i).FSellyn = "N" Then
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
		If oiMall.FItemList(i).FItemdiv = "06" OR oiMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oiMall.FItemList(i).FLtimallStatCd > 0) Then
			If Not IsNULL(oiMall.FItemList(i).FLtimallPrice) Then
				If (oiMall.FItemList(i).Fsellcash + oiMall.FItemList(i).FOptaddprice <> oiMall.FItemList(i).FLtimallPrice) Then
	%>
					<strong><%= formatNumber(oiMall.FItemList(i).FLtimallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oiMall.FItemList(i).FLtimallPrice,0)&"<br>"
				End If

				If (oiMall.FItemList(i).FSellyn="Y" and oiMall.FItemList(i).FLtimallSellYn<>"Y") or (oiMall.FItemList(i).FSellyn<>"Y" and oiMall.FItemList(i).FLtimallSellYn="Y") Then
	%>
					<strong><%= oiMall.FItemList(i).FLtimallSellYn %></strong>
	<%
				Else
					response.write oiMall.FItemList(i).FLtimallSellYn
				End If
			End If
		End If
	%>
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
					Response.Write "<br>(" & oiMall.FItemList(i).FLtiMallTmpGoodNo & ")"
				end if
			else
				Response.Write "<img src='/images/i_delete.gif' width='8' height='9' border='0'>"
			end if
		end if

		if (oiMall.FItemList(i).FLTiMallStatCd<>7) then
		    response.write "<br>"&oiMall.FItemList(i).getLTIMallStatCDName
		end if
	%>
	</td>
	<td align="center"><%= oiMall.FItemList(i).Freguserid %></td>
	<td align="center"><%= oiMall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<% If oiMall.FItemList(i).FCateMapCnt > 0 Then %>
	    ��Ī��
	<% Else %>
		<font color="darkred">��Ī�ȵ�</font>
	<% End If %>

	<% If (oiMall.FItemList(i).FaccFailCNT > 0) Then %>
	    <br><font color="red" title="<%= oiMall.FItemList(i).FlastErrStr %>">ERR:<%= oiMall.FItemList(i).FaccFailCNT %></font>
	<% End If %>
	</td>
    <td align="center"><%= oiMall.FItemList(i).FinfoDiv %>
    <% if (oiMall.FItemList(i).FoptAddPrcCnt>0) then %>
    <br><a href="javascript:popManageOptAddPrc('<%=oiMall.FItemList(i).FItemID%>','1');">
    	<font color="<%=CHKIIF(oiMall.FItemList(i).FoptAddPrcRegType<>0,"gray","red")%>">�ɼǱݾ�</font>
	    <% If oiMall.FItemList(i).FoptAddPrcRegType<>0 Then %>
	    (<%=oiMall.FItemList(i).FoptAddPrcRegType%>)
	    <% End If %>
    	</a>
    <% End If %>
    </td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
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
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% SET oiMall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->