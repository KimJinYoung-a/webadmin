<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/11st/11stcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, st11KeepSell, isSpecialPrice
Dim bestOrdMall, st11GoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, st11Yes10x10No, st11No10x10Yes, reqEdit, reqExpire, failCntExists, priceOption
Dim page, i, research, st11GoodNoArray, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim o11st, xl, kjypageSize
Dim startMargin, endMargin
Dim purchasetype
page    				= request("page")
kjypageSize				= request("kjypageSize")
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
st11GoodNo				= request("st11GoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
st11Yes10x10No			= request("st11Yes10x10No")
st11No10x10Yes			= request("st11No10x10Yes")
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
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = ""
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

'11���� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If st11GoodNo <> "" then
	Dim iA2, arrTemp2, arrst11GoodNo
	st11GoodNo = replace(st11GoodNo,",",chr(10))
	st11GoodNo = replace(st11GoodNo,chr(13),"")
	arrTemp2 = Split(st11GoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrst11GoodNo = arrst11GoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	st11GoodNo = left(arrst11GoodNo,len(arrst11GoodNo)-1)
End If

Set o11st = new C11st
	o11st.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	o11st.FPageSize					= kjypageSize
Else
	o11st.FPageSize					= 100
End If
	o11st.FRectCDL					= request("cdl")
	o11st.FRectCDM					= request("cdm")
	o11st.FRectCDS					= request("cds")
	o11st.FRectItemID				= itemid
	o11st.FRectItemName				= itemname
	o11st.FRectSellYn				= sellyn
	o11st.FRectLimitYn				= limityn
	o11st.FRectSailYn				= sailyn
	o11st.FRectStartMargin			= startMargin
	o11st.FRectEndMargin			= endMargin
	o11st.FRectMakerid				= makerid
	o11st.FRectSt11GoodNo			= st11GoodNo
	o11st.FRectMatchCate			= MatchCate
	o11st.FRectIsMadeHand			= isMadeHand
	o11st.FRectIsOption				= isOption
	o11st.FRectIsReged				= isReged
	o11st.FRectNotinmakerid			= notinmakerid
	o11st.FRectNotinitemid			= notinitemid
	o11st.FRectExcTrans				= exctrans
	o11st.FRectPriceOption			= priceOption
	o11st.FRectIsSpecialPrice     	= isSpecialPrice
	o11st.FRectDeliverytype			= deliverytype
	o11st.FRectMwdiv				= mwdiv
	o11st.FRectScheduleNotInItemid	= scheduleNotInItemid
	o11st.FRectIsextusing			= isextusing
	o11st.FRectCisextusing			= cisextusing
	o11st.FRectRctsellcnt			= rctsellcnt

	o11st.FRectExtNotReg			= ExtNotReg
	o11st.FRectExpensive10x10		= expensive10x10
	o11st.FRectdiffPrc				= diffPrc
	o11st.FRectSt11Yes10x10No		= st11Yes10x10No
	o11st.FRectSt11No10x10Yes		= st11No10x10Yes
	o11st.FRectSt11KeepSell			= st11KeepSell
	o11st.FRectExtSellYn			= extsellyn
	o11st.FRectInfoDiv				= infoDiv
	o11st.FRectFailCntOverExcept	= ""
	o11st.FRectFailCntExists		= failCntExists
	o11st.FRectReqEdit				= reqEdit
	o11st.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    o11st.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    o11st.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	o11st.get11streqExpireItemList
Else
	o11st.get11stRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=11stList"& replace(DATE(), "-", "") &"_xl.xls"
Else
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=11st1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=11st1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=11st1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=11st1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="st11KeepSell")&&(frm.st11KeepSell.checked)){ frm.st11KeepSell.checked=false }
	if ((comp.name!="diffPrc")&&(frm.diffPrc.checked)){ frm.diffPrc.checked=false }
	if ((comp.name!="st11Yes10x10No")&&(frm.st11Yes10x10No.checked)){ frm.st11Yes10x10No.checked=false }
	if ((comp.name!="st11No10x10Yes")&&(frm.st11No10x10Yes.checked)){ frm.st11No10x10Yes.checked=false }
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

    if ((comp.name=="st11Yes10x10No")&&(comp.checked)){
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

    if ((comp.name=="st11No10x10Yes")&&(comp.checked)){
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
        if (comp.form.st11Yes10x10No.checked){
            comp.form.st11Yes10x10No.checked = false;
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
	if ((comp.name!="st11Yes10x10No")&&(frm.st11Yes10x10No.checked)){ frm.st11Yes10x10No.checked=false }
	if ((comp.name!="st11No10x10Yes")&&(frm.st11No10x10Yes.checked)){ frm.st11No10x10Yes.checked=false }
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
	var pCM2 = window.open("/admin/etc/11st/pop11stcateList.asp","popCate11stmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
//�����ڵ� �˻�
function fn11stCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "<%=apiURL%>/outmall/11st/act11stReq.asp?cmdparam=11stCommonCode&CommCD="+ccd+"";
	}
}

// ���õ� ��ǰ �ϰ� ���
function st11SelectRegProcess() {
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

    if (confirm('11������ �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function st11SellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function st11PriceEditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditPrice").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��� ��ȸ
function st11checkStockProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��� ��ȸ �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTOCK";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ٹ����� ��ǰ�ڵ�� 11���� ��ǰ ��ȸ
function st11ChkStatProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� ��ȸ �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "CHKSTAT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function st11EditProcess() {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
		document.frmSvArr.submit();
    }
}
//��ǰ ����
function st11DeleteProcess(){
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
    if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n11���� ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
			document.frmSvArr.submit();
		}
    }
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=11st1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

<% if request("auto") = "Y" then %>
function st11EditProcessAuto() {
	var cnt = <%= o11st.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/11st/act11stReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= o11st.FResultCount %>;
	if (cnt === 0) {
		// 45�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		st11EditProcessAuto();
		// 5�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 5*60*1000);
	}
}

$(document).ready(function() {
    $('table').hide();
});
<% end if %>
function popXL()
{
    frmXL.submit();
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
		<a href="http://soffice.11st.co.kr" target="_blank">11����Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ tenbyten001 | store10x10* ]</font>"
			End If
		%>
		&nbsp;<a href="http://openapi.11st.co.kr/openapi/OpenApiGuide.tmall" target="_blank">11���� ���߰��̵�</a>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		11���� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="st11GoodNo" id="itemid"><%=replace(st11GoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >11���� ��Ͻõ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >11���� ��Ͽ����̻�
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >11���� ��ϿϷ�(����)
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
		<% If (session("ssBctID")="kjy8517") Then %>
			<input class="text" size="5" type="text" name="kjypageSize" value="<%= kjypageSize %>">
		<% End If %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>11���� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="st11Yes10x10No" <%= ChkIIF(st11Yes10x10No="on","checked","") %> ><font color=red>11�����Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="st11No10x10Yes" <%= ChkIIF(st11No10x10Yes="on","checked","") %> ><font color=red>11����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="st11KeepSell" <%= ChkIIF(st11KeepSell="on","checked","") %> ><font color=red>�Ǹ�����</font> �ؾ��� ��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����), ������ �ø�ó��(11������ ������ �Ⱦ�)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, ��۹���� �ù�(�Ϲ�) �ƴѰ�, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ɼǰ� 0�� �Ǹ��� ��ǰ�� ����(�ɼ� �������� 5�� ���� ����), �ɼǰ��� �ǸŰ� 100% �̻��� ��ǰ

<p />
<% end if %>
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
				<input class="button" type="button" value="��� ���� ��ǰ" onclick="NotInItemid();">&nbsp;
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('11st1010');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="st11SelectRegProcess();">
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="st11PriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="�����ȸ" onClick="st11checkStockProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStock" value="����" onClick="st11EditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnChkStat" value="��ȸ" onClick="st11ChkStatProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="��ǰ����" onClick="st11DeleteProcess();" style=font-weight:bold>
			<% End If %>
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="category">ī�װ�
					<option value="outboundarea">������ּ��ڵ�
					<option value="inboundarea">��ǰ��ȯ���ּ��ڵ�
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="fn11stCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="st11SellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="delitemid" value="">
<input type="hidden" name="chgSellYn" value="">
<input type="hidden" name="chgStatItemCode" value="">
<input type="hidden" name="ckLimit">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= FormatNumber(o11st.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(o11st.FTotalPage,0) %></b>
	</td>
	<td align="right">
		<input type="button" class="button" value="�����ޱ�" onClick="popXL()">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
<% If (xl <> "Y") Then %>
	<td width="50">�̹���</td>
<% End If %>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">11st�����<br>11st����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">11st<br>���ݹ��Ǹ�</td>
	<td width="70">11st<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to o11st.FResultCount - 1 %>
<tr align="center" <% If o11st.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= o11st.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= o11st.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= o11st.FItemList(i).FItemID %>','11st1010','<%=o11st.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=o11st.FItemList(i).FItemID%>" target="_blank"><%= o11st.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If o11st.FItemList(i).FSt11Statcd <> 7 Then
	%>
		<br><%= o11st.FItemList(i).getSt11StatName %>
	<%
			End If
			response.write o11st.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= o11st.FItemList(i).FMakerid %> <%= o11st.FItemList(i).getDeliverytypeName %><br><%= o11st.FItemList(i).FItemName %></td>
	<td align="center"><%= o11st.FItemList(i).FRegdate %><br><%= o11st.FItemList(i).FLastupdate %></td>
	<td align="center"><%= o11st.FItemList(i).FSt11Regdate %><br><%= o11st.FItemList(i).FSt11LastUpdate %></td>
	<td align="right">
		<% If o11st.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(o11st.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(o11st.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(o11st.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If o11st.FItemList(i).Fsellcash = 0 Then
			'//
		' elseIf (o11st.FItemList(i).FSt11StatCd > 0) and Not IsNULL(o11st.FItemList(i).FSt11Price) Then
		' 	If (o11st.FItemList(i).FSaleYn = "Y") and (CLng((1.0*o11st.FItemList(i).FSellcash/10)*10) < o11st.FItemList(i).FSt11Price) Then
		' 		'// ���޸� ���� �Ǹ���
		' %>
		' <strike><%= CLng(10000-o11st.FItemList(i).Fbuycash/o11st.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-o11st.FItemList(i).Fbuycash/o11st.FItemList(i).FSt11Price*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-o11st.FItemList(i).Fbuycash/o11st.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (o11st.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (o11st.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-o11st.FItemList(i).FOrgSuplycash/o11st.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-o11st.FItemList(i).Fbuycash/o11st.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-o11st.FItemList(i).Fbuycash/o11st.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If o11st.FItemList(i).IsSoldOut Then
			If o11st.FItemList(i).FSellyn = "N" Then
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
		If o11st.FItemList(i).FItemdiv = "06" OR o11st.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (o11st.FItemList(i).FSt11StatCd > 0) Then
			If Not IsNULL(o11st.FItemList(i).FSt11Price) Then
				If (o11st.FItemList(i).Mustprice <> o11st.FItemList(i).FSt11Price) Then
	%>
					<strong><%= formatNumber(o11st.FItemList(i).FSt11Price,0) %></strong><br>
	<%
				Else
					response.write formatNumber(o11st.FItemList(i).FSt11Price,0)&"<br>"
				End If

				If Not IsNULL(o11st.FItemList(i).FSpecialPrice) Then
					If (now() >= o11st.FItemList(i).FStartDate) And (now() <= o11st.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(o11st.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (o11st.FItemList(i).FSellyn="Y" and o11st.FItemList(i).FSt11SellYn<>"Y") or (o11st.FItemList(i).FSellyn<>"Y" and o11st.FItemList(i).FSt11SellYn="Y") Then
	%>
					<strong><%= o11st.FItemList(i).FSt11SellYn %></strong>
	<%
				Else
					response.write o11st.FItemList(i).FSt11SellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If o11st.FItemList(i).FSt11GoodNo <> "" Then %>
			<a target="_blank" href="http://www.11st.co.kr/product/SellerProductDetail.tmall?method=getSellerProductDetail&prdNo=<%=o11st.FItemList(i).FSt11GoodNo%>"><%=o11st.FItemList(i).FSt11GoodNo%></a>
		<% End If %>
	</td>
	<td align="center"><%= o11st.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=o11st.FItemList(i).FItemID%>','0');"><%= o11st.FItemList(i).FoptionCnt %>:<%= o11st.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= o11st.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If o11st.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= o11st.FItemList(i).FinfoDiv %>
		<%
		If (o11st.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(o11st.FItemList(i).FlastErrStr) &"'>ERR:"& o11st.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% st11GoodNoArray = st11GoodNoArray & o11st.FItemList(i).FSt11GoodNo & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= st11GoodNoArray %></textarea>
	<button onclick="copyArr();">Copy</button>
	<script>
	function copyArr() {
		var tt = document.getElementById("itemidArr");
		tt.select();
		document.execCommand("copy");
	}
	</script>
<% End If %>
<tr height="20">
    <td colspan="17" align="center" bgcolor="#FFFFFF">
        <% if o11st.HasPreScroll then %>
		<a href="javascript:goPage('<%= o11st.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + o11st.StartScrollPage to o11st.FScrollCount + o11st.StartScrollPage - 1 %>
    		<% if i>o11st.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if o11st.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value= <%= page %>>
	<input type="hidden" name="research" value= <%= research %>>
	<input type="hidden" name="itemid" value= <%= itemid %>>
	<input type="hidden" name="makerid" value= <%= makerid %>>
	<input type="hidden" name="itemname" value= <%= itemname %>>
	<input type="hidden" name="bestOrd" value= <%= bestOrd %>>
	<input type="hidden" name="bestOrdMall" value= <%= bestOrdMall %>>
	<input type="hidden" name="sellyn" value= <%= sellyn %>>
	<input type="hidden" name="limityn" value= <%= limityn %>>
	<input type="hidden" name="sailyn" value= <%= sailyn %>>
	<input type="hidden" name="onlyValidMargin" value= <%= onlyValidMargin %>>
	<input type="hidden" name="startMargin" value= <%= startMargin %>>
	<input type="hidden" name="endMargin" value= <%= endMargin %>>
	<input type="hidden" name="isMadeHand" value= <%= isMadeHand %>>
	<input type="hidden" name="isOption" value= <%= isOption %>>
	<input type="hidden" name="infoDiv" value= <%= infoDiv %>>
	<input type="hidden" name="morningJY" value= <%= morningJY %>>
	<input type="hidden" name="extsellyn" value= <%= extsellyn %>>
	<input type="hidden" name="st11GoodNo" value= <%= st11GoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="st11Yes10x10No" value= <%= st11Yes10x10No %>>
	<input type="hidden" name="st11No10x10Yes" value= <%= st11No10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET o11st = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
