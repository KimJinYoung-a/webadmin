<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/interpark/interparkcls.asp"-->
<%
Server.ScriptTimeOut = 60 * 5		' 5��
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY, isSpecialPrice
Dim bestOrdMall, interparkPrdno, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, priceOption, deliverytype, mwdiv, exctrans
Dim expensive10x10, diffPrc, interparkYes10x10No, interparkNo10x10Yes, reqEdit, reqExpire, failCntExists, isextusing, cisextusing, rctsellcnt
Dim page, i, research, InterparkGoodNoArray, scheduleNotInItemid
Dim oInterpark, xl
Dim startMargin, endMargin
Dim sSdate, sEdate
Dim purchasetype

sSdate					= requestCheckVar(request("iSD"),10)
sEdate					= requestCheckVar(request("iED"),10)

If sSdate = "" Then sSdate = DateSerial(Year(Now()), Month(Now()), 1)
If sEdate = "" Then sEdate = Date()

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
interparkPrdno			= request("interparkPrdno")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
interparkYes10x10No		= request("interparkYes10x10No")
interparkNo10x10Yes		= request("interparkNo10x10Yes")
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
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid="1548380,1707616,1854914,1824546,1827550,1827550,1824546,1754245,932141,1428639,932141,1818430,1755622,1824546,419748,1622289"

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
'������ũ ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If interparkPrdno <> "" then
	Dim iA2, arrTemp2, arrinterparkPrdno
	interparkPrdno = replace(interparkPrdno,",",chr(10))
	interparkPrdno = replace(interparkPrdno,chr(13),"")
	arrTemp2 = Split(interparkPrdno,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrinterparkPrdno = arrinterparkPrdno & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	interparkPrdno = left(arrinterparkPrdno,len(arrinterparkPrdno)-1)
End If

Set oInterpark = new CInterpark


If (session("ssBctID")="kjy8517") Then
	oInterpark.FPageSize					= 100
Else
	oInterpark.FPageSize					= 50
End If
	oInterpark.FCurrPage					= page
	oInterpark.FRectMakerid					= makerid
	oInterpark.FRectItemID					= itemid
	oInterpark.FRectItemName				= itemname
	oInterpark.FRectInterparkPrdno			= interparkPrdno
	oInterpark.FRectCDL						= request("cdl")
	oInterpark.FRectCDM						= request("cdm")
	oInterpark.FRectCDS						= request("cds")
	oInterpark.FRectExtNotReg				= ExtNotReg
	oInterpark.FRectIsReged					= isReged
	oInterpark.FRectNotinmakerid			= notinmakerid
	oInterpark.FRectNotinitemid				= notinitemid
	oInterpark.FRectExcTrans				= exctrans
	oInterpark.FRectPriceOption				= priceOption
	oInterpark.FRectIsSpecialPrice     		= isSpecialPrice
	oInterpark.FRectDeliverytype			= deliverytype
	oInterpark.FRectMwdiv					= mwdiv
	oInterpark.FRectIsextusing				= isextusing
	oInterpark.FRectCisextusing				= cisextusing
	oInterpark.FRectRctsellcnt				= rctsellcnt

	oInterpark.FRectSellYn					= sellyn
	oInterpark.FRectLimitYn					= limityn
	oInterpark.FRectSailYn					= sailyn
'	oInterpark.FRectonlyValidMargin			= onlyValidMargin
	oInterpark.FRectStartMargin				= startMargin
	oInterpark.FRectEndMargin				= endMargin
	oInterpark.FRectIsMadeHand				= isMadeHand
	oInterpark.FRectIsOption				= isOption
	oInterpark.FRectInfoDiv					= infoDiv
	oInterpark.FRectExtSellYn				= extsellyn
	oInterpark.FRectFailCntExists			= failCntExists
	oInterpark.FRectMatchCate				= MatchCate
	oInterpark.FRectExpensive10x10			= expensive10x10
	oInterpark.FRectdiffPrc					= diffPrc
	oInterpark.FRectInterparklYes10x10No	= interparkYes10x10No
	oInterpark.FRectInterparkNo10x10Yes		= interparkNo10x10Yes
	oInterpark.FRectReqEdit					= reqEdit
	oInterpark.FRectScheduleNotInItemid		= scheduleNotInItemid
	oInterpark.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oInterpark.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oInterpark.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oInterpark.getInterParkreqExpireItemList
Else
	oInterpark.getInterParkRegedItemList			'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=interparkList"& replace(DATE(), "-", "") &"_xl.xls"
Else
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
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

function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
// ������� �귣��
function NotInMakerid(){
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=interpark","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=interpark','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=interpark','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function checkNDelItem(){
	alert('�������� ����..���ȹ�� �����븮���� ����');
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
	if ((comp.name!="interparkYes10x10No")&&(frm.interparkYes10x10No.checked)){ frm.interparkYes10x10No.checked=false }
	if ((comp.name!="interparkNo10x10Yes")&&(frm.interparkNo10x10Yes.checked)){ frm.interparkNo10x10Yes.checked=false }
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

    if ((comp.name=="interparkYes10x10No")&&(comp.checked)){
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
			comp.form.failCntExists.value = "N";
    	}
    }

    if ((comp.name=="interparkNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.interparkYes10x10No.checked){
            comp.form.interparkYes10x10No.checked = false;
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
	if ((comp.name!="interparkYes10x10No")&&(frm.interparkYes10x10No.checked)){ frm.interparkYes10x10No.checked=false }
	if ((comp.name!="interparkNo10x10Yes")&&(frm.interparkNo10x10Yes.checked)){ frm.interparkNo10x10Yes.checked=false }
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

//ī�װ� ����
// function pop_CateManager() {
// 	var pCM2 = window.open("/admin/etc/interpark/popinterparkcateList.asp","popinterparkcateList","width=1100,height=700,scrollbars=yes,resizable=yes");
// 	pCM2.focus();
// }

//ī�װ�New ����
function pop_NewCateManager() {
	var pCate = window.open("/admin/etc/interpark/popNewinterparkcateList.asp","popCateinterparkmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCate.focus();
}
function popItem2CategoryRedirect(itemid){
    var popwin = window.open('/admin/etc/interpark/InterParkMatcheDispCateByitemRedirect.asp?itemid=' + itemid,'MatcheDispCate','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=interpark&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function EditIParkSupplyCtrtSeq(iitemid){
    var popwin = window.open('/admin/etc/interpark/EditIParkSupplyCtrtSeq.asp?itemid=' + iitemid,'EditIParkSupplyCtrtSeq','width=800,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=interpark','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ���õ� ��ǰ �ϰ� ���
function interparkregIMSI(isreg) {
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
        if (confirm('������ũ�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
        }
    }else{
        if (confirm('������ũ�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnDelImsi").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
        }
    }
}
//���� ��ǰ ��ȸ
function interparkStatCheckProcess(){
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

    if (confirm('������ũ�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹ� ���¸� ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ ����
function interparkDeleteProcess(){
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
    if (confirm('API�� �����ϴ� ����� �ƴմϴ�.\n\n������ũ ���ο��� ������ ó���ؾ� �մϴ�.\n\n ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
		if (confirm('���� �����Ͻðڽ��ϱ�? Ȯ�ι�ư Ŭ���� DB���� ��ǰ�� �����˴ϴ�.')){
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DELETE";
			document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
			document.frmSvArr.submit();
		}
    }
}
function interparkCateProcess(){
	if (confirm('������ũ ī�װ��� ���� ���ðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";iSD
		document.frmSvArr.param1.value = $("#iSD").val();
		document.frmSvArr.param2.value = $("#iED").val();
		document.frmSvArr.cmdparam.value = "CATE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
	}
}
// ���õ� ��ǰ �Ǹſ��� ����
function InterParkSellYnProcess(chkYn) {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��������ũ���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� ������ũ���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ���
function interparkSelectRegProcess() {
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

    if (confirm('������ũ�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��������ũ���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ϰ� ����
function interparkEditProcess() {
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

    if (confirm('������ũ�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��������ũ�� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
		document.frmSvArr.submit();
    }
}

<% if request("auto") = "Y" then %>
function interparkEditProcessAuto() {
	var cnt = <%= oInterpark.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/interpark/actInterparkReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oInterpark.FResultCount %>;
	if (cnt === 0) {
		// 45�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		interparkEditProcessAuto();
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
		<a href="http://ipss.interpark.com/member/login.do?_method=initial&GNBLogin=Y&wid1=wgnb&wid2=wel_login&wid3=seller" target="_blank">(��)������ũAdmin�ٷΰ���</a>
		<a href="https://seller.interpark.com/login" target="_blank">(��)������ũAdmin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ coolhass | store10x10 ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		������ũ ��ǰ�ڵ� : <textarea rows="2" cols="20" name="interparkPrdno" id="itemid"><%=replace(interparkPrdno,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >������ũ ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >������ũ ��ϿϷ�(����)
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
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>������ũ ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="interparkYes10x10No" <%= ChkIIF(interparkYes10x10No="on","checked","") %> ><font color=red>������ũ�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="interparkNo10x10Yes" <%= ChkIIF(interparkNo10x10Yes="on","checked","") %> ><font color=red>������ũǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����), ������ �ø�ó��(������ũ�� ������ �Ⱦ�)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : <br />

<p />
<% end if %>
<form name="frmReg" method="post" action="interparkItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('interpark');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ī�װ�" onclick="pop_NewCateManager();" style=color:blue;font-weight:bold>
				<!--
				<input class="button" type="button" value="ī�װ�" onclick="pop_CateManager();">
				-->
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
				<input class="button" type="button" id="btnRegImsi" value="�������" onClick="interparkregIMSI(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnDelImsi" value="��������" onClick="interparkregIMSI(false);" >
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnRegSel" value="���õ��" onClick="interparkSelectRegProcess(true);">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="���ü���" onClick="interparkEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="������ȸ" onClick="interparkStatCheckProcess();">
			<% If (session("ssBctID")="kjy8517") Then %>
				&nbsp;&nbsp;<input class="button" type="button" id="btnODelete" value="��ǰ����" onClick="interparkDeleteProcess();" style=font-weight:bold>
			<% End If %>
				<br><br>
				ī�װ� ��ȸ :
				<input id="iSD" name="iSD" value="<%=sSdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				<input id="iED" name="iED" value="<%=sEdate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "iSD", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "iED", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
				<input class="button" type="button" id="btnCate" value="��ȸ" onClick="interparkCateProcess();">
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
					<option value="X">�Ǹ�����(����)</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="InterParkSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="param1">
<input type="hidden" name="param2">
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= FormatNumber(oInterpark.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oInterpark.FTotalPage,0) %></b>
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
	<td width="140">��ǰ���������<br>��ǰ����������</td>
	<td width="140">������ũ�����<br>������ũ����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">������ũ<br>���ݹ��Ǹ�</td>
	<!-- <td width="60">����</td> -->
	<td width="70">������ũ<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="100">ī�װ�<br>��Ī����</td>
	<td width="50">ǰ��</td>
</tr>
<% For i=0 to oInterpark.FResultCount - 1 %>
<tr align="center" <% If oInterpark.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oInterpark.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oInterpark.FItemList(i).Fsmallimage %>" width="50"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oInterpark.FItemList(i).FItemID%>" target="_blank"><%= oInterpark.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oInterpark.FItemList(i).FLimitYn="Y" Then
	%>
			<br><%= oInterpark.FItemList(i).getLimitHtmlStr %></font>
	<%
			End If
			response.write "<br>"&oInterpark.FItemList(i).getiParkRegStateName
		End If
	%>
	</td>
	<td align="left"><%= oInterpark.FItemList(i).FMakerid %><%= oInterpark.FItemList(i).getDeliverytypeName %><br><%= oInterpark.FItemList(i).FItemName %></td>
	<td align="center"><%= oInterpark.FItemList(i).FiparkTmpregdate %><br><%= oInterpark.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oInterpark.FItemList(i).FinterparkRegdate %><br><%= oInterpark.FItemList(i).FinterparkLastUpdate %></td>

	<td align="right">
	<% If oInterpark.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(oInterpark.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(oInterpark.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(oInterpark.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
		<%
		If oInterpark.FItemList(i).Fsellcash = 0 Then
		elseif (oInterpark.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oInterpark.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oInterpark.FItemList(i).FOrgSuplycash/oInterpark.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oInterpark.FItemList(i).Fbuycash/oInterpark.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oInterpark.FItemList(i).Fbuycash/oInterpark.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oInterpark.FItemList(i).IsSoldOut Then
			If oInterpark.FItemList(i).FSellyn = "N" Then
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
		If oInterpark.FItemList(i).FItemdiv = "06" OR oInterpark.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
    <% if Not IsNULL(oInterpark.FItemList(i).FmayiParkPrice) then %>
        <% if (oInterpark.FItemList(i).Fsellcash<>oInterpark.FItemList(i).FmayiParkPrice) then %>
        <strong><%= formatNumber(oInterpark.FItemList(i).FmayiParkPrice,0) %></strong>
        <% else %>
        <%= formatNumber(oInterpark.FItemList(i).FmayiParkPrice,0) %>
        <% end if %>
        <br>

		<%
			If Not IsNULL(oInterpark.FItemList(i).FSpecialPrice) Then
				If (now() >= oInterpark.FItemList(i).FStartDate) And (now() <= oInterpark.FItemList(i).FEndDate) Then
					response.write "<font color='orange'><strong>(Ư)" & formatNumber(oInterpark.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
				End If
			End If
		%>

        <% if (oInterpark.FItemList(i).FmayiParkSellYn="X") then %>
        <a href="javascript:checkNDelItem('<%= oInterpark.FItemList(i).FItemID %>')">
        <% end if %>

        <% if (oInterpark.FItemList(i).FSellyn<>oInterpark.FItemList(i).FmayiParkSellYn) then %>
        <strong><%= oInterpark.FItemList(i).FmayiParkSellYn %></strong>
        <% else %>
        <%= oInterpark.FItemList(i).FmayiParkSellYn %>
        <% end if %>

        <% if (oInterpark.FItemList(i).FmayiParkSellYn="X") then %>
        </a>
        <% end if %>
    <% end if %>
	</td>
<!--
	<td align="center"><a href="javascript:EditIParkSupplyCtrtSeq('oInterpark.FItemList(i).FItemID')"> oInterpark.FItemList(i).GetExtStoreSeqName ( oInterpark.FItemList(i).FSupplyCtrtSeq )</a></td>
-->
	<td align="center">
		<a target=_blank href="https://shopping.interpark.com/product/productInfo.do?prdNo=<%= oInterpark.FItemList(i).FInterparkPrdno %>"><%= oInterpark.FItemList(i).FInterparkPrdno %></a>
		<% If IsNULL(oInterpark.FItemList(i).FInterparkPrdno) Then %>
			<a href="javascript:DelTenIparkItem('<%= oInterpark.FItemList(i).FItemID %>')"><img src="/images/i_delete.gif" width="8" height="9" border="0"></a>
		<% End If %>
	</td>
	<td align="center"><%= oInterpark.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oInterpark.FItemList(i).FItemID%>','0');"><%= oInterpark.FItemList(i).FoptionCnt %>:<%= oInterpark.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oInterpark.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
'		If IsNULL(oInterpark.FItemList(i).FCatekey) then
'			<font color="darkred">��Ī�ȵ�</font><br>
'	 	Else
'			<a href="javascript:popItem2CategoryRedirect('oInterpark.FItemList(i).FItemID');">oInterpark.FItemList(i).Finterparkdispcategory</a>
'		End If

		If oInterpark.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If

		If (oInterpark.FItemList(i).FaccFailCNT > 0) Then
	    	response.write "<br><font color='red' title='" & oInterpark.FItemList(i).FlastErrStr & "'>ERR: " & oInterpark.FItemList(i).FaccFailCNT & "</font>"
		End If
	%>
	</td>
    <td align="center"><%= oInterpark.FItemList(i).FinfoDiv %></td>
</tr>
<% InterparkGoodNoArray = InterparkGoodNoArray & oInterpark.FItemList(i).FInterparkPrdno & VBCRLF %>
<% Next %>
<% If (session("ssBctID")="kjy8517") Then %>
	<textarea id="itemidArr"><%= InterparkGoodNoArray %></textarea>
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
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oInterpark.HasPreScroll then %>
		<a href="javascript:goPage('<%= oInterpark.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oInterpark.StartScrollPage to oInterpark.FScrollCount + oInterpark.StartScrollPage - 1 %>
    		<% if i>oInterpark.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oInterpark.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
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
	<input type="hidden" name="interparkPrdno" value= <%= interparkPrdno %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="interparkYes10x10No" value= <%= interparkYes10x10No %>>
	<input type="hidden" name="interparkNo10x10Yes" value= <%= interparkNo10x10Yes %>>
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
<% SET oInterpark = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
