<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, coupangGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, morningJY, deliverytype, mwdiv, GosiEqual, MatchShipping, regedOptOver, exctrans
Dim expensive10x10, diffPrc, coupangYes10x10No, coupangNo10x10Yes, reqEdit, reqExpire, failCntExists, notinmakerid, notinitemid, priceOption, isSpecialPrice
Dim page, i, research, j, productId, kjypageSize
Dim oCoupang, splitMetaname, changMetaname, splitCoupangGosi, changeCoupangInfoDiv, isextusing, cisextusing, rctsellcnt
Dim startMargin, endMargin, scheduleNotInItemid, xl
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
coupangGoodNo			= request("coupangGoodNo")
productId				= request("productId")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchShipping			= request("MatchShipping")
regedOptOver			= request("regedOptOver")
GosiEqual				= request("GosiEqual")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
coupangYes10x10No		= request("coupangYes10x10No")
coupangNo10x10Yes		= request("coupangNo10x10Yes")
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
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
If kjypageSize = "" Then kjypageSize = 100
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchShipping = ""
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

'���� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If coupangGoodNo <> "" then
	Dim iA2, arrTemp2, arrcoupangGoodNo
	coupangGoodNo = replace(coupangGoodNo,",",chr(10))
	coupangGoodNo = replace(coupangGoodNo,chr(13),"")
	arrTemp2 = Split(coupangGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrcoupangGoodNo = arrcoupangGoodNo & "'"& trim(arrTemp2(iA2)) & "',"
			End If
		End If
		iA2 = iA2 + 1
	Loop
	coupangGoodNo = left(arrcoupangGoodNo,len(arrcoupangGoodNo)-1)
End If

'���� ���� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If productId <> "" then
	Dim iA3, arrTemp3, arrproductId
	productId = replace(productId,",",chr(10))
	productId = replace(productId,chr(13),"")
	arrTemp3 = Split(productId,chr(10))
	iA3 = 0
	Do While iA3 <= ubound(arrTemp3)
		If Trim(arrTemp3(iA3))<>"" then
			If Not(isNumeric(trim(arrTemp3(iA3)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp3(iA3) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrproductId = arrproductId & trim(arrTemp3(iA3)) & ","
			End If
		End If
		iA3 = iA3 + 1
	Loop
	productId = left(arrproductId,len(arrproductId)-1)
End If

Set oCoupang = new CCoupang
	oCoupang.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oCoupang.FPageSize					= kjypageSize
Else
	oCoupang.FPageSize					= 100
End If
	oCoupang.FRectCDL					= request("cdl")
	oCoupang.FRectCDM					= request("cdm")
	oCoupang.FRectCDS					= request("cds")
	oCoupang.FRectItemID				= itemid
	oCoupang.FRectItemName				= itemname
	oCoupang.FRectSellYn				= sellyn
	oCoupang.FRectLimitYn				= limityn
	oCoupang.FRectSailYn				= sailyn
'	oCoupang.FRectonlyValidMargin		= onlyValidMargin
	oCoupang.FRectStartMargin			= startMargin
	oCoupang.FRectEndMargin				= endMargin
	oCoupang.FRectMakerid				= makerid
	oCoupang.FRectCoupangGoodNo			= coupangGoodNo
	oCoupang.FRectProductId				= productId
	oCoupang.FRectMatchCate				= MatchCate
	oCoupang.FRectMatchShipping			= MatchShipping
	oCoupang.FRectregedOptOver			= regedOptOver
	oCoupang.FRectGosiEqual				= GosiEqual
	oCoupang.FRectIsMadeHand			= isMadeHand
	oCoupang.FRectIsOption				= isOption
	oCoupang.FRectIsReged				= isReged
	oCoupang.FRectNotinmakerid			= notinmakerid
	oCoupang.FRectNotinitemid			= notinitemid
	oCoupang.FRectExcTrans				= exctrans
	oCoupang.FRectPriceOption			= priceOption
	oCoupang.FRectIsSpecialPrice        = isSpecialPrice
	oCoupang.FRectDeliverytype			= deliverytype
	oCoupang.FRectMwdiv					= mwdiv
	oCoupang.FRectScheduleNotInItemid	= scheduleNotInItemid
	oCoupang.FRectIsextusing			= isextusing
	oCoupang.FRectCisextusing			= cisextusing
	oCoupang.FRectRctsellcnt			= rctsellcnt

	oCoupang.FRectExtNotReg				= ExtNotReg
	oCoupang.FRectExpensive10x10		= expensive10x10
	oCoupang.FRectdiffPrc				= diffPrc
	oCoupang.FRectCoupangYes10x10No		= coupangYes10x10No
	oCoupang.FRectCoupangNo10x10Yes		= coupangNo10x10Yes
	oCoupang.FRectExtSellYn				= extsellyn
	oCoupang.FRectInfoDiv				= infoDiv
	oCoupang.FRectFailCntOverExcept		= ""
	oCoupang.FRectFailCntExists			= failCntExists
	oCoupang.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oCoupang.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oCoupang.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oCoupang.getCoupangreqExpireItemList
Else
	oCoupang.getCoupangRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=coupangList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=coupang","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=coupang','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=coupang','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=coupang','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="coupangYes10x10No")&&(frm.coupangYes10x10No.checked)){ frm.coupangYes10x10No.checked=false }
	if ((comp.name!="coupangNo10x10Yes")&&(frm.coupangNo10x10Yes.checked)){ frm.coupangNo10x10Yes.checked=false }
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

    if ((comp.name=="coupangYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="coupangNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.coupangYes10x10No.checked){
            comp.form.coupangYes10x10No.checked = false;
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
	if ((comp.name!="coupangYes10x10No")&&(frm.coupangYes10x10No.checked)){ frm.coupangYes10x10No.checked=false }
	if ((comp.name!="coupangNo10x10Yes")&&(frm.coupangNo10x10Yes.checked)){ frm.coupangNo10x10Yes.checked=false }
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
//�귣���ڵ� �ù�� / ��ǰ���ڵ� ����
function pop_brandDeliver(){
	var pCM4 = window.open("/admin/etc/coupang/popCoupangBrandDeliveryList.asp","popbrandDelivergsshop","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM4.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/coupang/popCoupangCateList.asp","popCateCoupangmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//�����������
function popCouponList(){
	var popwin2=window.open('/admin/etc/coupang/popCoupangCouponCateList.asp','popCouponList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//Ű���� ����
function popKeywordItemList(){
	var popwin=window.open('/admin/etc/common/popKeywordList.asp?mallgubun=coupang','popKeywordItemList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_coupang.asp?itemid="+iitemid+'&mallid=coupang&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
// ���õ� ��ǰ ���
function CoupangSelectRegProcess() {
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

    if (confirm('Coupang�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ȸ
function CoupangSelectViewProcess() {
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

    if (confirm('Coupang�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
        document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ���� ����
function coupangSellYnProcess(chkYn) {
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
		case "X": strSell="����";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        if (chkYn=="X"){
			document.frmSvArr.cmdparam.value = "DELETE";
        }else{
        	document.frmSvArr.cmdparam.value = "EditSellYn";
        }
		document.frmSvArr.target = "xLink";
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ����
function CoupangSelectEditProcess() {
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

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function CoupangSelectPriceProcess() {
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

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��� ����
function CoupangSelectQuantityProcess() {
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

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ ��� �ϰ� ���� �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "QTY";
		document.frmSvArr.action = "<%=apiURL%>/outmall/coupang/actCoupangReq.asp"
		document.frmSvArr.submit();
    }
}
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
		<a href="https://wing.coupang.com" target="_blank">����Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") Then
				response.write "<font color='GREEN'>[ 10x10 | cube101010* ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		���� ��ǰ�ڵ� : <textarea rows="2" cols="20" name="coupangGoodNo" id="itemid"><%= replace(replace(coupangGoodNo,",",chr(10)), "'", "")%></textarea>
		&nbsp;
		���� �����ǰ�ڵ� : <textarea rows="2" cols="20" name="productId" id="itemid"><%=replace(productId,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >���� ��ϼ���_���δ��
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >���� ���۽õ� �� ����
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >���� �ݷ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >���� ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >���� ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>���� ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="coupangYes10x10No" <%= ChkIIF(coupangYes10x10No="on","checked","") %> ><font color=red>�����Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="coupangNo10x10Yes" <%= ChkIIF(coupangNo10x10Yes="on","checked","") %> ><font color=red>����ǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>

<p />

* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, ������ �ƴ� �� �ǸŰ� 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ɼǼ� 50�� �ʰ� ��ǰ, �ֹ����۹��� ��ǰ, �ɼǻ��� ����ġ(���ٿɼ�����:���޾��� or ���ٿɼǾ���:��������)<br />
* �ɼǼ� = �ٹ����� �Ǹ����� �ɼ� �� : ���޸� ��ϵ� �ɼ� ��(ǰ������)

<p />

<form name="frmReg" method="post" action="coupangItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="��� ���� ī�װ�" onclick="NotInCategory();">&nbsp;
				<input class="button" type="button" value="�����������" onclick="popCouponList();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
				<input class="button" type="button" value="Ű����" onclick="popKeywordItemList();">&nbsp;
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('coupang');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="�����" onclick="pop_brandDeliver();">&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="CoupangSelectRegProcess();">&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="����" onClick="CoupangSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnPriceSel" value="����" onClick="CoupangSelectPriceProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnQtySel" value="���" onClick="CoupangSelectQuantityProcess();">&nbsp;&nbsp;
				<br><br>
				���ο��� ��ȸ :
				<input class="button" type="button" id="btnViewSel" value="��ȸ" onClick="CoupangSelectViewProcess();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
					<option value="X">����</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="coupangSellYnProcess(frmReg.chgSellYn.value);">
			</td>
		</tr>
		</table>
    </td>
</tr>
</table>
</form>
<br>
<% End If %>
<!-- ����Ʈ ���� -->
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
		�˻���� : <b><%= FormatNumber(oCoupang.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCoupang.FTotalPage,0) %></b>
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
	<td width="140">Coupang�����<br>Coupang����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Coupang<br>���ݹ��Ǹ�</td>
	<td width="70">Coupang<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">�����<br>��Ī����</td>
	<td width="50">ǰ��</td>
	<td width="60">���ΰ��</td>
	<td width="150">Meta����</td>
</tr>
<% For i=0 to oCoupang.FResultCount - 1 %>
<tr align="center" <% If oCoupang.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oCoupang.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oCoupang.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oCoupang.FItemList(i).FItemID %>','coupang','<%=oCoupang.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oCoupang.FItemList(i).FItemID%>" target="_blank"><%= oCoupang.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oCoupang.FItemList(i).FCoupangStatcd <> 7 Then
	%>
		<br><%= oCoupang.FItemList(i).getCoupangStatName %>
	<%
			End If
			response.write oCoupang.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oCoupang.FItemList(i).FMakerid %> <%= oCoupang.FItemList(i).getDeliverytypeName %><br><%= oCoupang.FItemList(i).FItemName %></td>
	<td align="center"><%= oCoupang.FItemList(i).FRegdate %><br><%= oCoupang.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oCoupang.FItemList(i).FCoupangRegdate %><br><%= oCoupang.FItemList(i).FCoupangLastUpdate %></td>
	<td align="right">
		<% If oCoupang.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oCoupang.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oCoupang.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oCoupang.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).Fsellcash = 0 Then
		elseif (oCoupang.FItemList(i).FSaleYn="Y") Then
	%>
		<% if (oCoupang.FItemList(i).FOrgPrice<>0) then %>
		<strike><%= CLng(10000-oCoupang.FItemList(i).FOrgSuplycash/oCoupang.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
		<% end if %>
		<font color="#CC3333"><%= CLng(10000-oCoupang.FItemList(i).Fbuycash/oCoupang.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
	<%
		else
			response.write CLng(10000-oCoupang.FItemList(i).Fbuycash/oCoupang.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
	%>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).IsSoldOut Then
			If oCoupang.FItemList(i).FSellyn = "N" Then
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
		If oCoupang.FItemList(i).FItemdiv = "06" OR oCoupang.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oCoupang.FItemList(i).FCoupangStatCd > 0) Then
			If Not IsNULL(oCoupang.FItemList(i).FCoupangPrice) Then
				If (oCoupang.FItemList(i).Fsellcash <> oCoupang.FItemList(i).FCoupangPrice) Then
	%>
					<strong><%= formatNumber(oCoupang.FItemList(i).FCoupangPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oCoupang.FItemList(i).FCoupangPrice,0)&"<br>"
				End If

				If Not IsNULL(oCoupang.FItemList(i).FSpecialPrice) Then
					If (now() >= oCoupang.FItemList(i).FStartDate) And (now() <= oCoupang.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oCoupang.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oCoupang.FItemList(i).FSellyn="Y" and oCoupang.FItemList(i).FCoupangSellYn<>"Y") or (oCoupang.FItemList(i).FSellyn<>"Y" and oCoupang.FItemList(i).FCoupangSellYn="Y") Then
	%>
					<strong><%= oCoupang.FItemList(i).FCoupangSellYn %></strong>
	<%
				Else
					response.write oCoupang.FItemList(i).FCoupangSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oCoupang.FItemList(i).FCoupangGoodNo)) Then
			Response.Write oCoupang.FItemList(i).FCoupangGoodNo & "<br />"
		End If

		If Not(IsNULL(oCoupang.FItemList(i).FProductId)) Then
			Response.Write "<a target='_blank' href='http://www.coupang.com/vp/products/"&oCoupang.FItemList(i).FProductId&"?vendorItemId="&oCoupang.FItemList(i).FFirstVendorItemId&"'><font color='blue'>"&oCoupang.FItemList(i).FProductId&"</font></a>"
		End If
	%>
	</td>
	<td align="center"><%= oCoupang.FItemList(i).Freguserid %></td>
	<td align="center">
		<a href="javascript:popManageOptAddPrc('<%=oCoupang.FItemList(i).FItemID%>','0');"><%= oCoupang.FItemList(i).FoptionCnt %>:<%= oCoupang.FItemList(i).FregedOptCnt %></a>
	</td>
	<td align="center"><%= oCoupang.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oCoupang.FItemList(i).FOutboundShippingPlaceCode > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oCoupang.FItemList(i).FinfoDiv %>
		<%
		If (oCoupang.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oCoupang.FItemList(i).FlastErrStr) &"'>ERR:"& oCoupang.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
	<%
		changeCoupangInfoDiv = ""
		splitCoupangGosi = Split(oCoupang.FItemList(i).FMallinfoDiv, ",")
		For j = 0 to Ubound(splitCoupangGosi)
			rw oCoupang.FItemList(i).getCoupangInfoDiv(Trim(splitCoupangGosi(j)))
		Next
	%>
	</td>
	<td align="center">
	<%
		changMetaname = ""
		splitMetaname = Split(oCoupang.FItemList(i).FMetaOption, ",")
		For j = 0 to Ubound(splitMetaname)
			If instr(splitMetaname(j), "***") > 0 Then
				changMetaname = changMetaname & "<font color='red'>" & Replace(splitMetaname(j), "***", "") & "</font>,"
			Else
				changMetaname = changMetaname & splitMetaname(j) & ","
			End If
		Next
		If Right(changMetaname,1) = "," Then
			changMetaname = Left(changMetaname, Len(changMetaname) - 1)
		End If
		response.write changMetaname
	%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oCoupang.HasPreScroll then %>
		<a href="javascript:goPage('<%= oCoupang.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oCoupang.StartScrollPage to oCoupang.FScrollCount + oCoupang.StartScrollPage - 1 %>
    		<% if i>oCoupang.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oCoupang.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="page" value= <%= page %>>
	<input type="hidden" name="kjypageSize" value= <%= kjypageSize %>>
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
	<input type="hidden" name="coupangGoodNo" value= <%= coupangGoodNo %>>
	<input type="hidden" name="productId" value= <%= productId %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchShipping" value= <%= MatchShipping %>>
	<input type="hidden" name="regedOptOver" value= <%= regedOptOver %>>
	<input type="hidden" name="GosiEqual" value= <%= GosiEqual %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="coupangYes10x10No" value= <%= coupangYes10x10No %>>
	<input type="hidden" name="coupangNo10x10Yes" value= <%= coupangNo10x10Yes %>>
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
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oCoupang = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
