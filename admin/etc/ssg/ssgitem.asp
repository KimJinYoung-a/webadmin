<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ssg
' Hieditor : ������ ����
'            2022.09.27 �ѿ�� ����(��������)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgitemcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, ssgGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, setMargin, exctrans
Dim expensive10x10, diffPrc, ssgYes10x10No, ssgNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, deliverytype, mwdiv, isSpecialPrice
Dim page, i, research, isextusing, scheduleNotInItemid, cisextusing, rctsellcnt
Dim oSsg, xl, kjypageSize
dim startsell, stopsell
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
ssgGoodNo				= request("ssgGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
ssgYes10x10No			= request("ssgYes10x10No")
ssgNo10x10Yes			= request("ssgNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
notinmakerid			= request("notinmakerid")
priceOption				= request("priceOption")
isSpecialPrice          = request("isSpecialPrice")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
setMargin				= request("setMargin")
startsell				= requestCheckVar(request("startsell"), 1)
stopsell				= requestCheckVar(request("stopsell"), 1)
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

	if (stopsell = "Y") then
		'// �Ǹ����� ��� ��ǰ���
		ExtNotReg = "D"
		sellyn = "N"
		extsellyn = "Y"
		ssgYes10x10No = "on"
		onlyValidMargin = ""
	elseif (startsell = "Y") then
		'// �Ǹ���ȯ ��� ��ǰ���
		ExtNotReg = "D"
		sellyn = "Y"
		extsellyn = "N"
		onlyValidMargin = "Y"
		notinmakerid = "N"
		notinitemid = "N"
		ssgNo10x10Yes = "on"
	end if
End If

If (session("ssBctID")="kjy8517") Then
'	itemid = "1844673,1844658,1840365,1840349,1840348,1836595,1836260,1835212,1834869,1833193,1833156,1832693,1832692,1832454,1830488,1830124,1830113,1829131,1829130,1829100,1829057,1829056,1828945,1828944,1828856,1828354,1826836,1826687,1826686,1826467,1826466,1825608,1824761,1824444,1823959,1823958,1823795,1823794,1823637,1823537,1823290,1823288,1823287,1823286,1823282,1823280,1823278,1823257,1823256,1823255,1823254,1823253,1823251,1823250,1823249,1823248,1823247,1823246,1823245,1823244,1823243,1823242,1823241,1823236,1823232,1823230,1823229,1823228,1823227,1823226,1823225,1823224,1823223,1823222,1823221,1823220,1823219,1823218,1823217,1822369,1820674,1820596,1819187,1819186,1819185,1819184,1819139,1818606,1818605,1818604,1818603,1818602,1817236,1817167,1816479,1816409,1816408,1815096,1815062,1814656,1814579,1814578,1813143,1813131,1812395,1812394,1812393,1812392,1811522,1811482,1811480,1811456,1811455,1811454,1811453,1811452,1811451,1811450,1811449,1811448,1811447,1811446,1811445,1811442,1811441,1811440,1811439,1811423,1811422,1811420,1811139,1810710,1810701,1808667,1808666,1808665,1808305,1808304,1805638,1805637,1805636,1805635,1805634,1805633,1805632,1805631,1804494,1804493,1804492,1804490,1804489,1804478,1804477,1803160,1803159,1803158,1803157,1802751,1800434,1800421,1800420,1798878,1798877,1798876,1798875,1796468,1796466,1795719,1795713,1795504,1795458,1795352,1795221,1795111,1792876,1791947,1779732,1778064,1777398,1777396,1772554,1771028,1767757,1764925,1764875,1764874,1680141,1622014,1533355,1413087,1404481,1396558,1363606,1361058,1196372,1143455,1143452,1143451,1143450,1143447,1143445,1135310,1097515,958669"
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
	if trim(arrItemid)<>"" and not(isnull(trim(arrItemid))) then
	itemid = left(trim(arrItemid),len(trim(arrItemid))-1)
	end if
End If

'ssg ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If ssgGoodNo <> "" then
	Dim iA2, arrTemp2, arrssgGoodNo
	ssgGoodNo = replace(ssgGoodNo,",",chr(10))
	ssgGoodNo = replace(ssgGoodNo,chr(13),"")
	arrTemp2 = Split(ssgGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arrssgGoodNo = arrssgGoodNo& "'"& trim(arrTemp2(iA2)) & "',"
		End If
		iA2 = iA2 + 1
	Loop
	ssgGoodNo = left(arrssgGoodNo,len(arrssgGoodNo)-1)
End If

Set oSsg = new Cssg
	oSsg.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oSsg.FPageSize					= kjypageSize
Else
	oSsg.FPageSize					= 100
End If
	oSsg.FRectCDL					= request("cdl")
	oSsg.FRectCDM					= request("cdm")
	oSsg.FRectCDS					= request("cds")
	oSsg.FRectItemID				= itemid
	oSsg.FRectItemName				= itemname
	oSsg.FRectSellYn				= sellyn
	oSsg.FRectLimitYn				= limityn
	oSsg.FRectSailYn				= sailyn
'	oSsg.FRectonlyValidMargin		= onlyValidMargin
	oSsg.FRectStartMargin			= startMargin
	oSsg.FRectEndMargin				= endMargin
	oSsg.FRectMakerid				= makerid
	oSsg.FRectssgGoodNo				= ssgGoodNo
	oSsg.FRectMatchCate				= MatchCate
	oSsg.FRectIsMadeHand			= isMadeHand
	oSsg.FRectIsOption				= isOption
	oSsg.FRectIsReged				= isReged
	oSsg.FRectNotinmakerid			= notinmakerid
	oSsg.FRectNotinitemid			= notinitemid
	oSsg.FRectExcTrans				= exctrans
	oSsg.FRectPriceOption			= priceOption
	oSsg.FRectIsSpecialPrice        = isSpecialPrice
	oSsg.FRectDeliverytype			= deliverytype
	oSsg.FRectMwdiv					= mwdiv
	oSsg.FRectScheduleNotInItemid	= scheduleNotInItemid
	oSsg.FRectSetMargin				= setMargin
	oSsg.FRectIsextusing			= isextusing
	oSsg.FRectCisextusing			= cisextusing
	oSsg.FRectRctsellcnt			= rctsellcnt

	oSsg.FRectExtNotReg				= ExtNotReg
	oSsg.FRectExpensive10x10		= expensive10x10
	oSsg.FRectdiffPrc				= diffPrc
	oSsg.FRectssgYes10x10No			= ssgYes10x10No
	oSsg.FRectssgNo10x10Yes			= ssgNo10x10Yes
	oSsg.FRectExtSellYn				= extsellyn
	oSsg.FRectInfoDiv				= infoDiv
	oSsg.FRectFailCntOverExcept		= ""
	oSsg.FRectFailCntExists			= failCntExists
	oSsg.FRectReqEdit				= reqEdit
	oSsg.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oSsg.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oSsg.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oSsg.getssgreqExpireItemList
Else
	oSsg.getssgRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=ssgList"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=ssg","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=ssg','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=ssg','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//���� ���� ī�װ�
function popMarginCateList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginCateList.asp?mallid=ssg','popMarginCateList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//���� ���� ��ǰ
function popMarginItemList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginItemList.asp?mallid=ssg','popMarginItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//Ű���� ����
function popKeywordItemList(){
	var popwin=window.open('/admin/etc/common/popKeywordList.asp?mallgubun=ssg','popKeywordItemList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//Ű���� ����
function popSourceAreaList(){
	var popwin=window.open('/admin/etc/common/popSourceareaList.asp?mallgubun=ssg','popSourceAreaList','width=1300,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=ssg','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
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
	if ((comp.name!="ssgYes10x10No")&&(frm.ssgYes10x10No.checked)){ frm.ssgYes10x10No.checked=false }
	if ((comp.name!="ssgNo10x10Yes")&&(frm.ssgNo10x10Yes.checked)){ frm.ssgNo10x10Yes.checked=false }
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

    if ((comp.name=="ssgYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="ssgNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.ssgYes10x10No.checked){
            comp.form.ssgYes10x10No.checked = false;
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
	if ((comp.name!="ssgYes10x10No")&&(frm.ssgYes10x10No.checked)){ frm.ssgYes10x10No.checked=false }
	if ((comp.name!="ssgNo10x10Yes")&&(frm.ssgNo10x10Yes.checked)){ frm.ssgNo10x10Yes.checked=false }
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
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/ssg/popssgcateList.asp","popCateSSGmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//ǥ�� ī�װ� ����
function pop_stdCateManager() {
	var stdCM2 = window.open("/admin/etc/ssg/popssgStdcateList.asp","popstdCateSSGManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	stdCM2.focus();
}
//���� ī�װ� ����
function pop_dispCateManager() {
	var stdCM3 = window.open("/admin/etc/ssg/popssgdispcateList.asp","popstdCateSSGManager","width=1200,height=600,scrollbars=yes,resizable=yes");
	stdCM3.focus();
}

// ���õ� ��ǰ ���
function ssgREGProcess() {
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

    if (confirm('SSG�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "REG";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
//������ȸ
function ssgConfirmProcess(){
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

    if (confirm('SSG�� �����Ͻ� ' + chkSel + '�� ��ǰ ���ο��θ� �˻��Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function ssgSellYnProcess(chkYn) {
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
		case "N": strSell="�Ͻ�����";break;
		case "X": strSell="�����ߴ�";break;
	}

	if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSellYn";
        document.frmSvArr.chgSellYn.value = chkYn;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ ����
function ssgEditProcess(){
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

    if (confirm('SSG�� �����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
//��ǰ ��ȸ
function ssgViewItemProcess(){
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

    if (confirm('SSG�� �����Ͻ� ' + chkSel + '�� ��ȸ �Ͻðڽ��ϱ�?')){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "VIEW";
        document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
        document.frmSvArr.submit();
    }
}
function ssgGosiViewProcess(){
	if (confirm('SSG�� ��� ������ ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "GOSI";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
function ssgAreaViewProcess(){
	if (confirm('SSG�� ������ ������ ��ȸ �Ͻðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "AREA";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
function ssgDisplayCateProcess(){
	if (confirm('SSG ����ī�װ��� ���� ���ðڽ��ϱ�?')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "DISPCATE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
		document.frmSvArr.submit();
	}
}
//����, ��ȿ�� ��� �˾�
function popssgDate(iitemid){
    var pdate = window.open("/admin/etc/ssg/popssgDate.asp?itemid="+iitemid+'&mallid=ssg',"popssgDate","width=500,height=200,scrollbars=yes,resizable=yes");
	pdate.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=ssg&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}

<% if request("auto") = "Y" then %>
function ssgEditProcessAuto() {
	var cnt = <%= oSsg.FResultCount %>;
	var i, obj;
	for (i = 0; ; i++) {
		obj = document.getElementById('cksel' + i);
		if (obj == undefined) { break; }
		obj.checked = true;
	}
    document.frmSvArr.target = "xLink";
    document.frmSvArr.cmdparam.value = "EDIT";
	document.frmSvArr.auto.value = "Y";
    document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actssgReq.asp"
    document.frmSvArr.submit();
}

window.onload = function() {
	var cnt = <%= oSsg.FResultCount %>;
	if (cnt === 0) {
		// 45�е� ���ΰ�ħ
		setTimeout(function() {
			location.reload();
		}, 45*60*1000);
	} else {
		ssgEditProcessAuto();
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
		<a href="http://po.ssgadm.com/" target="_blank">SSG Admin�ٷΰ���</a>
		<%
			If ((session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") OR (session("ssBctID")="as2304") ) Then
				response.write "<font color='GREEN'>[ 0000003198 | Cube1010**! ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		SSG ��ǰ�ڵ� : <textarea rows="2" cols="20" name="ssgGoodNo" id="itemid"><%= replace(replace(ssgGoodNo,",",chr(10)), "'", "")%></textarea>
		&nbsp;
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >SSG ��ϼ���_���δ��
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >SSG ���۽õ� �� ����
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >SSG �ݷ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >SSG ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >SSG ��ϿϷ�(����)
		</select>&nbsp;
		<input type="radio" id="AR" name="isReged" <%= ChkIIF(isReged="A","checked","") %> onClick="checkisReged(this)" value="A">��ü</label>&nbsp;
		<label><input type="radio" id="NR" name="isReged" <%= ChkIIF(isReged="N","checked","") %> onClick="checkisReged(this)" value="N">�̵��<font color="<%= CHKIIF(makerid="" and itemid="", "#000000", "#AAAAAA") %>">(�ֱ� 3���� ��ϻ�ǰ��)</font></label>&nbsp;
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>SSG ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="ssgYes10x10No" <%= ChkIIF(ssgYes10x10No="on","checked","") %> ><font color=red>SSG�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="ssgNo10x10Yes" <%= ChkIIF(ssgNo10x10Yes="on","checked","") %> ><font color=red>SSGǰ��&�ٹ������ǸŰ���</font>(�������ܻ�ǰ ����) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<% if request("auto") <> "Y" then %>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����), �Һ��ڰ� ��� 80% �ʰ������� ��� 80% ���ΰ�<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, �ǸŰ�(���ΰ�) 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : <br />

<p />
<% end if %>
<form name="frmReg" method="post" action="ssgitem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="���Ը�������(ī�װ�)" onclick="popMarginCateList();">&nbsp;
				<input class="button" type="button" value="���Ը�������(��ǰ)" onclick="popMarginItemList();">&nbsp;
				<input class="button" type="button" value="Ű����" onclick="popKeywordItemList();">&nbsp;
				<input class="button" type="button" value="������" onclick="popSourceAreaList();">&nbsp;
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">

			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('ssg');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ǥ��ī�װ�" onclick="pop_stdCateManager();" style=color:blue;font-weight:bold> &nbsp;
				<input class="button" type="button" value="����ī�װ�" onclick="pop_dispCateManager();" style=color:blue;font-weight:bold> &nbsp;
				<!-- <input class="button" type="button" value="ī�װ�" onclick="pop_CateManager();"> &nbsp; -->
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
				<input class="button" type="button" id="btnREG" value="���" onClick="ssgREGProcess();" style=color:red;font-weight:bold>&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnOEdit" value="����" onClick="ssgEditProcess();" style=color:blue;font-weight:bold>
				<br><br>
				��Ÿ�ڵ� ��ȸ :
				<input class="button" type="button" id="btnStat" value="��ǰ��ȸ" onClick="ssgViewItemProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnStat" value="����" onClick="ssgConfirmProcess();">&nbsp;&nbsp;
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<input class="button" type="button" id="btnGosi" value="���" onClick="ssgGosiViewProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnGosi" value="������" onClick="ssgAreaViewProcess();">
				<br><br>
				ī�װ� ��ȸ :
				<input class="button" type="button" id="btnGosi" value="����" onClick="ssgDisplayCateProcess();">
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ�</option>
					<option value="X">�����ߴ�</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="ssgSellYnProcess(frmReg.chgSellYn.value);">
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
<input type="hidden" name="auto" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
		�˻���� : <b><%= FormatNumber(oSsg.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oSsg.FTotalPage,0) %></b>
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
	<td width="140">SSG�����<br>SSG����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">SSG<br>���ݹ��Ǹ�</td>
	<td width="100">SSG<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="50">���븶��</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
</tr>
<% For i=0 to oSsg.FResultCount - 1 %>
<tr align="center" <% If oSsg.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" id="cksel<%= i %>" onClick="AnCheckClick(this);"  value="<%= oSsg.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oSsg.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oSsg.FItemList(i).FItemID %>','ssg','<%=oSsg.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oSsg.FItemList(i).FItemID%>" target="_blank"><%= oSsg.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oSsg.FItemList(i).FssgStatcd <> 7 Then
	%>
		<br><%= oSsg.FItemList(i).getssgStatName %>
	<%
			End If
			response.write oSsg.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oSsg.FItemList(i).FMakerid %> <%= oSsg.FItemList(i).getDeliverytypeName %><br><%= oSsg.FItemList(i).FItemName %></td>
	<td align="center"><%= oSsg.FItemList(i).FRegdate %><br><%= oSsg.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oSsg.FItemList(i).FssgRegdate %><br><%= oSsg.FItemList(i).FssgLastUpdate %></td>
	<td align="right">
		<% If oSsg.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oSsg.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oSsg.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oSsg.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oSsg.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).FssgPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oSsg.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oSsg.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oSsg.FItemList(i).FOrgSuplycash/oSsg.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oSsg.FItemList(i).Fbuycash/oSsg.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oSsg.FItemList(i).IsSoldOut Then
			If oSsg.FItemList(i).FSellyn = "N" Then
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
		If oSsg.FItemList(i).FItemdiv = "06" OR oSsg.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oSsg.FItemList(i).FssgStatCd > 0) Then
			If Not IsNULL(oSsg.FItemList(i).FssgPrice) Then
				If (oSsg.FItemList(i).Mustprice <> oSsg.FItemList(i).FssgPrice) Then
	%>
					<strong><%= formatNumber(oSsg.FItemList(i).FssgPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oSsg.FItemList(i).FssgPrice,0)&"<br>"
				End If

				If Not IsNULL(oSsg.FItemList(i).FSpecialPrice) Then
					If (now() >= oSsg.FItemList(i).FStartDate) And (now() <= oSsg.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oSsg.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oSsg.FItemList(i).FSellyn="Y" and oSsg.FItemList(i).FssgSellYn<>"Y") or (oSsg.FItemList(i).FSellyn<>"Y" and oSsg.FItemList(i).FssgSellYn="Y") Then
	%>
					<strong><%= oSsg.FItemList(i).FssgSellYn %></strong>
	<%
				Else
					response.write oSsg.FItemList(i).FssgSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oSsg.FItemList(i).FssgGoodNo)) Then
			Response.Write "<a target='_blank' href='http://www.ssg.com/item/itemView.ssg?itemId="&oSsg.FItemList(i).FssgGoodNo&"'>"&oSsg.FItemList(i).FssgGoodNo&"</a>"
			'Response.Write "<span style='cursor:pointer;' onclick=window.open('http://www.ssg.com/item/itemView.ssg?itemid="&oSsg.FItemList(i).FssgGoodNo&"')>"&oSsg.FItemList(i).FssgGoodNo&"</span><br>"
		End If
	%>
	</td>
	<td align="center"><%= oSsg.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oSsg.FItemList(i).FItemID%>','0');"><%= oSsg.FItemList(i).FoptionCnt %>:<%= oSsg.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oSsg.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= oSsg.FItemList(i).FSetMargin %></td>
	<td align="center">
	<%
		If oSsg.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��(ī)"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�(ī)</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oSsg.FItemList(i).FinfoDiv %>
		<%
		If (oSsg.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oSsg.FItemList(i).FlastErrStr) &"'>ERR:"& oSsg.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if oSsg.HasPreScroll then %>
		<a href="javascript:goPage('<%= oSsg.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oSsg.StartScrollPage to oSsg.FScrollCount + oSsg.StartScrollPage - 1 %>
    		<% if i>oSsg.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oSsg.HasNextScroll then %>
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
	<input type="hidden" name="ssgGoodNo" value= <%= ssgGoodNo %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="ssgYes10x10No" value= <%= ssgYes10x10No %>>
	<input type="hidden" name="ssgNo10x10Yes" value= <%= ssgNo10x10Yes %>>
	<input type="hidden" name="reqEdit" value= <%= reqEdit %>>
	<input type="hidden" name="reqExpire" value= <%= reqExpire %>>
	<input type="hidden" name="failCntExists" value= <%= failCntExists %>>
	<input type="hidden" name="optAddPrcRegTypeNone" value= <%= optAddPrcRegTypeNone %>>
	<input type="hidden" name="notinmakerid" value= <%= notinmakerid %>>
	<input type="hidden" name="priceOption" value= <%= priceOption %>>
	<input type="hidden" name="isSpecialPrice" value= <%= isSpecialPrice %>>
	<input type="hidden" name="deliverytype" value= <%= deliverytype %>>
	<input type="hidden" name="mwdiv" value= <%= mwdiv %>>
	<input type="hidden" name="setMargin" value= <%= setMargin %>>
	<input type="hidden" name="startsell" value= <%= startsell %>>
	<input type="hidden" name="stopsell" value= <%= stopsell %>>
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% SET oSsg = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
