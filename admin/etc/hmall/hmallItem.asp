<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/hmall/hmallCls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv
Dim bestOrdMall, hmallGoodNo, hmallGoodNo2, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, notinitemid, morningJY, deliverytype, mwdiv, exctrans, MatchIMG
Dim expensive10x10, diffPrc, hmallYes10x10No, hmallNo10x10Yes, reqEdit, reqExpire, failCntExists, priceOption, isSpecialPrice, setMargin, isextusing, scheduleNotInItemid
Dim page, i, research, cisextusing, rctsellcnt
Dim oHmall, xl
Dim startMargin, endMargin
Dim purchasetype
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
hmallGoodNo				= request("hmallGoodNo")
hmallGoodNo2			= request("hmallGoodNo2")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
MatchIMG				= request("MatchIMG")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
hmallYes10x10No			= request("hmallYes10x10No")
hmallNo10x10Yes			= request("hmallNo10x10Yes")
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
notinitemid				= requestCheckVar(request("notinitemid"), 1)
exctrans				= requestCheckVar(request("exctrans"), 1)
scheduleNotInItemid		= requestCheckVar(request("scheduleNotInItemid"), 1)
isextusing				= requestCheckVar(request("isextusing"), 1)
cisextusing				= requestCheckVar(request("cisextusing"), 1)
rctsellcnt				= requestCheckVar(request("rctsellcnt"), 1)
purchasetype			= request("purchasetype")
xl 						= request("xl")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchIMG = ""
	onlyValidMargin = "Y"
	bestOrd = "on"
	sellyn = "Y"
End If

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	MatchIMG = ""
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
'hmall ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If hmallGoodNo2 <> "" then
	Dim iA2, arrTemp2, arrhmallGoodNo2
	hmallGoodNo2 = replace(hmallGoodNo2,",",chr(10))
	hmallGoodNo2 = replace(hmallGoodNo2,chr(13),"")
	arrTemp2 = Split(hmallGoodNo2,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrhmallGoodNo2 = arrhmallGoodNo2 & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	hmallGoodNo2 = left(arrhmallGoodNo2,len(arrhmallGoodNo2)-1)
End If

Set oHmall = new CHmall
	oHmall.FCurrPage					= page
	oHmall.FPageSize					= 100
	oHmall.FRectCDL						= request("cdl")
	oHmall.FRectCDM						= request("cdm")
	oHmall.FRectCDS						= request("cds")
	oHmall.FRectItemID					= itemid
	oHmall.FRectItemName				= itemname
	oHmall.FRectSellYn					= sellyn
	oHmall.FRectLimitYn					= limityn
	oHmall.FRectSailYn					= sailyn
'	oHmall.FRectonlyValidMargin			= onlyValidMargin
	oHmall.FRectStartMargin				= startMargin
	oHmall.FRectEndMargin				= endMargin
	oHmall.FRectMakerid					= makerid
	oHmall.FRectHmallGoodNo				= hmallGoodNo2
	oHmall.FRectMatchCate				= MatchCate
	oHmall.FRectMatchIMG				= MatchIMG
	oHmall.FRectIsMadeHand				= isMadeHand
	oHmall.FRectIsOption				= isOption
	oHmall.FRectIsReged					= isReged
	oHmall.FRectNotinmakerid			= notinmakerid
	oHmall.FRectNotinitemid				= notinitemid
	oHmall.FRectExcTrans				= exctrans
	oHmall.FRectPriceOption				= priceOption
	oHmall.FRectIsSpecialPrice			= isSpecialPrice
	oHmall.FRectDeliverytype			= deliverytype
	oHmall.FRectMwdiv					= mwdiv
	oHmall.FRectSetMargin				= setMargin
	oHmall.FRectScheduleNotInItemid		= scheduleNotInItemid
	oHmall.FRectIsextusing				= isextusing
	oHmall.FRectCisextusing				= cisextusing
	oHmall.FRectRctsellcnt				= rctsellcnt

	oHmall.FRectExtNotReg				= ExtNotReg
	oHmall.FRectExpensive10x10			= expensive10x10
	oHmall.FRectdiffPrc					= diffPrc
	oHmall.FRectHmallYes10x10No			= hmallYes10x10No
	oHmall.FRectHmallNo10x10Yes			= hmallNo10x10Yes
	oHmall.FRectExtSellYn				= extsellyn
	oHmall.FRectInfoDiv					= infoDiv
	oHmall.FRectFailCntOverExcept		= ""
	oHmall.FRectFailCntExists			= failCntExists
	oHmall.FRectReqEdit					= reqEdit
	oHmall.FRectPurchasetype			= purchasetype
If (bestOrd = "on") Then
    oHmall.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oHmall.FRectOrdType = "BM"
End If


If isReged = "R" Then					'ǰ��ó����� ��ǰ���� ����Ʈ
	oHmall.getHmallreqExpireItemList
Else
	oHmall.getHmallRegedItemList		'�� �� ����Ʈ
End If

If (xl = "Y") Then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=hmall1010List"& replace(DATE(), "-", "") &"_xl.xls"
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
    var popwin = window.open("/admin/etc/outmall/popExtUse_Not_In_Makerid.asp?mallgubun=hmall1010","popNotInMakerid","width=1200,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=hmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ī�װ�
function NotInCategory(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Category.asp?mallgubun=hmall1010','notinCategory','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//���� ���� ī�װ�
function popMarginCateList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginCateList.asp?mallid=hmall1010','popMarginCateList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
//���� ���� ��ǰ
function popMarginItemList(){
	var popwin2=window.open('/admin/etc/ssg/popSsgMarginItemList.asp?mallid=hmall1010','popMarginItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin2.focus();
}
// ������ ���� ��ǰ
function NotInScheItemid(){
	var popwin=window.open('/admin/etc/schedule_Not_In_Itemid.asp?mallgubun=hmall1010','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallCateList.asp","popCateHmallmanager","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}

//���ø���
function pop_SectId() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallSectList.asp","popSectId","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
//�����Ī��
function pop_SectId2() {
	var pCM2 = window.open("/admin/etc/hmall/popHmallSectList2.asp","popSectId2","width=1200,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
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
	if ((comp.name!="hmallYes10x10No")&&(frm.hmallYes10x10No.checked)){ frm.hmallYes10x10No.checked=false }
	if ((comp.name!="hmallNo10x10Yes")&&(frm.hmallNo10x10Yes.checked)){ frm.hmallNo10x10Yes.checked=false }
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

    if ((comp.name=="hmallYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="hmallNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.hmallYes10x10No.checked){
            comp.form.hmallYes10x10No.checked = false;
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
	if ((comp.name!="hmallYes10x10No")&&(frm.hmallYes10x10No.checked)){ frm.hmallYes10x10No.checked=false }
	if ((comp.name!="hmallNo10x10Yes")&&(frm.hmallNo10x10Yes.checked)){ frm.hmallNo10x10Yes.checked=false }
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

//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet_hmall.asp?itemid="+iitemid+'&mallid=hmall1010&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// ���õ� ��ǰ �ϰ� ���
function hmallSelectRegProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnRegSel").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REG";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� �̹��� ���
function hmallSelectImagesProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� �̹����� ��� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "IMAGE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���	''2018-12-10 ������ �߰�
function hmallSelectRegItemProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ��� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGAddItem";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� �̹��� ���	''2018-12-10 ������ �߰�
function hmallSelectImageProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� �̹����� ��� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGImage";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���	''2018-12-10 ������ �߰�
function hmallSelectImageConfirmProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� �̹����� Ȯ�� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "REGImageConfirm";
		document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ���� ����
function hmallSellYnProcess(chkYn) {
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
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//����
function hmallSelectEditProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//���� ''2018-12-10 ������ �߰�
function hmallSelectEditItemProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ���� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDITItem";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//���� ����
function hmallSelectPriceEditProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "PRICE";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//�ɼ� ����
function hmallSelectOptionEditProcess() {
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

    if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ɼ��� ���� �Ͻðڽ��ϱ�?\n\n��HMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSelOption").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTEDIT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ �� ��ȸ
function hmallSelectViewProcess() {
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

	if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ �� ��ȸ
function hmallSelectViewProcess2() {
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

	if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��ǰ��ȸ �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
        document.frmSvArr.submit();
    }
}

//��ǰ ��� ��ȸ
function hmallSelectOptionViewProcess() {
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

	if (confirm('HMall�� �����Ͻ� ' + chkSel + '�� ��� ��ȸ �Ͻðڽ��ϱ�?')){
        //document.getElementById("btnOptViewSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "OPTSTAT";
        document.frmSvArr.action = "<%=apiURL%>/outmall/hmall/acthmallReq.asp"
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
		https://partner.hmall.com/
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then
				response.write "<font color='GREEN'>[ hs0027011 | cube101010 ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		hmall ��ǰ�ڵ� : <textarea rows="2" cols="20" name="hmallGoodNo2" id="itemid"><%=replace(hmallGoodNo2,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >Hmall ��Ͻ���
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >Hmall ��Ͽ���
			<option value="E" <%= CHkIIF(ExtNotReg="E","selected","") %> >Hmall ���
			<option value="C" <%= CHkIIF(ExtNotReg="C","selected","") %> >Hmall �ݷ�
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >Hmall ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >Hmall ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Hmall ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="hmallYes10x10No" <%= ChkIIF(hmallYes10x10No="on","checked","") %> ><font color=red>Hmall�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="hmallNo10x10Yes" <%= ChkIIF(hmallNo10x10Yes="on","checked","") %> ><font color=red>Hmallǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p />

* ���ظ��� : �����ǸŰ� ��� ���԰�, ������ �ݿø���<br />
* �����ǸŰ� : ���ΰ�(���ظ��� �̸��� ��� ����)<br />
* �������ܻ�ǰ1 : ������ܺ귣��, ������ܻ�ǰ, ���޸�������, ��ü����, ����ǰ, �ɹ��, ȭ�����, Ƽ��(����) ��ǰ, ������ �ƴ� �� �ǸŰ� 1���� �̸�, �������5�� ����, �ɼǺ�������� ���� 5�� ����<br />
* �������ܻ�ǰ2 : �ֹ����۹��� ��ǰ, �ɼ��߰��ݾ� �ִ� ��ǰ<br />

<p />
<form name="frmReg" method="post" action="hmallItem.asp" style="margin:0px;">
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
				<input class="button" type="button" value="������ ���� ��ǰ" onclick="NotInScheItemid();">
			</td>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('hmall1010');">&nbsp;&nbsp;
				<font color="RED">���� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="ī�װ�" onclick="pop_CateManager();">

				<input class="button" type="button" value="���ø���" onclick="pop_SectId();">
				<input class="button" type="button" value="�����Ī��" onclick="pop_SectId2();">
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="hmallSelectRegProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnImages" value="�̹���" onClick="hmallSelectImagesProcess();">
				<% If (session("ssBctID")="kjy8517") Then %>
			<!--
					&nbsp;&nbsp;<input class="button" type="button" id="btnRegItem" value="��ǰ" onClick="hmallSelectRegItemProcess();">
					&nbsp;&nbsp;<input class="button" type="button" id="btnImage" value="�̹������" onClick="hmallSelectImageProcess();">
					&nbsp;&nbsp;<input class="button" type="button" id="btnImageConfirm" value="�̹���Ȯ��" onClick="hmallSelectImageConfirmProcess();">
			-->
				<% End If %>
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEditSel" value="����" onClick="hmallSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="����" onClick="hmallSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelOption" value="�ɼ�" onClick="hmallSelectOptionEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnOptViewSel" value="�����ȸ" onClick="hmallSelectOptionViewProcess();">&nbsp;&nbsp;
				<% If (session("ssBctID")="kjy8517") Then %>
			<!--
					<input class="button" type="button" id="btnEditItem" value="��ǰ����" onClick="hmallSelectEditItemProcess();">
			-->
				<% End If %>
				<br><br>
				���ο��� ��ǰ :
				<!-- 
				<input class="button" type="button" id="btnViewSel" value="����ȸ" onClick="hmallSelectViewProcess();">&nbsp;&nbsp;
				--> 
				<input class="button" type="button" id="btnViewSel" value="����ȸ2" onClick="hmallSelectViewProcess2();">&nbsp;&nbsp;
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">ǰ��</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="hmallSellYnProcess(frmReg.chgSellYn.value);">
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		�˻���� : <b><%= FormatNumber(oHmall.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oHmall.FTotalPage,0) %></b>
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
	<td width="140">Hmall�����<br>Hmall����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">Hmall<br>���ݹ��Ǹ�</td>
	<td width="70">Hmall<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="50">���븶��</td>
	<td width="80">��Ī����</td>
	<td width="80">ǰ��</td>
	<td width="100">�̹���<br />��� | Ȯ��</td>
</tr>
<% For i=0 to oHmall.FResultCount - 1 %>
<tr align="center" <% If oHmall.FItemList(i).FNotSchIdx <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oHmall.FItemList(i).FItemID %>"></td>
<% If (xl <> "Y") Then %>
	<td><img src="<%= oHmall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= oHmall.FItemList(i).FItemID %>','hmall1010','<%=oHmall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
<% End If %>
	<td align="center">
		<a href="<%=vwwwUrl%>/<%=oHmall.FItemList(i).FItemID%>" target="_blank"><%= oHmall.FItemList(i).FItemID %></a>
	<%
		If (xl <> "Y") Then
			If oHmall.FItemList(i).FHmallStatcd <> 7 Then
	%>
		<br><%= oHmall.FItemList(i).getHmallStatName %>
	<%
			End If
			response.write oHmall.FItemList(i).getLimitHtmlStr
		End If
	%>
	</td>
	<td align="left"><%= oHmall.FItemList(i).FMakerid %> <%= oHmall.FItemList(i).getDeliverytypeName %><br><%= oHmall.FItemList(i).FItemName %></td>
	<td align="center"><%= oHmall.FItemList(i).FRegdate %><br><%= oHmall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oHmall.FItemList(i).FHmallRegdate %><br><%= oHmall.FItemList(i).FHmallLastUpdate %></td>
	<td align="right">
		<% If oHmall.FItemList(i).FSaleYn = "Y" Then %>
			<strike><%= FormatNumber(oHmall.FItemList(i).FOrgPrice,0) %></strike><br>
			<font color="#CC3333"><%= FormatNumber(oHmall.FItemList(i).FSellcash,0) %></font>
		<% Else %>
			<%= FormatNumber(oHmall.FItemList(i).FSellcash,0) %>
		<% End If %>
	</td>
	<td align="center">
		<%
		If oHmall.FItemList(i).Fsellcash = 0 Then
		%>
		' <strike><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></strike><br>
		' <font color="#CC3333"><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).FHmallPrice*100*100)/100 & "%" %></font>
		' <%
		' 	else
		' 		response.write CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		' 	end if
		elseif (oHmall.FItemList(i).FSaleYn="Y") Then
		%>
			<% if (oHmall.FItemList(i).FOrgPrice<>0) then %>
			<strike><%= CLng(10000-oHmall.FItemList(i).FOrgSuplycash/oHmall.FItemList(i).FOrgPrice*100*100)/100 & "%" %></strike><br>
			<% end if %>
			<font color="#CC3333"><%= CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%" %></font>
		<%
		else
			response.write CLng(10000-oHmall.FItemList(i).Fbuycash/oHmall.FItemList(i).Fsellcash*100*100)/100 & "%"
		end if
		%>
	</td>
	<td align="center">
	<%
		If oHmall.FItemList(i).IsSoldOut Then
			If oHmall.FItemList(i).FSellyn = "N" Then
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
		If oHmall.FItemList(i).FItemdiv = "06" OR oHmall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oHmall.FItemList(i).FHmallStatCd > 0) Then
			If Not IsNULL(oHmall.FItemList(i).FHmallPrice) Then
				If (oHmall.FItemList(i).Mustprice <> oHmall.FItemList(i).FHmallPrice) Then
	%>
					<strong><%= formatNumber(oHmall.FItemList(i).FHmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(oHmall.FItemList(i).FHmallPrice,0)&"<br>"
				End If

				If Not IsNULL(oHmall.FItemList(i).FSpecialPrice) Then
					If (now() >= oHmall.FItemList(i).FStartDate) And (now() <= oHmall.FItemList(i).FEndDate) Then
						response.write "<font color='orange'><strong>(Ư)" & formatNumber(oHmall.FItemList(i).FSpecialPrice,0)&"</strong></font><br />"
					End If
				End If

				If (oHmall.FItemList(i).FSellyn="Y" and oHmall.FItemList(i).FHmallSellYn<>"Y") or (oHmall.FItemList(i).FSellyn<>"Y" and oHmall.FItemList(i).FHmallSellYn="Y") Then
	%>
					<strong><%= oHmall.FItemList(i).FHmallSellYn %></strong>
	<%
				Else
					response.write oHmall.FItemList(i).FHmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
		<% If oHmall.FItemList(i).FHmallGoodNo <> "" Then %>
			<a target="_blank" href="https://www.hmall.com/pd/pda/itemPtc?slitmCd=<%=oHmall.FItemList(i).FHmallGoodNo%>"><%=oHmall.FItemList(i).FHmallGoodNo2%></a>
		<% End If %>
	</td>
	<td align="center"><%= oHmall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oHmall.FItemList(i).FItemID%>','0');"><%= oHmall.FItemList(i).FoptionCnt %>:<%= oHmall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oHmall.FItemList(i).FrctSellCNT %></td>
	<td align="center"><%= oHmall.FItemList(i).FSetMargin %></td>
	<td align="center">
	<%
		If oHmall.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
		<%= oHmall.FItemList(i).FinfoDiv %>
		<%
		If (oHmall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oHmall.FItemList(i).FlastErrStr) &"'>ERR:"& oHmall.FItemList(i).FaccFailCNT &"</font>"
		End If
		%>
	</td>
	<td align="center">
		<%= Chkiif(oHmall.FItemList(i).FAPIaddImg="Y","<font color='BLUE'>"&oHmall.FItemList(i).FAPIaddImg&"</font>", "<font color='RED'>"&oHmall.FItemList(i).FAPIaddImg&"</font>") %>&nbsp;|
		<%= Chkiif(oHmall.FItemList(i).FAPIconfirmImg="Y","<font color='BLUE'>"&oHmall.FItemList(i).FAPIconfirmImg&"</font>", "<font color='RED'>"&oHmall.FItemList(i).FAPIconfirmImg&"</font>") %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oHmall.HasPreScroll then %>
		<a href="javascript:goPage('<%= oHmall.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oHmall.StartScrollPage to oHmall.FScrollCount + oHmall.StartScrollPage - 1 %>
    		<% if i>oHmall.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oHmall.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
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
	<input type="hidden" name="hmallGoodNo" value= <%= hmallGoodNo %>>
	<input type="hidden" name="hmallGoodNo2" value= <%= hmallGoodNo2 %>>
	<input type="hidden" name="ExtNotReg" value= <%= ExtNotReg %>>
	<input type="hidden" name="isReged" value= <%= isReged %>>
	<input type="hidden" name="MatchCate" value= <%= MatchCate %>>
	<input type="hidden" name="MatchIMG" value= <%= MatchIMG %>>
	<input type="hidden" name="expensive10x10" value= <%= expensive10x10 %>>
	<input type="hidden" name="diffPrc" value= <%= diffPrc %>>
	<input type="hidden" name="hmallYes10x10No" value= <%= hmallYes10x10No %>>
	<input type="hidden" name="hmallNo10x10Yes" value= <%= hmallNo10x10Yes %>>
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
	<input type="hidden" name="notinitemid" value= <%= notinitemid %>>
	<input type="hidden" name="scheduleNotInItemid" value= <%= scheduleNotInItemid %>>
	<input type="hidden" name="exctrans" value= <%= exctrans %>>
	<input type="hidden" name="isextusing" value= <%= isextusing %>>
	<input type="hidden" name="cisextusing" value= <%= cisextusing %>>
	<input type="hidden" name="cdl" value= <%= request("cdl") %>>
	<input type="hidden" name="cdm" value= <%= request("cdm") %>>
	<input type="hidden" name="cds" value= <%= request("cds") %>>
</form>
<% Set oHmall = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->