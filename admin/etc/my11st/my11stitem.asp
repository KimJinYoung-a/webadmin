<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/my11st/my11stcls.asp"-->
<%
Dim makerid, itemid, itemname, bestOrd, sellyn, limityn, sailyn, onlyValidMargin, isMadeHand, isOption, infoDiv, morningJY
Dim bestOrdMall, my11stGoodNo, extsellyn, ExtNotReg, isReged, MatchCate, optAddPrcRegTypeNone, notinmakerid, deliverytype, mwdiv
Dim expensive10x10, diffPrc, my11stYes10x10No, my11stNo10x10Yes, reqEdit, reqExpire, failCntExists
Dim page, i, research, exctrans
Dim oMy11st

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
my11stGoodNo			= request("my11stGoodNo")
ExtNotReg				= request("ExtNotReg")
isReged					= request("isReged")
MatchCate				= request("MatchCate")
expensive10x10			= request("expensive10x10")
diffPrc					= request("diffPrc")
my11stYes10x10No		= request("my11stYes10x10No")
my11stNo10x10Yes		= request("my11stNo10x10Yes")
reqEdit					= request("reqEdit")
reqExpire				= request("reqExpire")
failCntExists			= request("failCntExists")
optAddPrcRegTypeNone	= request("optAddPrcRegTypeNone")
deliverytype			= request("deliverytype")
mwdiv					= request("mwdiv")
exctrans				= requestCheckVar(request("exctrans"), 1)

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"
''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "D"
	MatchCate = ""
	bestOrd = "on"
	sellyn = "Y"
End If

If (session("ssBctID")="kjy8517") Then
'	itemid="25413,37537,38164,41006,43212,43213,44845,45718,60874,60880,64049,64050,64052,64054,64056,64514,65043,68174,71953,72108,72213,82176,82181,82184,82200,86265,89707,91413,98444,103028,104254,104260,104261,104265,112056,113192,115581,118346,120595,125365,125668,128783,128810,133036,133037,133685,137641,141122,142201,168056,168058,171033,171034,171036,176545,176556,176557,176623,186097,190419,205572,205573,208978,218571,227970,231538,234625,235439,235451,235452,235454,236962,236964,238006,238951,242623,243393,243394,246215,256335,258957,259555,259560,263243,272056,273802,283178,287785,287791,295365,303507,308139,309833,313255,315545,322163,323363,324540,324583,335545,336672,338335,339042,339678,349077,354391,357649,357650,366680,366686,368605,372915,377330,379555,380096,380097,380098,380099,387208,401047,440282,443152,443153,445558,458134,458969,459634,473849,480566,480567,482914,501831,511083,513664,517653,518847,525978,526943,528874,528910,548313,548314,548315,554874,556276,561587,565121,586616,589447,595775,620888,623572,623573,635152,639778,641473,641492,646758,654929,662850,674230,683737,688245,693231,698159,707732,707734,707854,724957,724969,726965,739614,746773,768343,768344,788776,792735,800183,804712,809047,813285,820044,822619,822648,824356,846137,849637,855394,862860,865503,886399,893602,896360,897652,898296,901895,904786,905019,905354,910047,920312,935881,937633,939001,939003,939004,939005,939007,939008,944937,945725,951143,951144,951145,960030,963386,967680,969601,969635,969636,970269,970288,996229,999954,999957,999959,1007013,1009539,1009540,1009541,1014574,1024590,1027524,1043157,1043159,1047268,1047275,1052165,1054553,1056660,1068190,1071304,1071309,1071310,1071311,1071404,1075254,1075745,1088037,1088073,1097062,1097530,1097540,1103426,1103489,1103524,1105335,1105856,1106488,1115015,1121238,1123693,1133673,1140399,1159848,1159853,1167545,1168410,1172031,1172035,1176396,1185473,1193519,1196348,1196350,1196408,1204614,1215160,1215861,1217388,1234713,1240835,1240854,1255389,1255390,1259852,1272590,1274885,1274886,1283028,1285357,1285358,1286684,1287747,1290789,1290791,1293039,1293045,1295907,1295908,1295918,1304358,1308116,1314099,1333952,1333988,1336060,1345854,1354335,1361393,1361394,1371122,1381148,1387369,1408275,1417383,1417608,1419418,1428543,1428552,1431356,1434911,1438034,1441088,1452878,1456459,1456948,1456951,1456952,1463137,1464643,1477002,1477008,1481648,1484115,1485872,1485873,1489493,1490328,1493912,1500811,1514009,1540225,1543558,1548837,1588838,1589009,1589010,1594211,1598258,1598552,1612793,1615313,1617211,1619449,1619828,1620859,1640190"
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
'11st ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If my11stGoodNo <> "" then
	Dim iA2, arrTemp2, arrmy11stGoodNo
	my11stGoodNo = replace(my11stGoodNo,",",chr(10))
	my11stGoodNo = replace(my11stGoodNo,chr(13),"")
	arrTemp2 = Split(my11stGoodNo,chr(10))
	iA2 = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			If Not(isNumeric(trim(arrTemp2(iA2)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp2(iA2) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrmy11stGoodNo = arrmy11stGoodNo & trim(arrTemp2(iA2)) & ","
			End If
		End If
		iA2 = iA2 + 1
	Loop
	my11stGoodNo = left(arrmy11stGoodNo,len(arrmy11stGoodNo)-1)
End If

Set oMy11st = new CMy11st
	oMy11st.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	oMy11st.FPageSize					= 50
Else
	oMy11st.FPageSize					= 20
End If
	oMy11st.FRectCDL					= request("cdl")
	oMy11st.FRectCDM					= request("cdm")
	oMy11st.FRectCDS					= request("cds")
	oMy11st.FRectItemID					= itemid
	oMy11st.FRectItemName				= itemname
	oMy11st.FRectSellYn					= sellyn
	oMy11st.FRectLimitYn				= limityn
	oMy11st.FRectSailYn					= sailyn
	oMy11st.FRectonlyValidMargin		= onlyValidMargin
	oMy11st.FRectMakerid				= makerid
	oMy11st.FRectMy11stGoodNo			= my11stGoodNo
	oMy11st.FRectMatchCate				= MatchCate
	oMy11st.FRectIsMadeHand				= isMadeHand
	oMy11st.FRectIsOption				= isOption
	oMy11st.FRectIsReged				= isReged
	oMy11st.FRectDeliverytype			= deliverytype
	oMy11st.FRectMwdiv					= mwdiv

	oMy11st.FRectExtNotReg				= ExtNotReg
	oMy11st.FRectExpensive10x10			= expensive10x10
	oMy11st.FRectdiffPrc				= diffPrc
	oMy11st.FRectMy11stYes10x10No		= my11stYes10x10No
	oMy11st.FRectMy11stNo10x10Yes		= my11stNo10x10Yes
	oMy11st.FRectExtSellYn				= extsellyn
	oMy11st.FRectInfoDiv				= infoDiv
	oMy11st.FRectFailCntOverExcept		= ""
	oMy11st.FRectFailCntExists			= failCntExists
	oMy11st.FRectReqEdit				= reqEdit
If (bestOrd = "on") Then
    oMy11st.FRectOrdType = "B"
ElseIf (bestOrdMall = "on") Then
    oMy11st.FRectOrdType = "BM"
End If

If isReged = "R" Then						'ǰ��ó����� ��ǰ���� ����Ʈ
	oMy11st.getMy11streqExpireItemList
Else
	oMy11st.getMy11stRegedItemList		'�� �� ����Ʈ
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
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
				comp.form.MatchCate.value = "Y";
			}
	        if (comp.checked){
				comp.form.ExtNotReg.disabled = true;
	        }else if(comp.checked == false){
				comp.form.ExtNotReg.disabled = false;
	        }
	    }
    }

    if ((comp.name=="my11stYes10x10No")&&(comp.checked)){
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

    if ((comp.name=="my11stNo10x10Yes")&&(comp.checked)){
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
        if (comp.form.my11stYes10x10No.checked){
            comp.form.my11stYes10x10No.checked = false;
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
			//comp.form.onlyValidMargin.value="Y";
			//comp.form.extsellyn.value = "Y";
			/* �ϴ� 3�� �� ��ġ�� ���� �� �� ���� �ּ����� */
			comp.form.sellyn.value = "A";
			comp.form.extsellyn.value = "";
			comp.form.onlyValidMargin.value="";        }
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
	if ((comp.name!="my11stYes10x10No")&&(frm.my11stYes10x10No.checked)){ frm.my11stYes10x10No.checked=false }
	if ((comp.name!="my11stNo10x10Yes")&&(frm.my11stNo10x10Yes.checked)){ frm.my11stNo10x10Yes.checked=false }
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
	if ((comp.name!="my11stYes10x10No")&&(frm.my11stYes10x10No.checked)){ frm.my11stYes10x10No.checked=false }
	if ((comp.name!="my11stNo10x10Yes")&&(frm.my11stNo10x10Yes.checked)){ frm.my11stNo10x10Yes.checked=false }
	if ((comp.name!="reqEdit")&&(frm.reqEdit.checked)){ frm.reqEdit.checked=false }
}
//Que �α� ����Ʈ �˾�
function pop_quelog(mallid) {
	var pCM5 = window.open("/admin/etc/que/popQueLogList.asp?mallid="+mallid,"pop_quelog","width=1400,height=600,scrollbars=yes,resizable=yes");
	pCM5.focus();
}
// ���õ� ��ǰ �ϰ� ���
function my11stSelectRegProcess() {
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
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function my11stEditProcess(){
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

	if (confirm('11������ �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
	}
}

// ���õ� ��ǰ ���� �ϰ� ����
function my11stPriceEditProcess(){
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

    if (confirm('11������ �����Ͻ� ' + chkSel + '�� ������ ���� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "PRICE";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ����
function my11stOptEditProcess(){
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

    if (confirm('11������ �����Ͻ� ' + chkSel + '�� �ɼ��� ���� �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDITOPT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �ϰ� ��ȸ
function my11stViewEditProcess(){
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

    if (confirm('11������ �����Ͻ� ' + chkSel + '�� ������ �ϰ� ��ȸ �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "VIEW";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ �ɼ� ��ȸ
function my11stViewOptProcess(){
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

    if (confirm('11������ �����Ͻ� ' + chkSel + '�� �ɼ��� ��ȸ �Ͻðڽ��ϱ�?\n\n��11�������� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "VIEWOPT";
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ �Ǹſ��� ����
function my11stSellYnProcess(chkYn) {
	var chkSel=0;
	var strSell;
	var strcmdparam;
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
	if(chkYn == "Y"){
		strSell = "�Ǹ���";
		strcmdparam = "ONSALE";
	}else if(chkYn == "N"){
		strSell = "�Ǹ�����";
		strcmdparam = "SOLDOUT";
	}

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��������ũ���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = strcmdparam;
		document.frmSvArr.chgSellYn.value = chkYn;
		document.frmSvArr.action = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp"
		document.frmSvArr.submit();
    }
}
//�����ڵ� �˻�
function my11stCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	if(ccd == ''){
		alert('�����ڵ带 �����ϼ���');
		return;
	}
	if (confirm('�����Ͻ� �ڵ带 �˻� �Ͻðڽ��ϱ�?')){
		xLink.location.href = "<%=apiURL%>/outmall/11stmy/actmy11stReq.asp?cmdparam=my11stCommonCode&CommCD="+ccd+"";
	}
}
// 11���� ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/my11st/popmy11stCateList.asp","popmy11st","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
// 11���� ��ǰ ����
function pop_my11stManager(itemid){
	var pCM = window.open("/admin/etc/my11st/popmy11stManager.asp?itemid="+itemid,"popmy11stManager","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
function PopItemContent(iitemid){
	var popwin = window.open('/admin/itemmaster/overseas/popItemContent.asp?itemid=' + iitemid +'&sitename=11STMY&ml=EN','itemWeightEdit','width=1024,height=768,scrollbars=yes,resizable=yes')
	popwin.focus();
}
//�ɼ� �� �˾�
function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=11stmy&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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
		<a href="https://soffice.11street.my/login.do" target="_blank">11st_Admin�ٷΰ���</a>
		<%
			If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") OR (session("ssBctID")="hrkang97") Then
				response.write "<font color='GREEN'>[ llkkjj0906@10x10.co.kr | xpsqkdlxps1010 ]</font>"
			End If
		%>
		<br>
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		&nbsp;
		11st ��ǰ�ڵ� : <textarea rows="2" cols="20" name="my11stGoodNo" id="itemid"><%=replace(my11stGoodNo,",",chr(10))%></textarea>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��Ͽ��� :
		<select name="ExtNotReg" class="select" <%=Chkiif(isReged <> "", "disabled","") %> >
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >11st ��Ͻ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >11st ���۽õ�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >11st ��Ͽ���
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >11st ��ϿϷ�(����)
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
		<label><input onClick="checkisReged(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>11st ����<�ٹ����� �ǸŰ�</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����</label>
		<% If (session("ssBctID")="kjy8517") Then %>
		&nbsp;
		<label><input onClick="onlyJY(this)" type="checkbox" name="morningJY" <%= ChkIIF(morningJY="on","checked","") %> >��������</label>
		<% End If %>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="my11stYes10x10No" <%= ChkIIF(my11stYes10x10No="on","checked","") %> ><font color=red>11st�Ǹ���&�ٹ�����ǰ��</font>��ǰ����</label>
		&nbsp;
		<label><input onClick="checkisReged(this)" type="checkbox" name="my11stNo10x10Yes" <%= ChkIIF(my11stNo10x10Yes="on","checked","") %> ><font color=red>11stǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����</label>
		<br>
		<label><input onClick="checkisReged(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)</label>
	</td>
</tr>
</form>
</table>
<p>
<form name="frmReg" method="post" action="lotteItem.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="delitemid" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="30" bgcolor="#FFFFFF">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="right">
				<input class="button" type="button" value="QUE LOG" onclick="pop_quelog('11stmy');">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���" onClick="my11stSelectRegProcess(true);">
				&nbsp;&nbsp;
				<br><br>
				������ǰ ���� :
				<input class="button" type="button" id="btnEdit" value="����" onClick="my11stEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditPrice" value="����" onClick="my11stPriceEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditOpt" value="�ɼ�" onClick="my11stOptEditProcess();">
			<% If (session("ssBctID")="kjy8517") OR (session("ssBctID")="icommang") Then %>
				<br><br>
				������ǰ ��ȸ :
				<input class="button" type="button" id="btnEditView" value="��ǰ" onClick="my11stViewEditProcess();">
				&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditView" value="�ɼ�" onClick="my11stViewOptProcess();">
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="">- Choice -
					<option value="CATEGORYLIST">ī�װ�
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="my11stCommCDSubmit();" >
			<% End If %>
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ǹ�����</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="my11stSellYnProcess(frmReg.chgSellYn.value);">
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
		�˻���� : <b><%= FormatNumber(oMy11st.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oMy11st.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">11���������<br>11��������������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">��ǰ<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">�Ǹŵɰ���</td>
	<td width="70">11����<br>���ݹ��Ǹ�</td>
	<td width="70">11����<br>��ǰ��ȣ</td>
	<td width="50">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">����</td>
</tr>
<% For i=0 to oMy11st.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= oMy11st.FItemList(i).FItemID %>"></td>
	<td><img src="<%= oMy11st.FItemList(i).Fsmallimage %>" width="50"></td>
	<td align="center">
		<a href="<%=wwwURL%>/<%=oMy11st.FItemList(i).FItemID%>" target="_blank"><%= oMy11st.FItemList(i).FItemID %></a>
		<% If oMy11st.FItemList(i).FMy11stStatcd <> 7 Then %>
		<br><%= oMy11st.FItemList(i).getMy11stStatName %>
		<% End If %>
		<%= oMy11st.FItemList(i).getLimitHtmlStr %>
	</td>
	<td align="left"><%= oMy11st.FItemList(i).FMakerid %> <%= oMy11st.FItemList(i).getDeliverytypeName %><br><%= oMy11st.FItemList(i).FItemName %></td>
	<td align="center"><%= oMy11st.FItemList(i).FRegdate %><br><%= oMy11st.FItemList(i).FLastupdate %></td>
	<td align="center"><%= oMy11st.FItemList(i).FMy11stRegdate %><br><%= oMy11st.FItemList(i).FMy11stLastUpdate %></td>
	<td align="right">
		<%= FormatNumber(oMy11st.FItemList(i).FOrgprice,0) %>(��)<br>
		<font color="red"><%= FormatNumber(oMy11st.FItemList(i).FSellcash,0) %>(��)</font>
	</td>
	<td align="center">
	<%
		If oMy11st.FItemList(i).Fsellcash <> 0 Then
			response.write CLng(10000-oMy11st.FItemList(i).Fbuycash / oMy11st.FItemList(i).Fsellcash*100*100)/100 & "%" &" <br>"
		End If
	%>
	</td>
	<td align="center"><%= FormatNumber((oMy11st.FItemList(i).FitemWeight/1000),3) %>kg</td>
	<td align="center">
	<%
		If oMy11st.FItemList(i).IsSoldOut Then
			If oMy11st.FItemList(i).FSellyn = "N" Then
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
		If oMy11st.FItemList(i).FItemdiv = "06" OR oMy11st.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If oMy11st.FItemList(i).FMaySellPrice <> "" Then
			response.write CDBL(formatNumber(oMy11st.FItemList(i).FMaySellPrice,2))
		End If
	%>
	</td>
	<td align="center">
	<%
		If (oMy11st.FItemList(i).FMy11stStatCd > 0) Then
			If Not IsNULL(oMy11st.FItemList(i).FMy11stPrice) Then
				If (oMy11st.FItemList(i).FOrgprice <> oMy11st.FItemList(i).FRegOrgprice) Then
	%>
					<strong><%= CDBL(formatNumber(oMy11st.FItemList(i).FMy11stPrice,2)) %></strong><br>
	<%
				Else
					response.write CDBL(formatNumber(oMy11st.FItemList(i).FMy11stPrice,2))&"<br>"
				End If

				If (oMy11st.FItemList(i).FSellyn="Y" and oMy11st.FItemList(i).FMy11stSellYn<>"Y") or (oMy11st.FItemList(i).FSellyn<>"Y" and oMy11st.FItemList(i).FMy11stSellYn="Y") Then
	%>
					<strong><%= oMy11st.FItemList(i).FMy11stSellYn %></strong>
	<%
				Else
					response.write oMy11st.FItemList(i).FMy11stSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(oMy11st.FItemList(i).FMy11stGoodNo)) Then
			Response.Write "<a target='_blank' href='http://www.11street.my/product/ProductDetailAction/getProductDetail.do?prdNo="&oMy11st.FItemList(i).FMy11stGoodNo&"'>"&oMy11st.FItemList(i).FMy11stGoodNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= oMy11st.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=oMy11st.FItemList(i).FItemID%>','0');"><%= oMy11st.FItemList(i).FoptionCnt %>:<%= oMy11st.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= oMy11st.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If oMy11st.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
		If (oMy11st.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& db2html(oMy11st.FItemList(i).FlastErrStr) &"'>ERR:"& oMy11st.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
	<td>
		<input type="button" class="button" value="����" onclick="PopItemContent('<%=oMy11st.FItemList(i).FItemid%>')">
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if oMy11st.HasPreScroll then %>
		<a href="javascript:goPage('<%= oMy11st.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oMy11st.StartScrollPage to oMy11st.FScrollCount + oMy11st.StartScrollPage - 1 %>
    		<% if i>oMy11st.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oMy11st.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->