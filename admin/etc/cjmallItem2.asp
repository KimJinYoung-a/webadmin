<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/cjmall2/cjmallitemcls.asp"-->
<!-- #include virtual="/admin/etc/cjmall2/incCJMallFunction.asp"-->
<%
Dim makerid, cjmallitemid, itemname, itemid, eventid, ExtNotReg, bestOrd, bestOrdMall, MatchCate, sellyn, limityn, sailyn, onlyValidMargin, showminusmargin, MatchPrddiv
Dim optAddprcExists, optAddprcExistsExcept
Dim cjSell10x10Soldout, expensive10x10, cjshowminusmagin
Dim failCntExists, optExists, diffPrc, optnotExists, isMadeHand
Dim page, i, research, infodiv, reqExpire, extsellyn
page    				= request("page")
itemid  				= request("itemid")

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

makerid					= request("makerid")
eventid					= request("eventid")
itemname				= request("itemname")
cjmallitemid			= request("cjmallitemid")
ExtNotReg				= request("ExtNotReg")
bestOrd					= request("bestOrd")
bestOrdMall				= request("bestOrdMall")
MatchCate				= request("MatchCate")
MatchPrddiv				= request("MatchPrddiv")
sellyn					= request("sellyn")
limityn					= request("limityn")
sailyn					= request("sailyn")
onlyValidMargin			= request("onlyValidMargin")
showminusmargin			= request("showminusmargin")
research				= request("research")
failCntExists			= request("failCntExists")
infodiv					= request("infodiv")
optAddprcExists 		= request("optAddprcExists")
optAddprcExistsExcept	= request("optAddprcExistsExcept")
optExists   			= request("optExists")
optnotExists   			= request("optnotExists")
isMadeHand				= request("isMadeHand")
cjSell10x10Soldout      = request("cjSell10x10Soldout")
cjshowminusmagin		= request("cjshowminusmagin")
expensive10x10          = request("expensive10x10")
reqExpire				= request("reqExpire")
extsellyn				= request("extsellyn")
diffPrc					= request("diffPrc")

If page = "" Then page = 1
If sellyn="" Then sellyn = "Y"

''�⺻���� ��Ͽ����̻�
If (research = "") Then
	ExtNotReg = "J"
	MatchCate = ""
	MatchPrddiv = ""
	onlyValidMargin = "on"  ''on
	bestOrd = "on"
	sellyn = "Y"
End If

Dim cjMall
Set cjMall = new CCjmall
	cjMall.FCurrPage					= page
If (session("ssBctID")="kjy8517") Then
	cjMall.FPageSize					= 50
Else
	cjMall.FPageSize					= 20
End If
	cjMall.FRectMakerid					= makerid
	cjMall.FRectItemName				= itemname
	cjMall.FRectCJMallPrdNo				= cjmallitemid
	cjMall.FRectCDL 					= request("cdl")
	cjMall.FRectCDM 					= request("cdm")
	cjMall.FRectCDS 					= request("cds")
	cjMall.FRectItemID					= itemid
	cjMall.FRectEventid					= eventid
	cjMall.FRectExtNotReg				= ExtNotReg
	cjMall.FRectMatchCate				= MatchCate
	cjMall.FRectPrdDivMatch				= MatchPrddiv
	cjMall.FRectSellYn					= sellyn
	cjMall.FRectLimitYn					= limityn
	cjMall.FRectSailYn					= sailyn
	cjMall.FRectonlyValidMargin 		= onlyValidMargin
	cjMall.FRectMinusMargin 			= showminusmargin
	cjMall.FRectFailCntExists			= failCntExists
	cjMall.Finfodiv						= infodiv
	cjMall.FRectdiffPrc 				= diffPrc
	cjMall.FRectoptAddprcExists			= optAddprcExists
	cjMall.FRectoptAddprcExistsExcept	= optAddprcExistsExcept
	cjMall.FRectoptExists				= optExists
	cjMall.FRectoptnotExists			= optnotExists
	cjMall.FRectisMadeHand				= isMadeHand
	cjMall.FRectCjSell10x10Soldout      = CjSell10x10Soldout
	cjMall.FRectCjshowminusmagin		= cjshowminusmagin
	cjMall.FRectExpensive10x10          = expensive10x10
	cjMall.FRectExtSellYn  				= extsellyn
	If (bestOrd="on") Then
	    cjMall.FRectOrdType = "B"
	ElseIf (bestOrdMall="on") Then
	    cjMall.FRectOrdType = "BM"
	End If

	If reqExpire <> "" Then
	    cjMall.getCjmallreqExpireItemList
	Else
	    cjMall.GetCjmallRegedItemList
	End If

If (session("ssBctID")="kjy8517") Then
	Dim JYSQL, qq, ww
	ww = ""
	JYSQL = ""
	JYSQL = JYSQL & " select TOP 100 itemid from "
	JYSQL = JYSQL & " db_outmall.dbo.tbl_item "
	JYSQL = JYSQL & " where itemid in ( "
	JYSQL = JYSQL & " 	select itemid from [db_outmall].[dbo].tbl_OutMall_regedoption where mallid = 'cjmall' and itemoption = '0000' and outmallsellyn = 'N' and lastupdate < '2013-10-07' group by itemid "
	JYSQL = JYSQL & " ) "
	JYSQL = JYSQL & " and isusing = 'Y' "
	JYSQL = JYSQL & " and sellyn = 'Y' "
	JYSQL = JYSQL & " and itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='cjmall') "
	JYSQL = JYSQL & " and makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='cjmall') "
	JYSQL = JYSQL & " and itemid not in (147391,147393,230552,359218,359220,368619,394415,399605,483146,483153,496532,552817,598861,598862,598867,598880,598883,598889,625614,642212,643690,662015,662022,662027,672830,673715,691890,695865,704132,742944,743918,745879,755202,755203,755204,755205,755206,755207,755227,755230,755232,755247,763029,763495,771158,771160,795016,795021,795022,795025,795026,795468,800146,802496,802497,802499,805812,818832,818833,818834,818835,819415,819434,819436,819446,819452,819457,819459,819462,819465,819469,821876,826952,827518,827528,846214,848259,848260,848794,848887,850412,850413,850592,850593,850594,858009,858395,858654,858658,858661,858667,858684,858689,858702,859978,859979,859980,859982,864586,864609,864612,864614,864619,864620,864621,864623,864628,864629,864670,867525,877189,877342,877343,877345,880644,882650,882652,883089,884608,884609,888093,892290,892992,893000,897417,898864,899538,899539,899540,899541,899542,899543,899544,899545,899546,899547,899548,899549,899550,899551,899553,899555,899556,899557,899558,899561,899562,899563,899605,900836,910593,913629,916273,916274,916324,916330,918488,918491,918492,918493,918495,918496,918497,918498) "			'���ǹ���̸鼭 10000�� �̸���ǰ
	JYSQL = JYSQL & " group by itemid "
	JYSQL = JYSQL & " ORDER by itemid ASC "
'	rsCTget.Open JYSQL, dbCTget
'	If Not(rsCTget.EOF or rsCTget.BOF) Then
'		For qq = 1 to rsCTget.RecordCount
'			if qq = rsCTget.RecordCount Then
'				ww = ww & trim(rsCTget("itemid"))
'			Else
'				ww = ww & trim(rsCTget("itemid"))&","
'			End If
'			rsCTget.MoveNext
'		Next
'	End If
'	rsCTget.Close
End If

%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}


function popManageOptAddPrc(iitemid,mngOptAdd){
    var pwin = window.open("/admin/etc/popOptionAddPrcSet.asp?itemid="+iitemid+'&mallid=cjmall&mngOptAdd='+mngOptAdd,"popOptionAddPrc","width=800,height=600,scrollbars=yes,resizable=yes");
	pwin.focus();
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

    if ((comp.name=="cjSell10x10Soldout")&&(comp.checked)){
        if (comp.form.expensive10x10.checked){
            comp.form.expensive10x10.checked = false;
        }

        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "A";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
    }

    if ((comp.name=="expensive10x10")&&(comp.checked)){
        if (comp.form.cjSell10x10Soldout.checked){
            comp.form.cjSell10x10Soldout.checked = false;
        }

        comp.form.ExtNotReg.value = "D";
        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "Y";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
        comp.form.onlyValidMargin.checked=false;
    }

    if ((comp.name=="cjshowminusmagin")&&(comp.checked)){

        comp.form.ExtNotReg.value = "D";
        comp.form.MatchCate.value = "";
        comp.form.MatchPrddiv.value = "";
        comp.form.sellyn.value = "Y";
        comp.form.limityn.value = "";
        comp.form.infodiv.value = "";
        comp.form.onlyValidMargin.checked=false;
    }

	if ((comp.name=="reqExpire")&&(comp.checked)){
	    comp.form.ExtNotReg.value="D";
	    comp.form.MatchCate.value="Y";
	    comp.form.sellyn.value="A";
	    comp.form.limityn.value="";
	    comp.form.onlyValidMargin.checked=false;
	}
	if ((comp.name!="cjshowminusmagin")&&(frm.cjshowminusmagin.checked)){ frm.cjshowminusmagin.checked=false }
	if ((comp.name!="reqExpire")&&(frm.reqExpire.checked)){ frm.reqExpire.checked=false }
	if ((comp.name=="diffPrc")){frm.onlyValidMargin.checked=true;}
}

// ������� �귣��
function NotInMakerid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Makerid.asp?mallgubun=cjmall','notin','width=300,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
// ������� ��ǰ
function NotInItemid(){
	var popwin=window.open('/admin/etc/JaehyuMall_Not_In_Itemid.asp?mallgubun=cjmall','notinItem','width=500,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}
//ī�װ� ����
function pop_CateManager() {
	var pCM = window.open("/admin/etc/cjmall2/popcjmallCateList.asp","popCateMancjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM.focus();
}
//��ǰ�з� ����
function pop_prdDivManager() {
	var pCM2 = window.open("/admin/etc/cjmall2/popcjmallprdDivList.asp","popprdDivcjMall","width=1000,height=600,scrollbars=yes,resizable=yes");
	pCM2.focus();
}
// ���õ� ��ǰ �ϰ� ���
function CjregIMSI(isreg) {
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
		if (confirm('CjMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelectWait";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}else{
		if (confirm('CjMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� ���� ��� ���� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "DelSelectWait";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
		}
	}
}

// ���õ� ��ǰ �ϰ� ���
function CjSelectRegProcess(isreal) {
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
		if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ��� �Ͻðڽ��ϱ�?')){
			document.getElementById("btnRegSel").disabled=true;
			document.frmSvArr.target = "xLink";
			document.frmSvArr.cmdparam.value = "RegSelect";
			document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
			document.frmSvArr.submit();
        }
	}
}

// ���õ� ��ǰ �ϰ� ����
function CjSelectEditProcess() {
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

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

// ���õ� ��ǰ ����+��ǰ �ϰ� ����
function CjSelectEdit2Process() {
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

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ�� �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnEditSel").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditSelect2";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}


function CjSelectPriceEditProcess() {
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

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
       // document.getElementById("btnEditSelPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

function CjSelectPriceEditProcess2() {
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

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        //document.getElementById("btnPrice").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditPriceSelect2";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}

//���û�ǰ ��������
function CjSelectQTYEditProcess(){
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

    if (confirm('CJMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ ���� �Ͻðڽ��ϱ�?\n\n��CJMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        document.getElementById("btnEditqty").disabled=true;
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EditQty";
        document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
        document.frmSvArr.submit();
    }
}


// ���õ� ��ǰ �Ǹſ��� ����
function CjmallSellYnProcess(chkYn) {
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

    if (confirm('�����Ͻ� ' + chkSel + '�� ��ǰ�� �Ǹſ��θ� "' + strSell + '"(��)�� ���� �Ͻðڽ��ϱ�?\n\n��cjmall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
        if (chkYn=="X"){
            if (!confirm(strSell + '�� �����ϸ� cjmall���� ���� �Ұ�/��ϸ�Ͽ��� �����Ǹ� ���ǸŽ�  ���� ���� ����ϼž� �մϴ�. ��� �Ͻðڽ��ϱ�?')) return;
        }

        //document.getElementById("btnSellYn").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EditSellYn";
		document.frmSvArr.subcmd.value = chkYn;
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
    }
}
// ���õ� ��ǰ ��ǰ, �ǸŻ��� ����
function CjSelectSaleStatEditProcess() {
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

	if (confirm('CjMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CjMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditDanpum").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdSaleDTSel";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}

//���û�ǰ �������
function CjSelectDateEditProcess() {
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

	if (confirm('CjMall�� �����Ͻ� ' + chkSel + '�� ��ǰ ������ �ϰ� ���� �Ͻðڽ��ϱ�?\n\n��CjMall���� ��Ż��¿� ���� �ð��� �ټ� �ɸ� �� �ֽ��ϴ�.')){
		document.getElementById("btnEditDate").disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EdDateSel";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	}
}

//���û�ǰ ����Ȯ�� �� �ǸŻ��� check - batch
function batchStatCheck(){
    document.frmSvArr.target = "xLink";
	document.frmSvArr.cmdparam.value = "confirmItemAuto";
	document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
	document.frmSvArr.submit();
}


//���û�ǰ ����Ȯ�� �� �ǸŻ��� check
function checkCjItemConfirm(comp) {
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

	//if (confirm('���� ��ǰ ���ο��� �� �ǸŻ��� ��ȸ �Ͻðڽ��ϱ�?')){
		//comp.disabled=true;
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "confirmItem";
		document.frmSvArr.action = "/admin/etc/cjmall/actCjMallReq.asp"
		document.frmSvArr.submit();
	//}
}

function CjSellynSubmit(yn){

	if(yn == true){
	    if (document.getElementById('theday').value.length!=10) {
	        alert('��¥ �������� �Է��� �ּ���.yyyy-mm-dd');
	        return;
	     }
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=LIST&sday="+document.getElementById('theday').value+"";
	}else{
		xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=DayLIST&sday="+document.getElementById('sday').value+"";
	}
	document.getElementById("btnSell1").disabled=true;
	document.getElementById("btnSell2").disabled=true;
}


function NumObj(obj){
	if (event.keyCode >= 48 && event.keyCode <= 57) { //����Ű�� �Է�
		return true;
	} else {
		event.returnValue = false;
	}
}
function popCjCommCDSubmit() {
	var ccd;
	ccd = document.getElementById('CommCD').value;
	xLink.location.href = "/admin/etc/cjmall/actCjMallreq.asp?cmdparam=cjmallCommonCode&CommCD="+ccd+"";
}
function popOutMallEtcLink(itemid,mallid,poomok){
    var popwin = window.open('/admin/etc/common/popOutMallEtcLink.asp?mallid='+mallid+'&itemid='+itemid+'&poomok='+poomok+'','popOutMallEtcLink','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
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
		Cjmall��ǰ��ȣ: <input type="text" name="cjmallitemid" value="<%= cjmallitemid %>" size="9" maxlength="9" class="input">
		��ǰ��: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="input">
		&nbsp;
		<a href="http://partner.cjmall.com/login.jsp" target="_blank">CJ��Admin�ٷΰ���</a>
		<%
			If C_ADMIN_AUTH Then
				response.write "<font color='GREEN'>[ 411378 | store10x10 | 1010cube* ]</font>"
			End If
		%>
		<br>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		��ǰ��ȣ: <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
   		�̺�Ʈ��ȣ: <input type="text" name="eventid" value="<%= eventid %>" size="6" maxlength="6" class="input">
    	<br>
    	��Ͽ��� :
		<select name="ExtNotReg" class="select">
			<option value="">��ü
			<option value="M" <%= CHkIIF(ExtNotReg="M","selected","") %> >CJmall �̵��(��ϰ���)
			<option value="Q" <%= CHkIIF(ExtNotReg="Q","selected","") %> >CJmall ��Ͻ���
			<option value="J" <%= CHkIIF(ExtNotReg="J","selected","") %> >CJmall ��Ͽ����̻�
			<option value="W" <%= CHkIIF(ExtNotReg="W","selected","") %> >CJmall ��Ͽ���
			<option value="A" <%= CHkIIF(ExtNotReg="A","selected","") %> >CJmall ���۽õ��߿���
			<option value="F" <%= CHkIIF(ExtNotReg="F","selected","") %> >CJmall ����� ���δ��(�ӽ�)
			<option value="D" <%= CHkIIF(ExtNotReg="D","selected","") %> >CJmall ��ϿϷ�(����)
			<option value="R" <%= CHkIIF(ExtNotReg="R","selected","") %> >CJmall �������
		</select>&nbsp;
		<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
		<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��(����)</b>&nbsp;
		ī�׸�Ī :
		<select name="MatchCate" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchCate="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchCate="N","selected","") %> >�̸�Ī
		</select>&nbsp;
		��ǰ�з���Ī :
		<select name="MatchPrddiv" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(MatchPrddiv="Y","selected","") %> >��Ī
			<option value="N" <%= CHkIIF(MatchPrddiv="N","selected","") %> >�̸�Ī
		</select>&nbsp;
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
		���Ͽ��� :
		<select name="sailyn" class="select">
			<option value="">��ü
			<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
			<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
		</select>&nbsp;
		ǰ������ :
		<select name="infodiv" class="select">
			<option value="" <%= CHkIIF(infoDiv="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >�Է�
			<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >���Է�
			<option value="01" <%= CHkIIF(infodiv="01","selected","") %> >01
			<option value="02" <%= CHkIIF(infodiv="02","selected","") %> >02
			<option value="03" <%= CHkIIF(infodiv="03","selected","") %> >03
			<option value="04" <%= CHkIIF(infodiv="04","selected","") %> >04
			<option value="05" <%= CHkIIF(infodiv="05","selected","") %> >05
			<option value="06" <%= CHkIIF(infodiv="06","selected","") %> >06
			<option value="07" <%= CHkIIF(infodiv="07","selected","") %> >07
			<option value="08" <%= CHkIIF(infodiv="08","selected","") %> >08
			<option value="09" <%= CHkIIF(infodiv="09","selected","") %> >09
			<option value="10" <%= CHkIIF(infodiv="10","selected","") %> >10
			<option value="11" <%= CHkIIF(infodiv="11","selected","") %> >11
			<option value="35" <%= CHkIIF(infodiv="12","selected","") %> >35
		</select>&nbsp;
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >���� <%= CMAXMARGIN %>%�̻� ��ǰ�� ����
		&nbsp;
		<input type="checkbox" name="failCntExists" <%= ChkIIF(failCntExists="on","checked","") %> >��ϼ���������ǰ
		<br>
		<input type="checkbox" name="optAddprcExists" <%= ChkIIF(optAddprcExists="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ
		&nbsp;
		<input type="checkbox" name="optAddprcExistsExcept" <%= ChkIIF(optAddprcExistsExcept="on","checked","") %> onClick="checkComp(this)">�ɼ��߰��ݾ������ǰ����
		&nbsp;
		<input type="checkbox" name="optExists" <%= ChkIIF(optExists="on","checked","") %> >�ɼ������ǰ
		&nbsp;
		<input type="checkbox" name="optnotExists" <%= ChkIIF(optnotExists="on","checked","") %> >��ǰ��ǰ(�ɼ�=0)
		&nbsp;
		�ֹ����ۿ��� :
		<select name="isMadeHand" class="select">
			<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
			<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
			<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
		</select>
		<br>
		<input type="checkbox" name="cjshowminusmagin"  <%= ChkIIF(cjshowminusmagin="on","checked","") %> onClick="checkComp(this)"  ><font color=red>������</font>��ǰ���� (MaxMagin : <%= CMAXMARGIN %>%) (CJ �Ǹ���)
		&nbsp;
		<input type="checkbox" name="cjSell10x10Soldout" <%= ChkIIF(cjSell10x10Soldout="on","checked","") %> onClick="checkComp(this)"><font color=red>CJ�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
		&nbsp;
		<input type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> onClick="checkComp(this)"><font color=red>CJ ����<�ٹ����� �ǸŰ�</font>��ǰ����
		&nbsp;
		<input onClick="checkComp(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
		<br>
		<input onClick="checkComp(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (���޸� �����Ե�)
		&nbsp;&nbsp;�����ǸŻ��� :
		<select name="extsellyn" class="select">
		<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
		<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
		<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
		<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
		<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
		</select>
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>

<br>
<form name="frmReg" method="post" action="cjmallItem.asp" style="margin:0px;">
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
				<font color="RED">���� 2�� ���۾� �ʿ�! :</font>
				<input class="button" type="button" value="cjMall��ǰ�з���Ī" onclick="pop_prdDivManager();">&nbsp;&nbsp;
				<input class="button" type="button" value="cjMallī�װ���Ī" onclick="pop_CateManager();">&nbsp;&nbsp;
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
				<input class="button" type="button" id="btnRegSel" value="���û�ǰ ���� ���" onClick="CjSelectRegProcess(true);">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSel" value="���û�ǰ ���� ����" onClick="CjSelectEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditSelPrice" value="���û�ǰ ��ǰ���� ���� ����" onClick="CjSelectPriceEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditqty" value="���û�ǰ ���� ����" onClick="CjSelectQTYEditProcess();">&nbsp;&nbsp;
				<input class="button" type="button" id="btnEditDanpum" value="���� ��ǰ ����" onClick="CjSelectSaleStatEditProcess();">
				<input class="button" type="button" id="btnEditDate" value="���û�ǰ ����+��ǰ ����" onClick="CjSelectEdit2Process();">
				<% If C_ADMIN_AUTH or (session("ssBctID")="kjy8517") or (session("ssBctID")="cogusdk") or (session("ssBctID")="areum531") or (session("ssBctID")="therthis") or (session("ssBctID")="joohyun49") then %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input class="button" type="button" id="btnPrice" value="���û�ǰ �ǸŰ��� ���� ����" onClick="CjSelectPriceEditProcess2();">
				<% End If %>
				<!--
				<input class="button" type="button" id="btnEditDate" value="���û�ǰ �������" onClick="CjSelectDateEditProcess();">
				-->
				<br><br>
				�������� ���� :
				<input class="button" type="button" id="btnRegSel" value="���û�ǰ ���� ���" onClick="CjregIMSI(true);">&nbsp;&nbsp;
				<input class="button" type="button" id="btnRegSel" value="���û�ǰ ���� ����" onClick="CjregIMSI(false);" >
				<br><br>
				���ο��� �˻� :
				<!--
				<input type="text" name="theday" value="" size="10" maxlength="10">
				<input class="button" type="button" id="btnSell1" value="Ư����¥ ���ο��� Ȯ��" onClick="CjSellynSubmit(true);">&nbsp;&nbsp;
				<select name="sday" class="select" id="sday">
				<% For i = 0 to 9 %>
					<option value="<%=i%>"><%=i%>
				<% Next %>
				</select>����&nbsp;
				<input class="button" type="button" id="btnSell2" value="���� �Ⱓ ���ο��� Ȯ��" onClick="CjSellynSubmit(false);" >
				-->
				<input class="button" type="button"  value="���û�ǰ ����(�ǸŻ���) Ȯ��" onClick="checkCjItemConfirm(this);" >
				<br><br>
				�����ڵ� �˻� :
				<select name="CommCD" class="select" id="CommCD">
					<option value="L126">�ù���ڵ�
					<option value="6009">����Ÿ��
					<option value="8047">�����ä�α���
				</select>
				<input class="button" type="button" id="btnCommcd" value="�����ڵ�Ȯ��" onClick="popCjCommCDSubmit();" >
			</td>
			<td align="right" valign="top">
				<br><br>
				���û�ǰ��
				<Select name="chgSellYn" class="select">
					<option value="N">�Ͻ��ߴ�</option>
					<option value="Y">�Ǹ���</option>
				</Select>(��)��
				<input class="button" type="button" id="btnSellYn" value="����" onClick="CjmallSellYnProcess(frmReg.chgSellYn.value);">

				<br><br><input type="button" value="�ǸŻ���Check(������)" onClick="batchStatCheck();">
			</td>
		</tr>
		</table>
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
	<td colspan="17" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(cjMall.FTotalPage,0) %> �ѰǼ�: <%= FormatNumber(cjMall.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
	<td width="20"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="50">�̹���</td>
	<td width="60">�ٹ�����<br>��ǰ��ȣ</td>
	<td>�귣��<br>��ǰ��</td>
	<td width="140">��ǰ�����<br>��ǰ����������</td>
	<td width="140">CJMall�����<br>CJMall����������</td>
	<td width="70">�ٹ�����<br>�ǸŰ�</td>
	<td width="70">�ٹ�����<br>����</td>
	<td width="70">ǰ������</td>
	<td width="70">�ֹ�����<br>����</td>
	<td width="70">CJMall<br>���ݹ��Ǹ�</td>
	<td width="70">CJMall<br>��ǰ��ȣ</td>
	<td width="80">�����ID</td>
	<td width="50">�ɼǼ�</td>
	<td width="50">3����<br>�Ǹŷ�</td>
	<td width="60">ī�װ�<br>��Ī����</td>
	<td width="60">��ǰ�з�<br>��Ī����</td>
</tr>
<% For i = 0 To cjMall.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  value="<%= cjMall.FItemList(i).FItemID %>"></td>
	<td><img src="<%= cjMall.FItemList(i).Fsmallimage %>" width="50" onClick="popOutMallEtcLink('<%= cjMall.FItemList(i).FItemID %>','cjMall','<%=cjMall.FItemList(i).FinfoDiv%>')" style="cursor:pointer"></td>
	<td align="center"><a href="<%=wwwURL%>/<%=cjMall.FItemList(i).FItemID%>" target="_blank"><%= cjMall.FItemList(i).FItemID %></a><br><%= cjMall.FItemList(i).getcjmallStatName %></td>
	<td><%= cjMall.FItemList(i).FMakerid %><%= cjMall.FItemList(i).getDeliverytypeName %><br><%= cjMall.FItemList(i).FItemName %></td>
	<td align="center"><%= cjMall.FItemList(i).FRegdate %><br><%= cjMall.FItemList(i).FLastupdate %></td>
	<td align="center"><%= cjMall.FItemList(i).FcjmallRegdate %><br><%= cjMall.FItemList(i).FcjmallLastUpdate %></td>
	<td align="right">
	<% If cjMall.FItemList(i).FSaleYn="Y" Then %>
		<strike><%= FormatNumber(cjMall.FItemList(i).FOrgPrice,0) %></strike><br>
		<font color="#CC3333"><%= FormatNumber(cjMall.FItemList(i).FSellcash,0) %></font>
	<% Else %>
		<%= FormatNumber(cjMall.FItemList(i).FSellcash,0) %>
	<% End If %>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-cjMall.FItemList(i).Fbuycash/cjMall.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).IsSoldOut Then
			If cjMall.FItemList(i).FSellyn = "N" Then
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
		If cjMall.FItemList(i).FItemdiv = "06" OR cjMall.FItemList(i).FItemdiv = "16" Then
			response.write "<font color='green'>�ֹ�����</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If (cjMall.FItemList(i).FcjmallStatCd > 0) Then
			If Not IsNULL(cjMall.FItemList(i).FcjmallPrice) Then
				If (cjMall.FItemList(i).Fsellcash <> cjMall.FItemList(i).FcjmallPrice) Then
	%>
					<strong><%= formatNumber(cjMall.FItemList(i).FcjmallPrice,0) %></strong><br>
	<%
				Else
					response.write formatNumber(cjMall.FItemList(i).FcjmallPrice,0)&"<br>"
				End If

				If (cjMall.FItemList(i).FSellyn="Y" and cjMall.FItemList(i).FcjmallSellYn<>"Y") or (cjMall.FItemList(i).FSellyn<>"Y" and cjMall.FItemList(i).FcjmallSellYn="Y") Then
	%>
					<strong><%= cjMall.FItemList(i).FcjmallSellYn %></strong>
	<%
				Else
					response.write cjMall.FItemList(i).FcjmallSellYn
				End If
			End If
		End If
	%>
	</td>
	<td align="center">
	<%
		If Not(IsNULL(cjMall.FItemList(i).FcjmallPrdNo)) Then
			Response.Write "<a target='_blank' href='http://www.cjmall.com/prd/detail_cate.jsp?item_cd="&cjMall.FItemList(i).FcjmallPrdNo&"'>"&cjMall.FItemList(i).FcjmallPrdNo&"</a>"
		End If
	%>
	</td>
	<td align="center"><%= cjMall.FItemList(i).Freguserid %></td>
	<td align="center"><a href="javascript:popManageOptAddPrc('<%=cjMall.FItemList(i).FItemID%>','0');"><%= cjMall.FItemList(i).FoptionCnt %>:<%= cjMall.FItemList(i).FregedOptCnt %></a></td>
	<td align="center"><%= cjMall.FItemList(i).FrctSellCNT %></td>
	<td align="center">
	<%
		If cjMall.FItemList(i).FCateMapCnt > 0 Then
			response.write "��Ī��"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If
	%>
	</td>
	<td align="center">
	<%
		If cjMall.FItemList(i).Fcddkey <> "" Then
			response.write "��Ī��("&cjMall.FItemList(i).Finfodiv&")"
		Else
			response.write "<font color='darkred'>��Ī�ȵ�</font>"
		End If

		If (cjMall.FItemList(i).FaccFailCNT > 0) Then
			response.write " <br><font color='red' title='"& cjMall.FItemList(i).FlastErrStr &"'>ERR:"& cjMall.FItemList(i).FaccFailCNT &"</font>"
		End If
	%>
	</td>
</tr>
<% Next %>
<tr height="20">
	<td colspan="17" align="center" bgcolor="#FFFFFF">
	<% If cjMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= cjMall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + cjMall.StartScrollPage To cjMall.FScrollCount + cjMall.StartScrollPage - 1 %>
		<% If i>cjMall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If cjMall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% Set cjMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->