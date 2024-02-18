<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ����
' Hieditor : 2011.03.07 ������ ����
'			 2011.07.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljuofflinecls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim notitemlist, itemlist ,pagesize ,research ,danpumcheck ,dcnt2 ,companyid ,locationid3PL
dim notitemlistinclude, itemlistinclude ,searchtypestring ,onejumuntype ,deliveryarea
dim notbrandlistinclude, brandlistinclude ,onlyOne,dcnt ,yyyy1,mm1,dd1,yyyymmdd,nowdate
dim imsi, sagawa, ems, epostmilitary, bigitem ,onejumuncount, onejumuncompare ,ix,iy
dim tenbeaonly ,upbeaInclude ,locationidto ,includeminus, includezerostock ,page ,ojumun, mwdiv
dim makerid
	yyyy1 = requestCheckVar(request("yyyy1"), 32)
	mm1 = requestCheckVar(request("mm1"), 32)
	dd1 = requestCheckVar(request("dd1"), 32)
	pagesize = requestCheckVar(request("pagesize"), 32)
	deliveryarea = requestCheckVar(request("deliveryarea"), 32)
	bigitem = requestCheckVar(request("bigitem"), 32)
	companyid = "10x10"
	locationid3PL = "10x10"
	upbeaInclude = request("upbeaInclude")
	tenbeaonly = request("tenbeaonly")
	notitemlistinclude = requestCheckVar(request("notitemlistinclude"), 32)
	itemlistinclude = requestCheckVar(request("itemlistinclude"), 32)
	notbrandlistinclude = requestCheckVar(request("notbrandlistinclude"), 32)
	brandlistinclude = requestCheckVar(request("brandlistinclude"), 32)
	notitemlist = requestCheckVar(request("notitemlist"), 32)
	itemlist = requestCheckVar(request("itemlist"), 32)
	research = requestCheckVar(request("research"), 32)
	onejumuntype = requestCheckVar(request("onejumuntype"), 32)
	onejumuncount = requestCheckVar(request("onejumuncount"), 32)
	onejumuncompare = requestCheckVar(request("onejumuncompare"), 32)
	locationidto = requestCheckVar(request("locationidto"), 32)
	includeminus = requestCheckVar(request("includeminus"), 32)
	includezerostock = requestCheckVar(request("includezerostock"), 32)
	makerid = requestCheckVar(request("makerid"), 32)
	'dcnt = trim(request("dcnt"))
	'dcnt2 = trim(request("dcnt2"))
	'onlyOne = request("onlyOne")
	'danpumcheck = request("danpumcheck")
	'imsi  = request("imsi")
	'sagawa= request("sagawa")
	'ems   = request("ems")
	'epostmilitary   = request("epostmilitary")
	page = request("page")
	mwdiv = requestCheckVar(request("mwdiv"), 32)

if deliveryarea = "" then deliveryarea = "KR"
if (page="") then page=1

if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if onejumuncount="" then
	onejumuncount = "1"
end if

if onejumuncompare="" then
	onejumuncompare = "less"
end if

if deliveryarea<>"" then
	if (deliveryarea = "ZZ") then
		ems   = ""
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		ems   = "on"
		epostmilitary   = ""
	else
		deliveryarea = "KR"
		ems   = ""
		epostmilitary   = ""
	end if
end if

''�ӽ�..
'if (research="") then
'    notitemlist = "311341"
'    notitemlistinclude="on"
'end if

if research="" then
	'notitemlist = "45718"
	''if notitemlist="" then notitemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	''if itemlist="" then itemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	'if notitemlistinclude="" then notitemlistinclude="on"
end if

if (pagesize="") then
	pagesize = request.cookies("offlinebaljupagesize")
end if

if (pagesize="") then pagesize=1000

response.cookies("offlinebaljupagesize") = pagesize

if (research = "") then
	notitemlistinclude = "on"
	notbrandlistinclude = "on"

	''includeminus		= "N"
	includezerostock	= "N"
end if

set ojumun = new CTenBaljuOffline
	ojumun.FRectCompanyId = companyid
	ojumun.FRectLocationid3PL = locationid3PL
	ojumun.FRectLocationidTo = locationidto
	ojumun.FPageSize = pagesize

	if notitemlistinclude="on" then
		ojumun.FRectNotIncludeItem = "Y"
	else
		ojumun.FRectNotIncludeItem = ""
	end if

	if itemlistinclude="on" then
		ojumun.FRectIncludeItem = "Y"
	else
		ojumun.FRectIncludeItem = ""
	end if

	if notbrandlistinclude="on" then
		ojumun.FRectNotIncludebrand = "Y"
	else
		ojumun.FRectNotIncludebrand = ""
	end if

	if brandlistinclude="on" then
		ojumun.FRectIncludebrand = "Y"
	else
		ojumun.FRectIncludebrand = ""
	end if

	ojumun.FCurrPage = page
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1

	''�簡�� ��۱ǿ�
	ojumun.FRectOnlySagawaDeliverArea = sagawa

	''��ü��� ���� �ֹ���.
	ojumun.FRectUpbeaInclude = upbeaInclude

	if tenbeaonly<>"" then
		ojumun.FRectTenbeaOnly = "Y"
	end if

	if deliveryarea<>"" then
		ojumun.FRectDeliveryArea = deliveryarea
	end if

	if bigitem<>"" then
		ojumun.FRectOnlyManyItem = "Y"
	end if

	if onejumuntype<>"" then
		ojumun.FRectOnlyOneJumun = "Y"

		ojumun.FRectOnlyOneJumunType = onejumuntype
		ojumun.FRectOnlyOneJumunCompare = onejumuncompare
		ojumun.FRectOnlyOneJumunCount = onejumuncount
	end if

	ojumun.FRectIncludeMinus		= includeminus
	ojumun.FRectIncludeZeroStock	= includezerostock
	ojumun.FRectMWDiv				= mwdiv
	ojumun.FRectMakerid				= makerid
	ojumun.GetBaljuItemListNewOffline

dim tenbaljucount
tenbaljucount =0

dim iStartDate, iEndDate

iStartDate  = Left(CStr(DateAdd("d",now(),-2)),10)
iEndDate    = Left(CStr(DateAdd("d",now(),+4)),10)

%>

<script language='javascript'>

var tenBaljuCnt = 0;

function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;
    var isEmsBalju = <%= chkIIF(ems="on","true","false")%>;
    var isMilitaryBalju = <%= chkIIF(epostmilitary="on","true","false")%>;
    var locationidto;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('��� �ù�縦 ���� �ϼ���.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('�۾� �׷��� ���� �ϼ���.');
		document.all.groupform.workgroup.focus();
		return;
	}

	if (document.all.frm.deliveryarea.value.length<1){
		alert('��������� ���� �ϼ���.');
		document.all.frm.deliveryarea.focus();
		return;
	}

	if (document.all.frm.deliveryarea.value == "EMS") {
		if (document.all.groupform.songjangdiv.value != "90"){
			if (confirm("�ؿ���� > EMS �̿� �ù�� ����!!\n\n�״�� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		}
	}

	if (document.all.frm.deliveryarea.value != "EMS") {
		if (document.all.groupform.songjangdiv.value == "90"){
			alert('�ؿܹ�۸� EMS ����� ������ �� �ֽ��ϴ�.');
			document.all.frm.deliveryarea.focus();
			return;
		}
	}

    if (document.all.groupform.pickingStationCd.value == '') {
        alert('��ŷ�����̼��� �����ϼ���.');
        return;
    }

	// ========================================================================
	var iszerobaljono, currordercode;

	upfrm.masteridx.value = "";
	upfrm.ordercode.value = "";
	upfrm.detailidx.value = "";
	upfrm.baljuno.value = "";

	upfrm.comment.value = "";
	upfrm.errorcd.value = "";

	currordercode = "";
	iszerobaljono = true;
	locationidto = "";
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if ((currordercode != "") && (currordercode != frm.ordercode.value)) {
					if (iszerobaljono == true) {
						alert("������ü����� ���� �ֹ��� �ֽ��ϴ�.[" + currordercode + "]");
						return;
					}
					iszerobaljono = true;
				}
				currordercode = frm.ordercode.value;

				if (locationidto == "") {
					locationidto = frm.locationidto.value;
				} else {
					if (locationidto != frm.locationidto.value) {
						alert("�������ֹ��� �ѹ��� ���ÿ� ��������� �� �����ϴ�.");
						return;
					}
				}

				upfrm.masteridx.value = upfrm.masteridx.value + "|" + frm.masteridx.value;
				upfrm.ordercode.value = upfrm.ordercode.value + "|" + frm.ordercode.value;
				upfrm.detailidx.value = upfrm.detailidx.value + "|" + frm.detailidx.value;
				upfrm.baljuno.value = upfrm.baljuno.value + "|" + frm.baljuno.value;

				upfrm.comment.value = upfrm.comment.value + "|" + frm.comment.value;
				upfrm.errorcd.value = upfrm.errorcd.value + "|" + frm.errorcd.value;

				if (frm.baljuno.value*1 != 0) {
					iszerobaljono = false;
				}
			}
		}
	}
	if (iszerobaljono == true) {
		alert("������ü����� ���� �ֹ��� �ֽ��ϴ�.[" + currordercode + "]");
		return;
	}
	upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
	upfrm.workgroup.value = document.all.groupform.workgroup.value;
    upfrm.pickingStationCd.value = document.all.groupform.pickingStationCd.value;
	upfrm.ems.value = "<%= ems %>";
	upfrm.epostmilitary.value = "<%= epostmilitary %>";

	//var count = (upfrm.masteridx.value.match(/\|/g) || []).length;
	//if (count > 410) {
	//	alert("�ʹ� ���� �ֹ��� �����߽��ϴ�. 400�� ���Ϸ� �ֹ����� �� ��������ϼ���.");
	//	return;
	//}

	// ========================================================================
	var ret = confirm('���� �ֹ��� �� ������ü��� �����Ͻðڽ��ϱ�?');
	if (ret) {
		upfrm.submit();
	}
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnableDiable(icomp){
	//return;
	var frm = document.frm;
	var ischecked = icomp.checked;
	if (ischecked){
		if (icomp.name=="notitemlistinclude"){
			frm.itemlistinclude.checked = !(ischecked);
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
		}

	}

	if (ischecked){
		if (icomp.name=="notbrandlistinclude"){
			frm.brandlistinclude.checked = !(ischecked);
		}else if (icomp.name=="brandlistinclude"){
			frm.notbrandlistinclude.checked = !(ischecked);
		}

	}

	if (icomp.name=="onlyOne"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function poponebrand(){
	var popwin = window.open("poponebrand.asp","poponebrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popViewOrderSheet(idx){
	var popwin = window.open("/admin/fran/jumuninputedit.asp?idx=" + idx,"popViewOrderSheet","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popViewRelatedOrderSheet(itemgubun, itemid, itemoption) {
	var popwin = window.open("popBaljuRelatedOrderList.asp?itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"popViewRelatedOrderSheet","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkUpbea(){
    var frm;
    var checkedExists = false;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.tenbeaexists.value!="Y"){
			    frm.cksel.checked = true;
			    AnCheckClick(frm.cksel);
			    checkedExists = true;
			}
		}
	}

	if (checkedExists){
	    document.groupform.songjangdiv.value="24";
	    document.groupform.workgroup.value="Z";
	    CheckNBalju();
	}
}

function chkAllitem(masteridx, ischecked) {
    var frm;

	for (var i = 0; ; i++) {
		frm = document.getElementById("frmBuyPrc_" + i);

		if (frm === null) {
			break;
		}

		if (frm.masteridx.value*1 == masteridx){
			frm.cksel.checked = ischecked;
			AnCheckClick(frm.cksel);
		}
	}
}

function ckAllLimit(icomp, maxnum) {
	var bool = icomp.checked;
	var frm, currnum = 0;
	var alartShowed = false;

	if (bool === false) {
		AnSelectAllFrame(bool);
	} else {
		for (var i = 0; i < document.forms.length; i++) {
			frm = document.forms[i];
			if (frm.name.substr(0,9) == "frmBuyPrc") {
				if (frm.cksel.disabled != true) {
					//if (currnum >= maxnum) {
					//	if (alartShowed == false) {
					//		alert("\n\n�ѹ��� " + maxnum + "�� �̻��� �����Ͽ� ��������� �� �����ϴ�.\n\n");
					//		alartShowed = true;
					//	}
					//	frm.cksel.disabled = true;
					//} else {
						if (frm.cksel.checked === true) {
							currnum = currnum + 1;
						} else {
							chkAllitem(frm.masteridx.value, bool);
							currnum = currnum + 1;
						}
					//}
				}
			}
		}
	}
	// AnSelectAllFrame(bool);
}

// �ֹ����� & ������ü��� ������ �׼�
function chkRealItemNo(tn){
	var frm = eval("frmBuyPrc_"+ tn);
	var v = frm.baljuno;
	var vv = frm.requestedno;

	if (isNaN(v.value)||v.value.length<1){
		return;
	}else{
		v.value = parseInt(v.value);
	}

	if (eval("seldiv" + tn) == false) {
		// �������ɻ��� �ƴ�.
		return;
	}

	// �ֹ������� ������ü����� �ٸ��� ǥ��
	if(parseInt(v.value) != parseInt(vv.value)) {
		ShowHide("seldiv" + tn, true);
		fnselcom(eval("frmBuyPrc_" + tn + ".dtstat.value"), tn);
	} else {
		ShowHide("seldiv" + tn, false);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "";
		frm.dtstat.selectedIndex = 0;
		frm.comment.readOnly = true;
		frm.errorcd.value = "";
	}
}

function ShowHide(divid, isshow) {

	if (isshow == true) {
		document.getElementById(divid).style.display='';
	} else {
		document.getElementById(divid).style.display='none';
	}

}

//������ ǥ��
function fnselcom(val, tn) {

	var frm = eval("frmBuyPrc_"+ tn);

	if(val=='ipt'){
		// ====================================================================
		// �����Է�
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, true);

		frm.comment.value = "";
		frm.comment.readOnly = false;

		frm.errorcd.value = "C";
	}else if(val=='sso'){
		// ====================================================================
		// �Ͻ�ǰ��
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "�Ͻ�ǰ��";
		frm.comment.readOnly = true;

		frm.errorcd.value = "T";
	}else if(val=='5day'){
		// ====================================================================
		// 5�ϳ����
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "5�ϳ����";
		frm.comment.readOnly = true;

		frm.errorcd.value = "C";
	}else if(val=='jaego'){
		// ====================================================================
		// ������
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "������";
		frm.comment.readOnly = true;

		frm.errorcd.value = "C";
	}else if(val=='so'){
		// ====================================================================
		// ����
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "����";
		frm.comment.readOnly = true;

		frm.errorcd.value = "E";
	}else{
		// ====================================================================
		// ����
		ShowHide("seldiv" + tn, false);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "";
		frm.comment.readOnly = true;

		frm.errorcd.value = "";
	}
}

</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="companyid" value="<%= companyid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<b>��ü</b> : <b><%= companyid %></b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<b>�������</b> :
		<select name="deliveryarea" >
			<option value="" 	<% if deliveryarea="" then response.write "selected" %> >��ü</option>
			<option value="KR" 	<% if deliveryarea="KR" then response.write "selected" %> >�������</option>
			<option value="EMS" <% if deliveryarea="EMS" then response.write "selected" %> >�ؿܹ��</option>
			<!--
			<option value="ZZ" 	<% if deliveryarea="ZZ" then response.write "selected" %> >���δ���</option>
			-->
		</select>
		<input type="checkbox" name="includeminus" value="N" <% if (includeminus = "N") then %>checked<% end if %>> ���̳ʽ��ֹ�����
		<input type="checkbox" name="includezerostock" value="N" <% if (includezerostock = "N") then %>checked<% end if %>> ��ü�������ֹ�����
		<input type="checkbox" name="includeminu11s" value="N"> �¶���7���Ǹź����ܼ�����
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<b>�Ⱓ : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ ����
		&nbsp;
		���� :
		<% 'drawSelectBoxOffShop "locationidto",locationidto %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("locationidto",locationidto, "21") %>
		&nbsp;
		�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		��۱��� :
		<select class="select" name="mwdiv">
			<option value="">��ü</option>
			<option value="U" <% if mwdiv="U" then response.write "selected" %> >����</option>
			<option value="T" <% if mwdiv="T" then response.write "selected" %> >�ٹ�</option>
			<option value="O" <% if mwdiv="O" then response.write "selected" %> >����</option>
		</select>
		<!--
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="bigitem" <% if bigitem="on" then response.write "checked" %> > <b>�ټ���ǰ�ֹ�</b>
		<font color="#AAAAAA">
		<input type="checkbox" name="upbeaInclude" <% if upbeaInclude="on" then response.write "checked" %> > <b>�������� �ֹ��Ǹ�</b>
		</font>
		<input type="checkbox" name="tenbeaonly" <% if tenbeaonly="on" then response.write "checked" %> > <b>�ٹ��ֹ��Ǹ�</b>
		<input type="checkbox" name="notitemlistinclude" <% if notitemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>��Ź��ǰ ���� �ֹ���</b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="itemlistinclude" <% if itemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>��Ź��ǰ ���� �ֹ���</b>
		<input type="checkbox" name="notbrandlistinclude" <% if notbrandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>��Ź����ó ���� �ֹ���</b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="brandlistinclude" <% if brandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>��Ź����ó ���� �ֹ���</b>
		<b>��ǰ�ֹ�</b> :
		<select name="onejumuntype" >
		<option value="" 	<% if onejumuntype="" then response.write "selected" %> >========</option>
		<option value="all" <% if onejumuntype="all" then response.write "selected" %> >��� ��ǰ�ֹ�</option>
		<option value="reg" <% if onejumuntype="reg" then response.write "selected" %> >������ ��ǰ�ֹ�</option>
		</select>

		<input type="text" name="onejumuncount" value="<%= onejumuncount %>" size=3>
		<select name="onejumuncompare" >
		<option value="less" 	<% if onejumuncompare="less" then response.write "selected" %> >�� ����</option>
		<option value="more" 	<% if onejumuncompare="more" then response.write "selected" %> >�� �̻�</option>
		<option value="equal" 	<% if onejumuncompare="equal" then response.write "selected" %> >��</option>
		</select>
		<input type="checkbox" name="onlyOne" <% if onlyOne="on" then response.write "checked" %> onclick="EnableDiable(this);">
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="danpumcheck" <% if danpumcheck="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<input type="button" value="����/����/��ǰ ��ǰ����" onclick="javascript:poponeitem();">
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="button" value="����/���� ����ó����" onclick="javascript:poponebrand();">
		<input type="text" name="dcnt2" value="<%= dcnt2 %>" size=1> �� (11 �Է½� 11�� �̻�, 0�� �Է½� 0�� �̻�)
		<input type="checkbox" name="ems" <% if ems="on" then response.write "checked" %> > <b>�ؿܹ��</b>
		<input type="checkbox" name="epostmilitary" <% if epostmilitary="on" then response.write "checked" %> > <b>���δ�</b>
		<input type="checkbox" name="imsi" <% if imsi="on" then response.write "checked" %> > <b>�ӽ�(���ѵ��� ���� ����)</b>
		<font color="#AAAAAA">
		<input type="checkbox" name="sagawa" <% if sagawa="on" then response.write "checked" %> onClick="alert('�Ϲ�������ø� ���� (��ǰ���,��Ź��ǰ �˻� ����ȵ�)');"> �ӽ�(�簡�ͱǿ�)
		</font>
		-->
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="groupform">
<tr>
	<td align="left">
        �� ��������� �Ǽ� : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
		�� �ݾ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
		��հ��ܰ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font><br>
		* �������� �ֹ����� ���� ��ǰ�� �ֹ��� ��� ����ɼ����� �޶����� �Ѵ�.
	</td>
	<td align="right">
	    <!--
	    <select name="baljutype">
	        <option value="">�Ϲ�
	        <option value="D">DAS
	        <option value="S">��ǰ���
	    </select>
	    -->
	    <select name="songjangdiv">
	        <option value="">�ù�缱��</option>
			<option value="1" >�����ù�</option>
			<option value="2" >�Ե��ù�</option>
            <option value="4" >CJ�ù�</option>
		   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS</option>
			<option value="91" >DHL</option>
		   	<option value="98" >������</option>
		   	<option value="99" >��Ÿ</option>
		   	<!--
		   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >��ü��(���δ�)
		   	-->
	    </select>
		<select name="workgroup">
		   	<option value="">�۾��׷�
		   	<option value="O" >O(��������)
	   	</select>
        <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", "") %>
		<!--
		<select name="workgroup">
		   	<option value="">�۾��׷�
		   	<option value="A" >A
		   	<option value="B" >B
		   	<option value="C" >C(DAS)
		   	<option value="D" >D
		   	<option value="F" >F
		   	<option value="" >===========
		   	<option value="T" >T(Ž������)
		   	<option value="" >===========
		   	<option value="I" >I(���̶��)
		   	<option value="" >===========
		   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)
		   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(���δ�)
		   	<option value="Z" >Z(����)
	   	</select>
	   	-->
		<input type="button" value="���û���������ü��ۼ�" onclick="CheckNBalju()" class="button">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<br>



<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22">
        <div id="currsearchno">�� �˻��ֹ��Ǽ� : </div>
        <!--
        <input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> �����������
        -->
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="ckAllLimit(this, 400)"></td>
	<td width="80">�����ڵ�</td>
	<td width="40">����</td>
	<td>UserID<br>���̸�</td>
	<td width="100">�����ڵ�</td>
	<td>�귣��ID</td>
	<td align="left">��ǰ��<br><font color=blue>[�ɼǸ�]</font></td>
	<td width="40">���<br>����</td>
	<!--
	<td width="60">�ǻ�<br>����</td>
	<td width="60">���������<br>(ON+OFF)</td>
	<td width="60">�¶���<br>����</td>
	<td width="60">N���ʿ�<br>����</td>
	<td width="60">�ֹ�<br>����</td>
	<td width="60">�����<br>����</td>
	-->

	<td width="60">�ǻ�<br>��ȿ���</td>
	<td width="60">ON<br>��ǰ�غ�</td>
	<td width="60">OFF<br>��ǰ�غ�</td>
	<td width="60">ON<br>�����Ϸ�</td>
	<td width="60">ON<br>�ֹ�����</td>
	<td width="60">���<br>���ɼ���</td>
	<td width="60">�ֹ�����</td>

	<td width="60">�������<br>����</td>
	<td width="240">���</td>
	<td width="40">����<br />�ֹ�</td>
</tr>
<% if ojumun.FresultCount>0 then %>
<% for ix=0 to ojumun.FresultCount-1 %>
<form name="frmBuyPrc_<%= ix %>" id="frmBuyPrc_<%= ix %>" method="post" >
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(ix).Fmasteridx %>">
<input type="hidden" name="ordercode" value="<%= ojumun.FItemList(ix).Fordercode %>">
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(ix).Fdetailidx %>">
<input type="hidden" name="locationidto" value="<%= ojumun.FItemList(ix).Flocationidto %>">
<tr align="center" bgcolor="#FFFFFF">
    <% if ((ems<>"") or (epostmilitary<>"")) then %>
    	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this); chkAllitem(<%= ojumun.FItemList(ix).Fmasteridx %>, this.checked)"></td>
    <% else %>
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this); chkAllitem(<%= ojumun.FItemList(ix).Fmasteridx %>, this.checked)" <%= CHKIIF(ojumun.FItemList(ix).FcountryCode<>"" and ojumun.FItemList(ix).FcountryCode<>"KR","disabled","") %> ></td>
	<% end if %>
		<td><a href="javascript:popViewOrderSheet(<%= ojumun.FItemList(ix).Fmasteridx %>)"><%= ojumun.FItemList(ix).Fordercode %></a></td>
	<td><%= ojumun.FItemList(ix).FcountryCode %></td>
	<td><%= ojumun.FItemList(ix).Flocationidto %><br><%= ojumun.FItemList(ix).Flocationnameto %></td>
	<td><%= ojumun.FItemList(ix).Fprdcode %></td>
	<td><%= ojumun.FItemList(ix).Fbrandid %></td>
	<td align="left">
		<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= BF_GetItemGubun(ojumun.FItemList(ix).Fprdcode) %>&itemid=<%= BF_GetItemId(ojumun.FItemList(ix).Fprdcode) %>&itemoption=<%= BF_GetItemOption(ojumun.FItemList(ix).Fprdcode) %>" target=_blank >
			<%= ojumun.FItemList(ix).Fprdname %><br><font color=blue>[<%= ojumun.FItemList(ix).Fitemoptionname %>]</font>
		</a>

	</td>
	<td>
		<% if ojumun.FItemList(ix).Fmwdiv = "M" or ojumun.FItemList(ix).Fmwdiv = "W" then %>
			<%= ojumun.FItemList(ix).GetMWDivName %>
		<% elseif ojumun.FItemList(ix).Fmwdiv = "U" then %>
			<font color="red"><%= ojumun.FItemList(ix).GetMWDivName %></font>
		<% elseif ojumun.FItemList(ix).Fmwdiv = "O" then %>
			<font color="blue"><%= ojumun.FItemList(ix).GetMWDivName %></font>
		<% end if %>
	</td>
	<!--
	<td><%= ojumun.FItemList(ix).Frealstockno %></td>
	<td><%= (ojumun.FItemList(ix).Fmoveoutdiv5 + ojumun.FItemList(ix).Fmoveoutdiv7 + ojumun.FItemList(ix).Fselldiv5 + ojumun.FItemList(ix).Fchulgodiv5) %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv4 %></td>
	<td><%= ojumun.FItemList(ix).GetOnlineRequireNo %></td>
	<td><%= ojumun.FItemList(ix).Frequestedno %></td>
	<td><%= ojumun.FItemList(ix).GetChulgoAvailableNo %></td>
	-->

	<input type="hidden" name="requestedno" value="<%= ojumun.FItemList(ix).Frequestedno %>">
	<td><%= ojumun.FItemList(ix).Frealstockno %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv5 %></td>
	<td><%= ojumun.FItemList(ix).Fchulgodiv5 %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv4 %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv2 %></td>
	<td>
		<b><%= ojumun.FItemList(ix).GetOffChulgoAvailableNo %></b>
	</td>
	<td>
		<b><font color="<%= CHKIIF((ojumun.FItemList(ix).GetOffChulgoAvailableNo-ojumun.FItemList(ix).Frequestedno) < 0, "red", "black") %>"><%= ojumun.FItemList(ix).Frequestedno %><%= CHKIIF(ojumun.FItemList(ix).Frequestedno<>ojumun.FItemList(ix).Foffjupno*-1, "/" & ojumun.FItemList(ix).Foffjupno*-1, "") %></font></b>
	</td>

	<td>
		<% if (ojumun.FItemList(ix).Frequestedno > 0) and (ojumun.FItemList(ix).GetOffChulgoAvailableNo > 0) then %>
			<% if ojumun.FItemList(ix).Frequestedno > ojumun.FItemList(ix).GetOffChulgoAvailableNo then %>
				<input type=text class="text" name=baljuno size=4 value="<%= ojumun.FItemList(ix).GetOffChulgoAvailableNo %>"  onKeyup="chkRealItemNo(<%= ix %>);">
			<% else %>
				<input type=text class="text" name=baljuno size=4 value="<%= ojumun.FItemList(ix).Frequestedno %>"  onKeyup="chkRealItemNo(<%= ix %>);">
			<% end if %>
		<% else %>
			<input type=text class="text" name=baljuno size=4 value="0"  onKeyup="chkRealItemNo(<%= ix %>);">
		<% end if %>
	</td>
	<input type="hidden" name="errorcd" value="">
	<td>
		<span id="seldiv<%= ix %>">
				<select class="select" name="dtstat" onchange="fnselcom(this.value,<%= ix %>);">
					<option value="ipt">�����Է�</option>
					<option value="5day" <%= CHKIIF(ojumun.FItemList(ix).Fdanjongyn="M", "selected", "") %> >5�ϳ����</option>
					<option value="jaego" <%= CHKIIF((ojumun.FItemList(ix).Fdanjongyn="S" or ojumun.FItemList(ix).Fdanjongyn = "Y"), "selected", "") %> >������</option>
					<option value="so">����</option>
					<option value="sso">�Ͻ�ǰ��</option>
				</select>
		</span>
		<script>ShowHide("seldiv<%= ix %>", false)</script>
		<span id="comdiv<%= ix %>">
			<input type="text" class="text" name="comment" value="" size="8"><br>
		</span>
		<script>ShowHide("comdiv<%= ix %>", false)</script>

		<% if ((Not IsNull(ojumun.FItemList(ix).FreipgoMayDate)) and (Left(ojumun.FItemList(ix).FreipgoMayDate, 10) >= iStartDate) and (Left(ojumun.FItemList(ix).FreipgoMayDate, 10) <= iEndDate)) then %>
			<%= Left(ojumun.FItemList(ix).FreipgoMayDate, 10) %>
		<% end if %>
		<%= fnColor(ojumun.FItemList(ix).Fdanjongyn,"dj") %>
		<% if ojumun.FItemList(ix).Fpreorderno<>0 then %>
			���ֹ�:
			<% if ojumun.FItemList(ix).Fpreorderno<>ojumun.FItemList(ix).Fpreordernofix then response.write CStr(ojumun.FItemList(ix).Fpreorderno) + "->" %>
			<%= ojumun.FItemList(ix).Fpreordernofix %>
		<% end if %>
	</td>
	<td><a href="javascript:popViewRelatedOrderSheet('<%= ojumun.FItemList(ix).Fitemgubun %>', <%= ojumun.FItemList(ix).Fitemid %>, '<%= ojumun.FItemList(ix).Fitemoption %>')">����</a></td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<form name="frmArrupdate" method="post" action="baljumaker_offline_new_process.asp">
	<input type="hidden" name="mode" value="arr">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="ordercode" value="">
	<input type="hidden" name="detailidx" value="">
	<input type="hidden" name="baljuno" value="">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="errorcd" value="">
	<input type="hidden" name="songjangdiv" value="">
	<input type="hidden" name="workgroup" value="">
	<input type="hidden" name="ems" value="">
	<input type="hidden" name="epostmilitary" value="">
    <input type="hidden" name="pickingStationCd" value="">
</form>
</table>

<script language='javascript'>

<% ''if (locationidto <> "") then %>

	for (var i = 0; i < <%= ojumun.FresultCount %>; i++) {
		chkRealItemNo(i);
	}

<% ''end if %>

	document.all.currsearchno.innerHTML = "�˻����� : <Font color='#3333FF'><%= ix %></font>";
	tenBaljuCnt = 1*<%= tenbaljucount %>;

	<% if onlyOne<>"" then %>
		EnableDiable(frm.onlyOne);
	<% end if %>

	<% if danpumcheck<>"" then %>
		EnableDiable(frm.danpumcheck);
	<% end if %>



</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
