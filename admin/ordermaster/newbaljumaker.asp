<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ü� �ۼ�
' History : �̻� ����
'			2018.03.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->

<%
DIM CBRAND_INEXCLUDE_USING : CBRAND_INEXCLUDE_USING = True
Dim FlushCount : FlushCount=100  ''2016/04/18 :: ASP �������� �����Ͽ� Response ������ ������ ������ �ʰ��Ǿ����ϴ�.

dim pagesize, notitemlist, itemlist, notitemlistinclude, itemlistinclude, notbrandlistinclude, brandlistinclude
dim research, yyyy1,mm1,dd1,yyyymmdd,nowdate, onlyOne,dcnt, danpumcheck, upbeaInclude, dcnt2, searchtypestring
dim imsi, sagawa, ems, epostmilitary, bigitem, fewitem, kpack, deliveryarea, onejumuntype, onejumuncount, onejumuncompare
dim tenbeaonly, tenbeamakeonorder, cn10x10, ecargo, extSiteName, stockLocationGubun, excMinusStock, excRealMinusStock
dim presentOnly, show100, repeatOrderCnt, standingorderinclude, page, ojumun, ix,iy, tenbaljucount

	standingorderinclude = requestcheckvar(request("standingorderinclude"),2)
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

pagesize = request("pagesize")

deliveryarea = request("deliveryarea")

bigitem = request("bigitem")
fewitem = request("fewitem")

upbeaInclude = request("upbeaInclude")

notitemlistinclude = request("notitemlistinclude")
itemlistinclude = request("itemlistinclude")

notbrandlistinclude = request("notbrandlistinclude")
brandlistinclude = request("brandlistinclude")

notitemlist = request("notitemlist")
itemlist = request("itemlist")

research = request("research")

onejumuntype = request("onejumuntype")
onejumuncount = request("onejumuncount")
onejumuncompare = request("onejumuncompare")

tenbeaonly = request("tenbeaonly")

tenbeamakeonorder = request("tenbeamakeonorder")

extSiteName = request("extSiteName")

stockLocationGubun = request("stockLocationGubun")
excMinusStock = request("excMinusStock")
excRealMinusStock = request("excRealMinusStock")
presentOnly = request("presentOnly")
show100 = request("show100")
repeatOrderCnt = request("repeatOrderCnt")

if (research = "") then
	notitemlistinclude = "on"
	if (CBRAND_INEXCLUDE_USING) then
	    notbrandlistinclude = "on"
    end if
	tenbeamakeonorder = "E"
	extSiteName = "10x10"
	presentOnly = "N"
	show100 = "Y"
end if

if (repeatOrderCnt = "") then
	repeatOrderCnt = "0"
end if

'dcnt = trim(request("dcnt"))
'dcnt2 = trim(request("dcnt2"))
'onlyOne = request("onlyOne")
'danpumcheck = request("danpumcheck")
'imsi  = request("imsi")
'sagawa= request("sagawa")
'ems   = request("ems")
'epostmilitary   = request("epostmilitary")



'==============================================================================
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


'// ============================================================================
ems = ""
kpack = ""
epostmilitary = ""
cn10x10 = ""
ecargo = ""
if deliveryarea<>"" then
	if (deliveryarea = "ZZ") then
		'// ���δ�
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		'// �ؿܹ��
		ems   = "on"
	elseif (deliveryarea = "KPACK") then
		'// �ؿܹ�� kpack
		kpack   = "on"
	elseif (deliveryarea = "CN10X10") then
		'// �߱������
		cn10x10 = "on"
	elseif (deliveryarea = "ECARGO") then
		'// ��ī����
		ecargo = "on"
	elseif (deliveryarea = "QQ") then
		'// �����
	else
		'// ��Ÿ : �������
		deliveryarea = "KR"
	end if
end if


'// ============================================================================
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

''��Ű ���� /2016/04/18
''if (pagesize="") then
''	pagesize = request.cookies("baljupagesize")
''end if

if (pagesize="") then pagesize=200
''if (pagesize>=2000) then pagesize=1000

''��Ű ���� /2016/04/18
''response.cookies("baljupagesize") = pagesize

page = request("page")
if (page="") then page=1

set ojumun = new CTenBalju

''�� ����¡�� 2�� �˻�
ojumun.FPageSize = pagesize * 3
''ojumun.FPageSize = pagesize

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

''��ü��� ���� �ֹ���.
ojumun.FRectUpbeaInclude = upbeaInclude

''�簡�� ��۱ǿ�
ojumun.FRectOnlySagawaDeliverArea = sagawa

if deliveryarea<>"" then
	ojumun.FRectDeliveryArea = deliveryarea
end if

if fewitem<>"" then
	ojumun.FRectOnlyFewItem = fewitem
end if

if onejumuntype<>"" then
	ojumun.FRectOnlyOneJumun = "Y"

	ojumun.FRectOnlyOneJumunType = onejumuntype
	ojumun.FRectOnlyOneJumunCompare = onejumuncompare
	ojumun.FRectOnlyOneJumunCount = onejumuncount
end if

if tenbeaonly<>"" then
	ojumun.FRectTenbeaOnly = "Y"
end if

if tenbeamakeonorder <> "" then
	ojumun.FRectTenbeaMakeOnOrder = tenbeamakeonorder
end if


ojumun.FRectSiteGubun = extSiteName

ojumun.FRectStockLocationGubun = stockLocationGubun
ojumun.FRectExcMinusStock = excMinusStock
ojumun.FRectExcRealMinusStock = excRealMinusStock
ojumun.FRectPresentOnly = presentOnly
ojumun.FRectRepeatOrderCnt = repeatOrderCnt
ojumun.FRectstandingorderinclude = standingorderinclude
ojumun.GetBaljuItemListProc
''ojumun.GetBaljuItemListNew

tenbaljucount =0

dim MaxTenBaljuCount : MaxTenBaljuCount = 100

%>
<script language='javascript'>
var tenBaljuCnt = 0;
function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;
    var isEmsBalju = <%= chkIIF(ems="on","true","false")%>;
    var isKpackBalju = <%= chkIIF(kpack="on","true","false")%>;
    var isMilitaryBalju = <%= chkIIF(epostmilitary="on","true","false")%>;
	var isCn10x10Balju = <%= chkIIF(cn10x10="on","true","false")%>;
	var isEcargoBalju = <%= chkIIF(ecargo="on","true","false")%>;
	var isQuickBeasongBalju = <%= chkIIF(deliveryarea="QQ","true","false")%>;

	var isGiftPojang = document.frm.presentOnly[1].checked ? 'Y' : 'N';
	upfrm.isGiftPojang.value = isGiftPojang;

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

	/*
    //C�۾��� DAS
    isDasBalju = (document.all.groupform.workgroup.value=="C");

    //DAS ������� üũ, �ٹ� 150�� ����.
    if ((isDasBalju)&&(tenBaljuCnt>150)){
        alert('DAS ������ô� �ٹ����� ��� 150�� �̸��� �����մϴ�. ');
		document.all.groupform.workgroup.focus();
		return;
    }
	 */

	// ========================================================================
    if (isEmsBalju){
        if (document.all.groupform.workgroup.value!="E"){
            alert('EMS(�ؿ�)������ô� E �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="E"){
            alert('�˻������� �ؿܹ���̾�� EMS(�ؿ�)������ð� �����մϴ�.');
            return;
        }
    }
    if (isKpackBalju){
        if (document.all.groupform.workgroup.value!="R"){
            alert('KPACK(�ؿ�)������ô� R �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="R"){
            alert('�˻������� �ؿܹ���̾�� KPACK(�ؿ�)������ð� �����մϴ�.');
            return;
        }
    }

	if ((isQuickBeasongBalju == true) && (document.all.groupform.songjangdiv.value != "98")) {
		alert('�ù�縦 ��������� �����ϼ���.');
           return;
	}

    if (((document.all.groupform.songjangdiv.value=="90")&&(document.all.groupform.workgroup.value!="E"))||((document.all.groupform.songjangdiv.value!="90")&&(document.all.groupform.workgroup.value=="E"))){
        alert('EMS(�ؿ�)������ô� E �۾��常 �����մϴ�.');
        return;
    }

    if (((document.all.groupform.songjangdiv.value=="93")&&(document.all.groupform.workgroup.value!="R"))||((document.all.groupform.songjangdiv.value!="93")&&(document.all.groupform.workgroup.value=="R"))){
        alert('KPACK(�ؿ�)������ô� R �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isMilitaryBalju){
        if (document.all.groupform.workgroup.value!="G"){
            alert('���δ� ������ô� G �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="G"){
            alert('�˻������� ���δ����̾�� ���δ�������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="8")&&(document.all.groupform.workgroup.value!="G"))||((document.all.groupform.songjangdiv.value!="8")&&(document.all.groupform.workgroup.value=="G"))){
        alert('���δ� ������ô� G �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isCn10x10Balju){
        if (document.all.groupform.workgroup.value!="H"){
            alert('EMS(�߱���)������ô� H �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="H"){
            alert('��������� �߱�������̾�� EMS(�߱���)������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="91")&&(document.all.groupform.workgroup.value!="H"))||((document.all.groupform.songjangdiv.value!="91")&&(document.all.groupform.workgroup.value=="H"))){
        alert('EMS(�߱���)������ô� H �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isEcargoBalju){
        if (document.all.groupform.workgroup.value!="J"){
            alert('�ؿ�(ECARGO)������ô� J �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="J"){
            alert('��������� �ؿ�(ECARGO)����̾�� �ؿ�(ECARGO)������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="92")&&(document.all.groupform.workgroup.value!="J"))||((document.all.groupform.songjangdiv.value!="92")&&(document.all.groupform.workgroup.value=="J"))){
        alert('�ؿ�(ECARGO)������ô� J �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isDasBalju){
        if (!confirm('DAS ������� �Դϴ�. ��� �Ͻðڽ��ϱ�?')){
            return;
        }
    }
	var ret = confirm('���� �ֹ��� �� ������ü��� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
		upfrm.workgroup.value = document.all.groupform.workgroup.value;
		upfrm.ems.value = "<%= ems %>";
		upfrm.kpack.value = "<%= kpack %>";
		upfrm.epostmilitary.value = "<%= epostmilitary %>";
		upfrm.cn10x10.value = "<%= cn10x10 %>";
		upfrm.ecargo.value = "<%= ecargo %>";
		upfrm.extSiteName.value = "<%= extSiteName %>";

		if (isDasBalju) {
		    upfrm.baljutype.value = "D";
		}else{
		    upfrm.baljutype.value = "";
		}

		if ((upfrm.songjangdiv.value == "91") || (upfrm.songjangdiv.value == "92") || (upfrm.songjangdiv.value == "93")) {
			upfrm.songjangdiv.value = "90";
		}

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
	var popwin = window.open("/admin/ordermaster/poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function poponebrand(){
	var popwin = window.open("/admin/ordermaster/poponebrand.asp","poponebrand","width=800 height=600 scrollbars=yes resizable=yes");
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
</script>

<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
        		<tr height="35">
        			<td width="320"><b>�Ⱓ : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ ����</td>
        			<td width="220">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>�ٹ����ٹ�� �Ǽ�</b> :
						<select name="pagesize" >
						<option value="10" <% if pagesize="10" then response.write "selected" %> >10</option>
						<option value="20" <% if pagesize="20" then response.write "selected" %> >20</option>
						<option value="50" <% if pagesize="50" then response.write "selected" %> >50</option>
						<option value="100" <% if pagesize="100" then response.write "selected" %> >100</option>
						<option value="120" <% if pagesize="120" then response.write "selected" %> >120</option>
						<option value="150" <% if pagesize="150" then response.write "selected" %> >150</option>
						<option value="200" <% if pagesize="200" then response.write "selected" %> >200</option>
						<option value="250" <% if pagesize="250" then response.write "selected" %> >250</option>
						<option value="300" <% if pagesize="300" then response.write "selected" %> >300</option>
						<option value="400" <% if pagesize="400" then response.write "selected" %> >400</option>
						<option value="500" <% if pagesize="500" then response.write "selected" %> >500</option>
						<option value="600" <% if pagesize="600" then response.write "selected" %> >600</option>
						<option value="800" <% if pagesize="800" then response.write "selected" %> >800</option>
						<option value="1000" <% if pagesize="1000" then response.write "selected" %> >1000</option>
						<option value="2000" <% if pagesize="2000" then response.write "selected" %> >2000</option>
						</select>
        			</td>
        			<td width="250">
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>�������</b> :
						<select name="deliveryarea" >
						<option value="" 			<% if deliveryarea="" then response.write "selected" %> >��ü</option>
						<option value="KR" 			<% if deliveryarea="KR" then response.write "selected" %> >�������</option>
						<option value="QQ" 			<% if deliveryarea="QQ" then response.write "selected" %> >�������(�����)</option>
						<option value="EMS" 		<% if deliveryarea="EMS" then response.write "selected" %> >�ؿܹ��(EMS)</option>
						<option value="KPACK" 		<% if deliveryarea="KPACK" then response.write "selected" %> >�ؿܹ��(KPACK)</option>
						<option value="ZZ" 			<% if deliveryarea="ZZ" then response.write "selected" %> >���δ���</option>
						<option value="CN10X10" 	<% if deliveryarea="CN10X10" then response.write "selected" %> >�߱���(CN10X10)</option>
						<option value="ECARGO" 	    <% if deliveryarea="ECARGO" then response.write "selected" %> >�ؿ�(ECARGO)</option>
						</select>
        			</td>
        			<td width="150">

        			</td>
        			<td width="200">
						&nbsp;&nbsp;|&nbsp;&nbsp;
						<b>��ǰ��ġ</b> :
						<select class="select" name="stockLocationGubun">
							<option value="">��ü</option>
							<option value="3" <% if (stockLocationGubun = "3") then %>selected<% end if %> >3����ǰ �ֹ�</option>
							<option value="4" <% if (stockLocationGubun = "4") then %>selected<% end if %> >4����ǰ �ֹ�</option>
						</select>
					</td>
					<td width="250">
						&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="excMinusStock" value="Y" <% if (excMinusStock = "Y") then %>checked<% end if %> > <b>������� ���̳ʽ� �ֹ�����</b>
					</td>
					<td>
						&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="show100" value="Y" <% if (show100 = "Y") then %>checked<% end if %> > <b>�ٹ��ֹ� 100�Ǹ� ǥ��</b>
					</td>
        		</tr>
        		<tr height="35">
        			<td>
						<font color="#AAAAAA">
						<input type="checkbox" name="upbeaInclude" <% if upbeaInclude="on" then response.write "checked" %> > <b>�������� �ֹ��Ǹ�</b>
						</font>
        			</td>
        			<td>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="tenbeaonly" <% if tenbeaonly="on" then response.write "checked" %> > <b>�ٹ��ֹ��Ǹ�</b>
        			</td>
					<td>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
        				<b>�ֹ�����Ʈ</b> :
						<select name="extSiteName" >
							<option value="" 				<% if extSiteName="" then response.write "selected" %> >��ü</option>
							<option value="10x10" 			<% if extSiteName="10x10" then response.write "selected" %> >����(���޸�����)</option>
							<option value="extSiteAll"		<% if extSiteName="extSiteAll" then response.write "selected" %> >���޸���ü(��������)</option>
							<option value="cjmall" 			<% if extSiteName="cjmall" then response.write "selected" %> >���޸�(cjmall)</option>
							<option value="interpark" 		<% if extSiteName="interpark" then response.write "selected" %> >���޸�(interpark)</option>
							<option value="lotteCom" 		<% if extSiteName="lotteCom" then response.write "selected" %> >���޸�(lotteCom)</option>
							<option value="lotteimall" 		<% if extSiteName="lotteimall" then response.write "selected" %> >���޸�(lotteimall)</option>
							<option value="itsSite" 		<% if extSiteName="itsSite" then response.write "selected" %> >���̶��(ITS)</option>
							<option value="etcExtSite" 		<% if extSiteName="etcExtSite" then response.write "selected" %> >��Ÿ���޸�</option>
						</select>
        			</td>
					<td colspan=4>
						&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="excRealMinusStock" value="Y" <% if (excRealMinusStock = "Y") then %>checked<% end if %> > <b>����ľ����  ���̳ʽ� �ֹ����� </b>
					</td>
        		</tr>
        		<tr height="35">
        			<td>
						<input type="checkbox" name="notitemlistinclude" <% if notitemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>��Ź��ǰ ���� �ֹ���</b>
        			</td>
        			<td>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="itemlistinclude" <% if itemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>��Ź��ǰ ���� �ֹ���</b>
        			</td>
        			<td colspan=5>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="standingorderinclude" <% if standingorderinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>���ⱸ����ǰ ���� �ֹ���</b>
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
						<input type="checkbox" name="notbrandlistinclude" <% if notbrandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);" <%=CHKIIF(NOT CBRAND_INEXCLUDE_USING,"disabled","")%> >
						<b>��Ź�귣�� ���� �ֹ���</b>
        			</td>
        			<td colspan=6>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="checkbox" name="brandlistinclude" <% if brandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
						<b>��Ź�귣�� ���� �ֹ���</b>
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
						<input type="radio" name="tenbeamakeonorder" <% if tenbeamakeonorder="E" then response.write "checked" %> value="E">
						<b>�ٹ� �ֹ����� ���� �ֹ���</b>
        			</td>
        			<td colspan=6>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="radio" name="tenbeamakeonorder" <% if tenbeamakeonorder="I" then response.write "checked" %> value="I">
						<b>�ٹ� �ֹ����� ���� �ֹ���</b>
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
						<input type="radio" name="presentOnly" <% if presentOnly="N" then response.write "checked" %> value="N">
						<b>�������� ���� �ֹ���</b>
        			</td>
        			<td colspan=6>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="radio" name="presentOnly" <% if presentOnly="Y" then response.write "checked" %> value="Y">
						<b>�������� ���� �ֹ���</b>
        			</td>
        		</tr>
        		<tr height="35">
        			<td>
        				<b>��ǰ�ֹ�</b> :
						<select name="onejumuntype" >
						<option value="" 	<% if onejumuntype="" then response.write "selected" %> ></option>
						<option value="all" <% if onejumuntype="all" then response.write "selected" %> >��� ��ǰ�ֹ�</option>
						<option value="reg" <% if onejumuntype="reg" then response.write "selected" %> >������ ��ǰ�ֹ�</option>
						</select>

						<input type="text" name="onejumuncount" value="<%= onejumuncount %>" size=3>
						<select name="onejumuncompare" >
						<option value="less" 	<% if onejumuncompare="less" then response.write "selected" %> >�� ����</option>
						<option value="more" 	<% if onejumuncompare="more" then response.write "selected" %> >�� �̻�</option>
						<option value="equal" 	<% if onejumuncompare="equal" then response.write "selected" %> >��</option>
						</select>

						<!--<input type="checkbox" name="onlyOne" <% if onlyOne="on" then response.write "checked" %> onclick="EnableDiable(this);">-->
        			</td>
        			<td colspan=6>
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<!--<input type="checkbox" name="danpumcheck" <% if danpumcheck="on" then response.write "checked" %> onclick="EnableDiable(this);">-->
						<input type="button" value="����/����/��ǰ ��ǰ����" onclick="javascript:poponeitem();">
						&nbsp;&nbsp;|&nbsp;&nbsp;
						<input type="button" value="����/���� �귣�弳��" onclick="javascript:poponebrand();" <%=CHKIIF(NOT CBRAND_INEXCLUDE_USING,"disabled","")%>>
						<!--<input type="text" name="dcnt2" value="<%= dcnt2 %>" size=1> �� (11 �Է½� 11�� �̻�, 0�� �Է½� 0�� �̻�)-->
						|
						* ���ϻ�ǰ
						<select class="select" name="repeatOrderCnt">
							<option value="0" <%= CHKIIF(repeatOrderCnt="0", "selected", "") %> >��ü</option>
							<option value="2" <%= CHKIIF(repeatOrderCnt="2", "selected", "") %> >2ȸ</option>
							<option value="3" <%= CHKIIF(repeatOrderCnt="3", "selected", "") %> >3ȸ</option>
							<option value="5" <%= CHKIIF(repeatOrderCnt="5", "selected", "") %> >5ȸ</option>
							<option value="10" <%= CHKIIF(repeatOrderCnt="10", "selected", "") %> >10ȸ</option>
							<option value="20" <%= CHKIIF(repeatOrderCnt="20", "selected", "") %> >20ȸ</option>
						</select>
						�ݺ����� �ֹ���
        				&nbsp;&nbsp;|&nbsp;&nbsp;
						<select class="select" name="fewitem">
							<option></option>
							<option value="2UP" <% if fewitem="2UP" then response.write "selected" %>>2 ���� �̻�</option>
							<option value="4UP" <% if fewitem="4UP" then response.write "selected" %>>4 ���� �̻�</option>
							<option value="5UP" <% if fewitem="5UP" then response.write "selected" %>>5 ���� �̻�</option>
							<option value="6UP" <% if fewitem="6UP" then response.write "selected" %>>6 ���� �̻�</option>
							<option value="10UP" <% if fewitem="10UP" then response.write "selected" %>>10 ���� �̻�</option>
							<option value="3DN" <% if fewitem="3DN" then response.write "selected" %>>3 ���� ����</option>
							<option value="2DN" <% if fewitem="2DN" then response.write "selected" %>>2 ���� ����</option>
							<option value="1DN" <% if fewitem="1DN" then response.write "selected" %>>1 ���� ����</option>
						</select>
						�ֹ���
        			</td>
        		</tr>

        	</table>
			<!--
			<input type="checkbox" name="ems" <% if ems="on" then response.write "checked" %> > <b>�ؿܹ��</b>

			<input type="checkbox" name="epostmilitary" <% if epostmilitary="on" then response.write "checked" %> > <b>���δ�</b>
			-->


			<!--
			<input type="checkbox" name="imsi" <% if imsi="on" then response.write "checked" %> > <b>�ӽ�(���ѵ��� ���� ����)</b>
			<font color="#AAAAAA">
			<input type="checkbox" name="sagawa" <% if sagawa="on" then response.write "checked" %> onClick="alert('�Ϲ�������ø� ���� (��ǰ���,��Ź��ǰ �˻� ����ȵ�)');"> �ӽ�(�簡�ͱǿ�)
			</font>
			-->

        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        �� ��������� �Ǽ� : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
			�� �ݾ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			��հ��ܰ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">�� �˻��ֹ��Ǽ� : </div>
	        <div id="currtensearchno">�ٹ����ٹ�� �ֹ��Ǽ� : </div>
	        <!-- input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> ����������� -->
        </td>
        <td align="right">
		<form name="groupform">
		    <!--
		    <select name="baljutype">
		        <option value="">�Ϲ�
		        <option value="D">DAS
		        <option value="S">��ǰ���
		    </select>
		    -->
			<b>
			<%
			Select Case extSiteName
				Case "10x10"
					response.write "����(���޸� ����) �ֹ�"
				Case "extSiteAll"
					response.write "���޸���ü(��������) �ֹ�"
				Case "cjmall"
					response.write "���޸�(cjmall) �ֹ�"
				Case "interpark"
					response.write "���޸�(interpark) �ֹ�"
				Case "lotteCom"
					response.write "���޸�(lotteCom) �ֹ�"
				Case "lotteimall"
					response.write "���޸�(lotteimall) �ֹ�"
				Case "etcExtSite"
					response.write "��Ÿ���޸� �ֹ�"
				Case Else
					response.write "��ü �ֹ�"
			End Select
			%>
			</b>
			&nbsp;
		    <select name="songjangdiv">
		        <option value="">�ù�缱��</option>
				<!-- <option value="2" >�����ù�</select> -->
                <% if (now()>"2010-04-01") then %>
                	<option value="4" >CJ�ù�</option>
					<option value="98" >�����</option>
                <% else %>
                	<option value="4" >CJ�ù�</option>
			   		<option value="24" >�簡��</option>
			   	<% end if %>
			   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS(�ؿ�)</option>
				<option value="93" <%= CHKIIF(kpack="on","selected","") %> >EMS(KPACK)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
			   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >��ü��(���δ�)</option>
				<option value="91" <%= CHKIIF(cn10x10="on","selected","") %> >EMS(�߱���)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
				<% if False then %>
				<option value="92" <%= CHKIIF(ecargo="on","selected","") %> >�ؿ�(ECARGO)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
				<% end if %>
		    </select>
			<select name="workgroup">
			   	<option value="">�۾��׷�</option>
			   	<option value="A" >A</option>
			   	<option value="B" >B</option>
			   	<option value="C" >C</option>
			   	<option value="D" >D</option>
			   	<option value="F" >F</option>
				<option value="K" >K</option>
				<option value="L" >L</option>
				<option value="N" >N(��ǰ)</option>
				<option value="M" >M(3PL��ǰ)</option>
			   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)</option>
			   	<option value="R" <%= CHKIIF(kpack="on","selected","") %> >R(KPACK)</option>
			   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(���δ�)</option>
				<option value="H" <%= CHKIIF(cn10x10="on","selected","") %> >H(�߱���)</option>
				<% if False then %>
				<option value="J" <%= CHKIIF(ecargo="on","selected","") %> >J(��ī��)</option>
				<% end if %>
			   	<option value="Z" >Z(����)</option>
		   	</select>
			<input type="button" value="�����ֹ� ������ü��ۼ�" onclick="CheckNBalju()">
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(true)"></td>
	<td width="80">�ֹ���ȣ</td>
	<td width="120">Site</td>
	<td width="50">����</td>
	<td width="120">UserID</td>
	<% if (FALSE) then %>
	<td width="120">������</td>
	<% end if %>
	<td width="120">������</td>
	<td width="60">�����ݾ�</td>
	<td width="60">�����Ѿ�</td>
	<td width="80">�������</td>
	<td width="80">�ŷ�����</td>
	<td width="110">�ֹ���</td>
	<td width="60">��ǰ<br />������</td>
	<td>
	    <% if upbeaInclude<>"" then %>
	    ��������
	    <% else %>
	    �ٹ�����
	    <% end if %>
	    </td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if tenbaljucount < CLng(pagesize) and (tenbaljucount < MaxTenBaljuCount or show100 <> "Y")  then %>
	<% if (ix/FlushCount)=CLNG(ix/FlushCount) then response.write CLNG(ix/FlushCount): response.flush %>
<form name="frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>" method="post" style="margin:0px;">
<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(ix).FOrderSerial %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sitename" value="<%= ojumun.FItemList(ix).FSiteName %>">
<input type="hidden" name="dlvcontrycode" value="<%= ojumun.FItemList(ix).FDlvcountryCode %>">
<tr align="center" bgcolor="#FFFFFF">

<!-- !!! EMS ���δ� �߱������ üũ�� Ŭ�������Ͽ��� �Ѵ�. !!! -->

<% if ((ems<>"") or (epostmilitary<>"") or (cn10x10<>"") or (ecargo<>"") or (kpack<>"") or (ojumun.FItemList(ix).FDlvcountryCode=deliveryarea)) then %>
<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
<% else %>
<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(ix).FDlvcountryCode<>"" and ojumun.FItemList(ix).FDlvcountryCode<>"KR","disabled","") %> ></td>
<% end if %>

<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
<td><%= ojumun.FItemList(ix).FDlvcountryCode %></td>
<td><%= printUserId(ojumun.FItemList(ix).FUserID,2,"*") %></td>
<% if (FALSE) then %>
<td><%= ojumun.FItemList(ix).FBuyName %></td>
<% end if %>
<td><%= ojumun.FItemList(ix).FReqName %></td>
<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
<td><%= Left(ojumun.FItemList(ix).FRegDate,16) %></td>
<td><%= ojumun.FItemList(ix).FTenbeaItemKindCnt %></td>
<td>
<% if ojumun.FItemList(ix).Ftenbeaexists then %>
<input type="hidden" name="tenbeaexists" value="Y">
<% tenbaljucount = tenbaljucount + 1 %>
�� <%= tenbaljucount %>
<% else %>
<input type="hidden" name="tenbeaexists" value="N">
<% end if %>
</td>
</tr>
</form>
	<% else %>
		<% exit for %>
	<% end if %>
	<% next %>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="dobaljumaker.asp" style="margin:0px;">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="sitename" value="">
<input type="hidden" name="songjangdiv" value="">
<input type="hidden" name="workgroup" value="">
<input type="hidden" name="baljutype" value="">
<input type="hidden" name="ems" value="">
<input type="hidden" name="epostmilitary" value="">
<input type="hidden" name="cn10x10" value="">
<input type="hidden" name="ecargo" value="">
<input type="hidden" name="kpack" value="">
<input type="hidden" name="extSiteName" value="">
<input type="hidden" name="isGiftPojang" value="">
</form>
<%
set ojumun = Nothing
%>

<script language='javascript'>
document.all.currsearchno.innerHTML = "�˻����� : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "�ٹ����ٹ�� �˻����� : <Font color='#3333FF'><%= tenbaljucount %></font>";
tenBaljuCnt = 1*<%= tenbaljucount %>;
<% if onlyOne<>"" then %>
EnableDiable(frm.onlyOne);
<% end if %>

<% if danpumcheck<>"" then %>
EnableDiable(frm.danpumcheck);
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
