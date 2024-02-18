<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����α� ī�װ�-ä��
' Hieditor : 2019.02.14 ������ ����
'			 2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim research
Dim i, yyyy1,mm1,yyyy2,mm2, dd1, dd2, fromDate ,toDate ,oCMaechulLog, page, vatinclude, targetGbn, mwdiv_beasongdiv
dim searchfield, searchtext, dategbn, actDivCode, makerid, excptdlv, exceptSite
dim excTPL
dim searchGbn
dim yyyy3, mm3, yyyy4, mm4, dd3, dd4, fromDate2, toDate2
dim useNewDB, vPurchasetype, showChannel, grpBy

	research = requestCheckvar(request("research"),10)

	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	dd2     = requestcheckvar(request("dd2"),10)
	yyyy4   = requestcheckvar(request("yyyy4"),10)
	mm4     = requestcheckvar(request("mm4"),10)
	dd4     = requestcheckvar(request("dd4"),10)
	vatinclude     = requestcheckvar(request("vatinclude"),1)
	targetGbn     = requestcheckvar(request("targetGbn"),16)
	mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),10)
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	dategbn     = requestCheckvar(request("dategbn"),32)
	actDivCode = requestCheckvar(request("actDivCode"),10)
	makerid   = requestcheckvar(request("makerid"),32)
    excptdlv  = requestcheckvar(request("excptdlv"),10)
    exceptSite = requestcheckvar(request("exceptSite"),10)
	searchGbn = requestcheckvar(request("searchGbn"),10)
	vPurchasetype = requestcheckvar(request("purchasetype"),10)

	excTPL 	= request("excTPL")
	useNewDB 	= request("useNewDB")
	showChannel = request("showChannel")

if dategbn="" then dategbn="ActDate"

if showChannel="Y" then
	grpBy = "cateChn"
else
	grpBy = "cate"
end if

if (research = "") then
	excTPL = "Y"
	useNewDB = "Y"
end if

if (yyyy2="") then yyyy2 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm2="") then mm2 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd2="") then dd2 = "01"
if (yyyy4="") then yyyy4 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm4="") then mm4 = Cstr(Month( dateadd("m",-1,date()) ))
''if (dd4="") then dd4 = Cstr(Day( dateadd("d",-1,DateSerial(Year(Date()), Month(Date()), 1)) ))
if (dd4="") then dd4 = "01"
if (targetGbn = "") then targetGbn = "ON"

yyyy1=yyyy2
mm1=mm2
dd1=dd2
yyyy3=yyyy4
mm3=mm4
dd3=dd4


fromDate = DateSerial(yyyy2, mm2, dd2)
toDate = DateSerial(yyyy4, mm4, dd4+1)
''fromDate2 = DateSerial(yyyy3, mm3,"01")
''toDate2 = DateSerial(yyyy4, mm4+1,"01")

''rw fromDate &"~"&toDate
set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 1000
	oCMaechulLog.FCurrPage = 1
	oCMaechulLog.FRectStartDate = fromDate
	oCMaechulLog.FRectEndDate = toDate

    ''������
	''oCMaechulLog.FRectStartDate2 = fromDate2
	''oCMaechulLog.FRectEndDate2 = toDate2

	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRecttargetGbn = targetGbn
	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
	oCMaechulLog.FRectDategbn = dategbn
	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectmakerid = makerid
	oCMaechulLog.FRectExceptDlv = excptdlv
	oCMaechulLog.FRectExceptSite = exceptSite

	oCMaechulLog.FRectExcTPL = excTPL
	oCMaechulLog.FRectGrpBy = grpBy
	oCMaechulLog.FRectUseNewDB = useNewDB
	oCMaechulLog.FRectPurchaseType = vPurchasetype

    oCMaechulLog.GetMaechul_month_item_Log
%>
<script type="text/javascript">

function searchSubmit(){
	frm.target = "";
	frm.action = "";
	frm.submit();
}

function pop_detail_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, vatinclude, mwdiv_beasongdiv){
	<% if dategbn="ActDate" then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?actDate_yyyy1='+yyyy1+'&actDate_mm1='+mm1+'&actDate_dd1='+dd1+'&actDate_yyyy2='+yyyy2+'&actDate_mm2='+mm2+'&actDate_dd2='+dd2+'&chkActDate=Y&vatinclude='+vatinclude+'&mwdiv_beasongdiv='+mwdiv_beasongdiv+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&makerid=<%=makerid%>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% elseif (dategbn="chulgoDate") then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?chulgoDate_yyyy1='+yyyy1+'&chulgoDate_mm1='+mm1+'&chulgoDate_dd1='+dd1+'&chulgoDate_yyyy2='+yyyy2+'&chulgoDate_mm2='+mm2+'&chulgoDate_dd2='+dd2+'&chkChulgoDate=Y&vatinclude='+vatinclude+'&mwdiv_beasongdiv='+mwdiv_beasongdiv+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&makerid=<%=makerid%>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% elseif (dategbn="jFixedDt") then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?jFixedDt_yyyy1='+yyyy1+'&jFixedDt_mm1='+mm1+'&jFixedDt_dd1='+dd1+'&jFixedDt_yyyy2='+yyyy2+'&jFixedDt_mm2='+mm2+'&jFixedDt_dd2='+dd2+'&chkjFixedDt=Y&vatinclude='+vatinclude+'&mwdiv_beasongdiv='+mwdiv_beasongdiv+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&makerid=<%=makerid%>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% else %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?orgPay_yyyy1='+yyyy1+'&orgPay_mm1='+mm1+'&orgPay_dd1='+dd1+'&orgPay_yyyy2='+yyyy2+'&orgPay_mm2='+mm2+'&orgPay_dd2='+dd2+'&chkOrgPay=Y&vatinclude='+vatinclude+'&mwdiv_beasongdiv='+mwdiv_beasongdiv+'&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&makerid=<%=makerid%>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% end if %>

	pop_detail_list.focus();
}

function pop_detail_list2(yyyy1, mm1, dd1, yyyy2, mm2, dd2, yyyy3, mm3, dd3, yyyy4, mm4, dd4, vatinclude, mwdiv_beasongdiv){
	var param = "";
	param = "?actDate_yyyy1="+yyyy1+"&actDate_mm1="+mm1+"&actDate_dd1="+dd1+"&actDate_yyyy2="+yyyy2+"&actDate_mm2="+mm2+"&actDate_dd2="+dd2+"&chkActDate=Y&vatinclude="+vatinclude+"&mwdiv_beasongdiv="+mwdiv_beasongdiv+"&targetGbn=<%= targetGbn %>&actDivCode=<%= actDivCode %>&makerid=<%=makerid%>&searchfield=<%=searchfield%>&searchtext=<%=searchtext%>&menupos=<%=menupos%>";
	param = param + "&chulgoDate_yyyy1="+yyyy1+"&chulgoDate_mm1="+mm1+"&chulgoDate_dd1="+dd1+"&chulgoDate_yyyy2="+yyyy2+"&chulgoDate_mm2="+mm2+"&chulgoDate_dd2="+dd2+"&chkChulgoDate=Y";

	var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp' + param,'pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');

	pop_detail_list.focus();
}

function jsSetYYYYMM4() {
	var frm = document.frm;
/*
	if (frm.dategbn.value == "actDateAndChulgoDate") {
		frm.yyyy4.disabled = false;
		frm.mm4.disabled = false;
	} else {
		frm.yyyy4.disabled = true;
		frm.mm4.disabled = true;
	}
	*/
}

window.onload=function(){
	jsSetYYYYMM4();
}

<% if C_ADMIN_AUTH then %>

//var tmp_url = window.location.href.split("?");
//alert(tmp_url[0]);
//alert(getParameter("dd2"));


function getParameter(paramName) {
  var searchString = window.location.search.substring(1),
      i, val, params = searchString.split("&");

  for (i=0;i<params.length;i++) {
    val = params[i].split("=");
    if (val[0] == paramName) {
      return val[1];
    }
  }
  return null;
}

<% end if %>

function reSearchExcelDown(){
    frm.target = "exceldown";
	frm.action = "/admin/maechul/maechul_month_category_log_exceldown.asp"
    frm.submit();
	frm.target = "";
	frm.action = "";
    frm.select_type.value='';
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>" rowspan="3">�˻�</td>
	<td align="left">
		* ��¥ :
		<select class="select" name="dategbn" onChange="jsSetYYYYMM4()">
			<option value="ipkumdate" <%=CHKIIF(dategbn="ipkumdate","selected","")%> >����������</option>
			<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >��������(ó������)</option>
			<option value="chulgoDate" <%=CHKIIF(dategbn="chulgoDate","selected","")%> >�������</option>
			<option value="jFixedDt" <%=CHKIIF(dategbn="jFixedDt","selected","")%> >����Ȯ������</option>
			<!--
			<option value="actDateAndChulgoDate" <%=CHKIIF(dategbn="actDateAndChulgoDate","selected","")%> >��������(ó������) + �������</select>
			-->
		</select>
		&nbsp;
		<% DrawOneDateBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"dd2",dd2,"", "", "", "" %>
		~
		<% DrawOneDateBoxdynamic "yyyy4",yyyy4,"mm4",mm4,"dd4",dd4,"", "", "", "" %>
		&nbsp;
		<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> >
		3PL ���� ����
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>" rowspan="3"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ���ⱸ�� :
        <!--
        �¶���
		<input type="hidden" name="targetGbn" value="ON">
        -->
		<% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
		&nbsp;&nbsp;
		* �������� : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
		&nbsp;&nbsp;
		* ���Ա��� : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
		<!--
		&nbsp;&nbsp;
		* �ֹ����� :
		<select class="select" name="actDivCode">
			<option value=""></option>
			<option value="A" <% if (actDivCode = "A") then %>selected<% end if %> >���ֹ�</option>
			<option value="C" <% if (actDivCode = "C") then %>selected<% end if %> >����ֹ�</option>
			<option value="H" <% if (actDivCode = "H") then %>selected<% end if %> >��ǰ����</option>
			<option value="E" <% if (actDivCode = "E") then %>selected<% end if %> >��ȯ�ֹ�</option>
			<option value="M" <% if (actDivCode = "M") then %>selected<% end if %> >��ǰ�ֹ�</option>
			<option value="CC" <% if (actDivCode = "CC") then %>selected<% end if %> >�������ȭ�ֹ�</option>
			<option value="HH" <% if (actDivCode = "HH") then %>selected<% end if %> >��ǰ��������ֹ�</option>
			<option value="EE" <% if (actDivCode = "EE") then %>selected<% end if %> >��ȯ����ֹ�</option>
			<option value="MM" <% if (actDivCode = "MM") then %>selected<% end if %> >��ǰ����ֹ�</option>
		</select>
		-->
		* �������� :
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
		&nbsp;&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
			<option value=""></option>
			<!-- option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option -->
			<option value="sitename" <% if (searchfield = "sitename") then %>selected<% end if %> >����ó</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">

		&nbsp;(<input type="checkbox" name="exceptSite" <%=CHKIIF(exceptSite="on","checked","")%> >�ش����ó����)
		&nbsp;&nbsp;
		* ��ۺ�/����� : <input type="checkbox" name="excptdlv" <%=CHKIIF(excptdlv<>"","checked","")%> >����
		&nbsp;&nbsp;
		* <input type="checkbox" name="useNewDB" <%=CHKIIF(useNewDB<>"","checked","")%> value="Y" /> DW�����
		&nbsp;&nbsp;
		* <input type="checkbox" name="showChannel" <%=CHKIIF(showChannel<>"","checked","")%> value="Y" /> ä�� ǥ��
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� �ӵ��� ������ ��� ������ ���ð� ��ٷ� �ּ���. ���ϰ� ū ������ �Դϴ�.
	</td>
	<td align="right">
		<input type="button" class="button" value="�����ٿ�ε�" onclick="reSearchExcelDown();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!--<h5>�۾���</h5>-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%= oCMaechulLog.FTotalcount %></b> �� �� <%= oCMaechulLog.FPageSize %> �Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td rowspan="2">���ؿ�</td>
	<% else %>
	<td rowspan="2">������<br>(ó����)</td>
	<td rowspan="2">����</td>
	<% end if %>

	<td rowspan="2">����<br>����</td>
	<td rowspan="2">(����)<br />ī�װ�</td>
	<% if showChannel="Y" then %><td rowspan="2">ä��</td><% end if %>
	<td rowspan="2">��������</td>
	<td rowspan="2">����<Br>����</td>
	<td rowspan="2">����<Br>����</td>
	<td rowspan="2">SKU</td>
    <td rowspan="2">�Ǹż���</td>
	<% if (C_InspectorUser = False) then %>
	<td rowspan="2">�Һ��ڰ�<br>�հ�</td>
	<td rowspan="2">�ǸŰ�<br>(���ΰ�)</td>
	<td rowspan="2">��ǰ����<br>���밡</td>
	<td colspan="3">���ʽ�����</td>
	<td rowspan="2">��Ÿ����<br>(�þ�)</td>
	<% end if %>
	<td rowspan="2">�����Ѿ�</td>
	<td rowspan="2"><b>���ް���</b></td>
	<td rowspan="2">����</td>
	<td rowspan="2">��ü<Br>�����</td>
	<td rowspan="2"><b>ȸ�����</b></td>
	<td rowspan="2">����<Br>���ϸ���</td>
	<td rowspan="2">���<br>���԰�</td>
	<td rowspan="2">���<br>����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="45">����<br>����</td>
	<td width="45">����<br>����</td>
	<td width="45">��ۺ�<br>����</td>
	<% end if %>
</tr>
<%
Dim ttl_orgTotalPrice,ttl_subtotalpriceCouponNotApplied, ttl_totalsum
Dim ttl_proCpnDiscount, ttl_totalPriceBonusCouponDiscount, ttl_totalBeasongBonusCouponDiscount, ttl_allatdiscountprice
Dim ttl_totalMaechulPrice,ttl_totalMileage ,ttl_totalBuycash, ttl_totalUpcheJungsanCash
dim ttl_avgipgoPrice, ttl_overValueStockPrice
dim ttl_itemno, ttl_sku
%>
<% if oCMaechulLog.FresultCount >0 then %>
<% for i=0 to oCMaechulLog.FresultCount -1 %>
<%
ttl_orgTotalPrice=ttl_orgTotalPrice+oCMaechulLog.FItemList(i).forgTotalPrice
ttl_subtotalpriceCouponNotApplied=ttl_subtotalpriceCouponNotApplied+oCMaechulLog.FItemList(i).fsubtotalpriceCouponNotApplied
ttl_totalsum=ttl_totalsum+oCMaechulLog.FItemList(i).ftotalsum

ttl_proCpnDiscount=ttl_proCpnDiscount+(oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount)
ttl_totalPriceBonusCouponDiscount=ttl_totalPriceBonusCouponDiscount+oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount
ttl_totalBeasongBonusCouponDiscount=ttl_totalBeasongBonusCouponDiscount+oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount
ttl_allatdiscountprice=ttl_allatdiscountprice+oCMaechulLog.FItemList(i).fallatdiscountprice

ttl_totalMaechulPrice=ttl_totalMaechulPrice+oCMaechulLog.FItemList(i).ftotalMaechulPrice

ttl_totalMileage=ttl_totalMileage+oCMaechulLog.FItemList(i).ftotalMileage
ttl_totalBuycash=ttl_totalBuycash+oCMaechulLog.FItemList(i).ftotalBuycash
ttl_totalUpcheJungsanCash=ttl_totalUpcheJungsanCash+oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash

ttl_avgipgoPrice = ttl_avgipgoPrice + oCMaechulLog.FItemList(i).FavgipgoPrice
ttl_overValueStockPrice = ttl_overValueStockPrice + CLng(oCMaechulLog.FItemList(i).FoverValueStockPrice)

ttl_itemno = ttl_itemno + oCMaechulLog.FItemList(i).Fitemno
ttl_sku = ttl_sku + oCMaechulLog.FItemList(i).Fsku

%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td>
		<a href="javascript:pop_detail_list('<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd2 %>','<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd4 %>','<%= oCMaechulLog.FItemList(i).fvatinclude %>','<%= oCMaechulLog.FItemList(i).fmwdiv_beasongdiv %>');" onfocus="this.blur()">
		<%= oCMaechulLog.FItemList(i).fyyyymm %></a>
	</td>
	<% else %>
	<td>
	    <% if (oCMaechulLog.FItemList(i).fyyyymm2<>"") then %>
		<a href="javascript:pop_detail_list2('<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd2 %>','<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd4 %>','<%= left(oCMaechulLog.fitemlist(i).Fyyyymm2,4) %>','<%= mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2) %>','<%= dd2 %>','<%= left(oCMaechulLog.fitemlist(i).Fyyyymm2,4) %>','<%= mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2) %>','<%= LastDayOfThisMonth( left(oCMaechulLog.fitemlist(i).Fyyyymm2,4),mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2)) %>','<%= oCMaechulLog.FItemList(i).fvatinclude %>','<%= oCMaechulLog.FItemList(i).fmwdiv_beasongdiv %>');" onfocus="this.blur()">
		<%= oCMaechulLog.FItemList(i).fyyyymm %></a>
		<% else %>
		<%= oCMaechulLog.FItemList(i).fyyyymm %>
	    <% end if %>
	</td>
	<td>
	    <% if (oCMaechulLog.FItemList(i).fyyyymm2<>"") then %>
		<a href="javascript:pop_detail_list2('<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd2 %>','<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= dd4 %>','<%= left(oCMaechulLog.fitemlist(i).Fyyyymm2,4) %>','<%= mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2) %>','<%= dd2 %>','<%= left(oCMaechulLog.fitemlist(i).Fyyyymm2,4) %>','<%= mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2) %>','<%= LastDayOfThisMonth( left(oCMaechulLog.fitemlist(i).Fyyyymm2,4),mid(oCMaechulLog.fitemlist(i).Fyyyymm2,6,2)) %>','<%= oCMaechulLog.FItemList(i).fvatinclude %>','<%= oCMaechulLog.FItemList(i).fmwdiv_beasongdiv %>');" onfocus="this.blur()">
		<%= oCMaechulLog.FItemList(i).fyyyymm2 %></a>
		<% else %>
		<%= oCMaechulLog.FItemList(i).fyyyymm2 %>
	    <% end if %>
	</td>
	<% end if %>

	<td><%= oCMaechulLog.FItemList(i).FtargetGbn %></td>
	<td><%= oCMaechulLog.FItemList(i).Fcatename %></td>
	<% if showChannel="Y" then %><td><%= oCMaechulLog.FItemList(i).FchannelName %></td><% end if %>
	<td><%= fnColor(oCMaechulLog.FItemList(i).fvatinclude,"tx") %></td>
	<td><%= getmwdiv_beasongdivname(oCMaechulLog.FItemList(i).fmwdiv_beasongdiv) %></td>
	<td><%=oCMaechulLog.FItemList(i).getMeaChulGubunName%></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Fsku, 0) %></td>
    <td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).Fitemno, 0) %></td>
	<% if (C_InspectorUser = False) then %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).forgTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalsum, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalBuycash, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice-oCMaechulLog.FItemList(i).ftotalBuycash, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMileage, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FavgipgoPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FoverValueStockPrice, 0) %></td>
</tr>
<% next %>
<tr bgcolor="FFFFFF" >

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td align="center">�հ�</td>
	<% else %>
	<td align="center" colspan="2">�հ�</td>
	<% end if %>

    <td></td>
	<% if showChannel="Y" then %><td></td><% end if %>
	<td></td>
    <td></td>
	<td></td>
	<td></td>
	<td align="right"><%=FormatNumber(ttl_sku,0)%></td>
    <td align="right"><%=FormatNumber(ttl_itemno,0)%></td>
	<% if (C_InspectorUser = False) then %>
    <td align="right"><%=FormatNumber(ttl_orgTotalPrice,0)%></td>
    <td align="right"><%=FormatNumber(ttl_subtotalpriceCouponNotApplied,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalsum,0)%></td><!-- ��ǰ�������밡 -->
    <td align="right"><%=FormatNumber(ttl_proCpnDiscount,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalPriceBonusCouponDiscount,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalBeasongBonusCouponDiscount,0)%></td>
    <td align="right"><%=FormatNumber(ttl_allatdiscountprice,0)%></td>
	<% end if %>
    <td align="right"><%=FormatNumber(ttl_totalMaechulPrice,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalBuycash,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalMaechulPrice-ttl_totalBuycash,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalUpcheJungsanCash,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalMaechulPrice-ttl_totalUpcheJungsanCash,0)%></td>
    <td align="right"><%=FormatNumber(ttl_totalMileage,0)%></td>
	<td align="right"><%= FormatNumber(ttl_avgipgoPrice, 0) %></td>
	<td align="right"><%= FormatNumber(ttl_overValueStockPrice, 0) %></td>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="30">�˻��� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>
<%
set oCMaechulLog = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
