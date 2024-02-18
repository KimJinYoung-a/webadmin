<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ������ �������(����)
' History : 2009.04.07 ������ ����
'			2010.05.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<%
dim page, shopid , yyyy1 , mm1 , dd1 , yyyy2 , mm2 , dd2 , designer, statecd , divcode
dim i, totalsellsum, totalsum, totalsuply, totalerr, totalbuy , fromDate , toDate ,shopdiv, totmatchedipkumsum, totcnt
dim tmpToDate, onlyITS, rmvDupp, selltype, sellBizCd, totdtlsuplysumITS, totdtlbuysumITS, totalsum_tax
dim datetype, research, curryyyy1, currmm1, curryyyy2, currmm2, currstartday, currendday
dim inc3pl
	research 	= RequestCheckvar(request("research"),10)
	yyyy1 		= RequestCheckvar(request("yyyy1"),10)
	mm1 		= RequestCheckvar(request("mm1"),10)
	dd1 		= RequestCheckvar(request("dd1"),10)
	yyyy2 		= RequestCheckvar(request("yyyy2"),10)
	mm2 		= RequestCheckvar(request("mm2"),10)
	dd2 		= RequestCheckvar(request("dd2"),10)
	designer 	= RequestCheckvar(request("designer"),32)
	statecd  	= RequestCheckvar(request("statecd"),10)
	shopid 		= RequestCheckvar(request("shopid"),32)
	divcode 	= RequestCheckvar(request("divcode"),10)
    shopdiv 	= RequestCheckvar(request("shopdiv"),10)
	datetype = RequestCheckvar(request("datetype"),32)
    onlyITS  = RequestCheckvar(request("onlyITS"),32)
    rmvDupp  = RequestCheckvar(request("rmvDupp"),32)
    selltype = RequestCheckvar(request("selltype"),32)
    sellBizCd= RequestCheckvar(request("sellBizCd"),32)
    inc3pl   = RequestCheckvar(request("inc3pl"),32)
if (yyyy1="") then yyyy1 = Cstr(Year(Dateadd("d",now(),-30)))
if (mm1="") then mm1 = Cstr(Month(Dateadd("d",now(),-30)))
''if (dd1="") then dd1 = Cstr(day(Dateadd("d",now(),-30)))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))

fromDate = yyyy1+"-"+mm1 + "-01"

tmpToDate = DateSerial(yyyy2, mm2, 1)
tmpToDate = DateAdd("m", 1, tmpToDate)
tmpToDate = DateAdd("d", -1, tmpToDate)
toDate = Left(tmpToDate, 10)

page = request("page")
if page="" then page=1

if (research = "") then
	datetype = "yyyymm"
end if

if (C_InspectorUser) THEN  datetype = "taxdt"

dim oetcmeachul
	set oetcmeachul = new CEtcMeachul
	oetcmeachul.FPageSize=2000
	oetcmeachul.FCurrpage = 1
	oetcmeachul.FRectShopDiv = shopdiv
	oetcmeachul.FRectshopid = shopid
	oetcmeachul.FRectdivcode = divcode
	oetcmeachul.FRectStateCd = statecd
	oetcmeachul.FRectDateType = datetype
	oetcmeachul.FRectStartDate = fromDate
	oetcmeachul.FRectendDate = toDate
    oetcmeachul.FRectOnlyDtlITS = onlyITS
    oetcmeachul.FRectRemoveDupp = rmvDupp
    oetcmeachul.FRectSelltype   = selltype
    oetcmeachul.FRectSellBizCd  = sellBizCd
    oetcmeachul.FRectInc3pl = inc3pl
	oetcmeachul.getEtcMeachulSumList()

dim chulgoinforows		: chulgoinforows = 3
dim paperinforows		: paperinforows = 3
dim depositinforows		: depositinforows = 2
dim otherinforows		: otherinforows = 16

%>

<script language='javascript'>

function popEtcMeachul(){
	var popwin = window.open('popetcmeachulreg.asp?shopid=' + document.frm.shopid.value,'popEtcMeachul','width=1100, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popMasterEdit(iid){
	var popwin = window.open('popetcmeachuledit.asp?idx=' + iid,'popMasterEdit','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popMasterAdd(){
	var popwin = window.open('popetcmeachuledit.asp','popMasterAdd','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popRegMeachulPaper(idx, shopdiv, papertype) {
	var popRegMeachulPaper = window.open('popregpaper.asp?idx=' + idx + '&shopdiv=' + shopdiv + '&papertype=' + papertype,'popRegMeachulPaper','width=400, height=200, scrollbars=yes, resizable=yes');
	popRegMeachulPaper.focus();
}

function DelThis(iid){
	if (!confirm('������ ���� �Ͻðڽ��ϱ�?')){
		return;
	}

	var popwin = window.open('etc_meachul_process.asp?mode=delmaster&idx=' + iid,'delfrm','width=400, height=400, scrollbars=yes, resizable=yes');

}

function popSubmasterEdit(iid){
	var popwin = window.open('popetcmeachul_submaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popIpkumSearch(jungsanidx, serchtype, searchstring, yyyy1, mm1, yyyy2, mm2) {
	var popwin;
	if (serchtype == "txammount") {
		popwin = window.open('pop_ipkum_search.asp?jungsanidx=' + jungsanidx + '&serchtype=' + serchtype + '&txammount=' + searchstring + '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2,'popIpkumSearch','width=900, height=500, scrollbars=yes, resizable=yes');
	} else {
		popwin = window.open('pop_ipkum_search.asp?jungsanidx=' + jungsanidx + '&serchtype=' + serchtype + '&jeokyo=' + searchstring + '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2,'popIpkumSearch','width=900, height=500, scrollbars=yes, resizable=yes');
	}
	popwin.focus();
}

function popIpkumList(jungsanidx) {
	var popwin = window.open('pop_ipkum_list.asp?jungsanidx=' + jungsanidx,'popIpkumList','width=800, height=500, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

function regOffTax(idx){
	var popwin = window.open("pop_offshop_TaxReg.asp?idx=" + idx,"popOffTaxReg","width=640 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function registerOffShopTax(idx){
	var popwin = window.open("/cscenter/taxsheet/tax_register_new.asp?issuetype=etcmeachul&idx=" + idx,"registerOffShopTax","width=850 height=650 scrollbars=yes resizable=yes");
	popwin.focus();
}

function modifyInvoice(shopid, idx, workidx, invoiceidx){
	if (workidx == "") {
		alert("���� �۾��� �����ϼ���");
		return;
	}

	var popwin = window.open("/admin/fran/offinvoice_modify.asp?shopid=" + shopid + "&jungsanidx=" + idx + "&workidx=" + workidx + "&invoiceidx=" + invoiceidx,"modifyInvoice","width=1000 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

// �����ڿ� ���ݰ�꼭
function popTaxPrint(taxNo, bizNo){
	var s_biz_no = "2118700620";	// �ٹ����� ����ڹ�ȣ

	//	���󼭹�	http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp
	//	�׽�Ʈ		http://ifs.neoport.net/jsp/dti/tx/dti_get_pin.jsp
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no="+taxNo+"&cur_biz_no="+s_biz_no+"&s_biz_no="+s_biz_no+"&b_biz_no="+bizNo,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");

	popwinsub.focus();
}

function goView_Bill36524(tax_no, b_biz_no)
{
		window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=670,height=620,status=no, scrollbars=auto, menubar=no");
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=1000,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ����ó :
		<% NewdrawSelectBoxShopAll "shopid", shopid %>
		<p>
		* ���� :
		<% Call DrawShopDivBox(shopdiv) %>
		&nbsp;&nbsp;
		<select class="select" name="divcode">
			<option value="">��ü
			<option value="MC" <% if divcode="MC" then response.write "selected" %> > ��������
			<option value="WS" <% if divcode="WS" then response.write "selected" %> > �Ǹź�����(��üƯ��)
			<option value="AA" <% if divcode="AA" then response.write "selected" %> > �Ǹź�����(���� ������)
			<option value="BB" <% if divcode="BB" then response.write "selected" %> > �Ǹź�����(�� ������)
			<option value="GC" <% if divcode="GC" then response.write "selected" %> > ���ͺ�
			<option value="ET" <% if divcode="ET" then response.write "selected" %> > ��Ÿ����(�뿪��)
		</select>
		&nbsp;&nbsp;
		* �ۼ����� :
		<select class="select" name="statecd">
			<option value="">��ü
			<option value="0" <% if statecd="0" then response.write "selected" %> >������
			<option value="1" <% if statecd="1" then response.write "selected" %> >��üȮ����
			<option value="3" <% if statecd="3" then response.write "selected" %> >��üȮ�οϷ�
			<option value="7" <% if statecd="7" then response.write "selected" %> >�Ϸ�
		</select>
		&nbsp;&nbsp;
		* ����ι� : <%= fndrawSaleBizSecCombo(true,"sellBizCd",sellBizCd,"") %>
		&nbsp;&nbsp;
		* ������� : <% drawPartnerCommCodeBox true,"sellacccd","selltype",selltype,"" %>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �˻��Ⱓ :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		&nbsp;&nbsp;
		<% if (Not C_InspectorUser) THEN %>
		<select class="select" name="datetype">
			<option value="yyyymm" <% if datetype="yyyymm" then response.write "selected" %> > ������
			<option value="taxdt" <% if datetype="taxdt" then response.write "selected" %> > ��꼭�����
		</select>

		<p>
		<input type="checkbox" name="onlyITS" value="on" <%= CHKIIF(onlyITS="on","checked","") %> >ithinkso �� ���� �и�(G02799) + etcithinkso (etcithinkso �ΰ�� ��ü[�󼼳�������])
		&nbsp;&nbsp;
		<input type="checkbox" name="rmvDupp" value="on" <%= CHKIIF(rmvDupp="on","checked","") %> >�Ǹź�����(�� ������), �Ǹź�����(���� ������ streetshop012) ����
		<% end if %>
		&nbsp;&nbsp;* ����ó����
	    <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p />

* ����, ����, �ؿܰ��ͺ�(������� : ��Ÿ) = ���� 0���Դϴ�.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=50>������</td>
	<td>���óID</td>
	<td>���ó��</td>
	<td>���ó<Br>����</td>
	<td>����ι�</td>
	<td>�������</td>
	<td width=30>����</td>
	<td></td>

	<td width=80>�ǸŰ���</td>
	<td width=80><b>�����</b></td>
	<td width=80><b>����ݾ�</b></td>
	<td width=80>���ް�</td>
	<td width=80>����<br>(����,����=0)</td>
	<td width=80><b>�������</b></td>
	<td width=80><b>�Ա�Ȯ�ξ�</b></td>

	<td width="30">�Ǽ�</td>

	<% if (onlyITS="on") then %>
		<td width=80><b>�����(ITS)</b></td>
		<td width=80><b>�������(ITS)</b></td>
	<% end if %>

	<td>��</td>
</tr>
<% if oetcmeachul.FResultCount >0 then %>
<% for i=0 to oetcmeachul.FResultCount-1 %>
<%

totalsellsum = totalsellsum + oetcmeachul.FItemList(i).Ftotalsellcash
totalsum = totalsum + oetcmeachul.FItemList(i).Ftotalsum
totalsuply  = totalsuply + oetcmeachul.FItemList(i).Ftotalsuplycash
totalbuy = totalbuy + oetcmeachul.FItemList(i).Ftotalbuycash
totmatchedipkumsum = totmatchedipkumsum + oetcmeachul.FItemList(i).Ftotmatchedipkumsum
totcnt = totcnt + oetcmeachul.FItemList(i).FCNT

totdtlsuplysumITS = totdtlsuplysumITS + oetcmeachul.FItemList(i).FdtlsuplysumITS
totdtlbuysumITS   = totdtlbuysumITS + oetcmeachul.FItemList(i).FdtlbuysumITS

totalsum_tax = totalsum_tax + oetcmeachul.FItemList(i).gettotalsum_Tax
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td align=center><%= oetcmeachul.FItemList(i).FYYYYMM %></td>
	<td align=center><%= oetcmeachul.FItemList(i).Fshopid %></td>
	<td align=center><%= oetcmeachul.FItemList(i).Fsocname_kor %></td>
	<td align=center><%= getPartnerCommCodeName("pcuserdiv",oetcmeachul.FItemList(i).FpcUserDiv) %></td>
	<td><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
	<td><%= oetcmeachul.FItemList(i).Fselltypenm %></td>
	<td align=center><%= oetcmeachul.FItemList(i).getShopDivName() %></td>
	<td align="left"><%= oetcmeachul.FItemList(i).GetDivCodeName %></td>

	<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsellcash,0) %></td>
	<td align=right><b><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsuplycash,0) %></b></td>
	<td align=right><b><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum,0) %></b></td>
	<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum-oetcmeachul.FItemList(i).gettotalsum_Tax,0) %></td>
	<td align=right><%= formatNumber(oetcmeachul.FItemList(i).gettotalsum_Tax,0) %></td>
	<td align=right><b><%= formatNumber(oetcmeachul.FItemList(i).Ftotalbuycash,0) %></b></td>
	<td align=right>
		<b><%= formatNumber(oetcmeachul.FItemList(i).Ftotmatchedipkumsum,0) %></b>
	</td>
	<td align=right><%= oetcmeachul.FItemList(i).FCNT %></td>

	<% if (onlyITS="on") then %>
		<td align=right><%= formatNumber(oetcmeachul.FItemList(i).FdtlsuplysumITS,0) %></td>
		<td align=right><%= formatNumber(oetcmeachul.FItemList(i).FdtlbuysumITS,0) %></td>
	<% end if %>

	<td align="center"><a href="/admin/offshop/etc_meachul.asp?menupos=1466&yyyy1=<%= Left(oetcmeachul.FItemList(i).FYYYYMM,4) %>&mm1=<%= right(oetcmeachul.FItemList(i).FYYYYMM,2) %>&yyyy2=<%= Left(oetcmeachul.FItemList(i).FYYYYMM,4) %>&mm2=<%= right(oetcmeachul.FItemList(i).FYYYYMM,2) %>&shopid=<%= oetcmeachul.FItemList(i).Fshopid %>&divcode=<%= oetcmeachul.FItemList(i).Fdivcode %>&sellBizCd=<%= oetcmeachul.FItemList(i).Fbizsection_cd %>&selltype=<%= oetcmeachul.FItemList(i).Fselltype%>" target="_etcmeachul">����</a></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=8>�Ѱ�</td>
	<td align=right><%= formatNumber(totalsellsum,0) %></td>
	<td align=right><%= formatNumber(totalsuply,0) %></td>
	<td align=right><%= formatNumber(totalsum,0) %></td>
	<td align=right><%= formatNumber(totalsum-totalsum_tax,0) %></td>
	<td align=right><%= formatNumber(totalsum_tax,0) %></td>
	<td align=right><%= formatNumber(totalbuy,0) %></td>
	<td align=right><%= formatNumber(totmatchedipkumsum,0) %></td>
	<td align=right><%= formatNumber(totcnt,0) %></td>
	<% if (onlyITS="on") then %>
	<td align=right><%= formatNumber(totdtlsuplysumITS,0) %></td>
	<td align=right><%= formatNumber(totdtlbuysumITS,0) %></td>
	<% end if %>
	<td></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan="20" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
</table>
<% end if %>

<%
set oetcmeachul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
