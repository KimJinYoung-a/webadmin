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

dim idx

dim page, shopid , yyyy1 , mm1 , dd1 , yyyy2 , mm2 , dd2 , designer, statecd , divcode
dim i, totalsellsum, totalsum, totalsuply, totalerr, totalbuy , fromDate , toDate ,shopdiv
dim bankinoutidx
dim chulgoinfoyn, paperinfoyn, depositinfoyn
dim research

	idx = RequestCheckvar(request("idx"),10)

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
    bankinoutidx = RequestCheckvar(request("bankinoutidx"),32)

    chulgoinfoyn = RequestCheckvar(request("chulgoinfoyn"),32)
    paperinfoyn = RequestCheckvar(request("paperinfoyn"),32)
    depositinfoyn = RequestCheckvar(request("depositinfoyn"),32)


if (yyyy1="") then yyyy1 = Cstr(Year(Dateadd("d",now(),-30)))
if (mm1="") then mm1 = Cstr(Month(Dateadd("d",now(),-30)))
''if (dd1="") then dd1 = Cstr(day(Dateadd("d",now(),-30)))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))

fromDate = yyyy1+"-"+mm1
toDate = yyyy2+"-"+mm2

page = request("page")
if page="" then page=1

if (research = "") then
	chulgoinfoyn = "Y"
	paperinfoyn = "Y"
end if


'// ===========================================================================
dim oetcmeachulone

set oetcmeachulone = new CEtcMeachul
oetcmeachulone.FRectidx = idx
oetcmeachulone.getOneEtcMeachul

'// ===========================================================================
dim oetcmeachul
	set oetcmeachul = new CEtcMeachul
	oetcmeachul.FPageSize=200
	oetcmeachul.FCurrpage = page
	oetcmeachul.FRectshopid = shopid
	oetcmeachul.FRectdivcode = divcode
	oetcmeachul.FRectStateCd = statecd

	oetcmeachul.FRectBeforeIssueOnly = "Y"

	if (bankinoutidx = "") then
		'// �����IDX �˻��� ��¥ ����
		oetcmeachul.FRectStartDate = fromDate
		oetcmeachul.FRectendDate = toDate
	else
		oetcmeachul.FRectBankInOutIdx = bankinoutidx
	end if

	oetcmeachul.FRectShopDiv = shopdiv

	oetcmeachul.getEtcMeachulList()

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

function addSelected() {
	var frm = document.frmdetail;
	var tottaxsum = 0;
	var prev_bizsection_cd = "";
	var prev_selltype = "";
	var arrchk = "";

	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var totalsum = document.getElementById("totalsum_" + i);
		var bizsection_cd = document.getElementById("bizsection_cd_" + i);
		var selltype = document.getElementById("selltype_" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked == true) {
			if (prev_bizsection_cd == "") {
				prev_bizsection_cd = bizsection_cd.value;
			} else {
				if (prev_bizsection_cd != bizsection_cd.value) {
					alert("����μ��� ���������� �ٸ� ������ �߰��� �� �����ϴ�.");
					return;
				}
			}

			if (prev_selltype == "") {
				prev_selltype = selltype.value;
			} else {
				if (prev_selltype != selltype.value) {
					alert("����μ��� ���������� �ٸ� ������ �߰��� �� �����ϴ�.");
					return;
				}
			}

			if (arrchk == "") {
				arrchk = chk.value;
			} else {
				arrchk = arrchk + "," + chk.value;
			}

			tottaxsum = tottaxsum + totalsum.value*1;
			hL(chk);
		} else {
			dL(chk);
		}
	}

	opener.ReactMeachulDetailList(arrchk, tottaxsum);

	opener.focus();
	window.close();
}

function selectChanged() {
	var frm = document.frmdetail;
	var tottaxsum = 0;

	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var totalsum = document.getElementById("totalsum_" + i);

		if (chk == undefined) {
			break;
		}

		if (chk.checked == true) {
			tottaxsum = tottaxsum + totalsum.value*1;
			hL(chk);
		} else {
			dL(chk);
		}
	}

	frm.tottaxsum.value = tottaxsum;
}

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){
	selectChanged();
}

window.onload = getOnload;

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		����ó :
		<% NewdrawSelectBoxShopAll "shopid", shopid %>
		&nbsp;
		���� :
		<% Call DrawShopDivBox(shopdiv) %>
		&nbsp;
		<select class="select" name="divcode">
			<option value="">��ü
			<option value="MC" <% if divcode="MC" then response.write "selected" %> > ��������
			<option value="WS" <% if divcode="WS" then response.write "selected" %> > �Ǹź�����(��üƯ��)
			<option value="AA" <% if divcode="AA" then response.write "selected" %> > �Ǹź�����(���� ������)
			<option value="BB" <% if divcode="BB" then response.write "selected" %> > �Ǹź�����(�� ������)
			<option value="GC" <% if divcode="GC" then response.write "selected" %> > ���ͺ�
			<option value="ET" <% if divcode="ET" then response.write "selected" %> > ��Ÿ����(�뿪��)
		</select>
		&nbsp;
		�ۼ����� :
		<select class="select" name="statecd">
			<option value="">��ü
			<option value="0" <% if statecd="0" then response.write "selected" %> >������
			<option value="1" <% if statecd="1" then response.write "selected" %> >��üȮ����
			<option value="3" <% if statecd="3" then response.write "selected" %> >��üȮ�οϷ�
			<option value="7" <% if statecd="7" then response.write "selected" %> >�Ϸ�
		</select>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�˻��Ⱓ :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>	(������)
		&nbsp;
		�����IDX :
		<input type="text" class="text" name="bankinoutidx" value="<%= bankinoutidx %>">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frmdetail" method="get" action="">
<tr>
	<td width="50%" align="left">
		* ���������� ���ݰ�꼭(������) �� ����� �����û���� ���� ������ ǥ�õ˴ϴ�.
	</td>
	<td width="50%" align="right">
		<b>����ݾ��հ� :</b>
		<input type="text" class="text" name="tottaxsum" style="text-align: right;" value="0">
		&nbsp;
		<input type="button" class="button" value="�߰��ϱ�" onClick="addSelected()">
	</td>
</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20></td>
	<td width=35>IDX</td>
	<td width=50>������</td>
	<td width=30>����<br>����</td>
	<td>�귣��<br>����</td>
	<td width=100>���óID</td>
	<td width=30>����</td>
	<td>����</td>
	<td>����</td>

	<td width=70><b>����ݾ�</b></td>

	<td width=30>����<br>����</td>
	<td>����μ�</td>
	<td>��������</td>
</tr>
<% if oetcmeachul.FResultCount >0 then %>
<% for i=0 to oetcmeachul.FResultCount-1 %>

<tr bgcolor="#FFFFFF" height="25">
	<td align=center>
		<input type="checkbox" id="chk_<%= i %>" value="<%= oetcmeachul.FItemList(i).Fidx %>" <% if (CStr(oetcmeachul.FItemList(i).Fidx) = idx) then %>checked disabled<% end if %> onClick="selectChanged()">
	</td>
	<input type="hidden" id="totalsum_<%= i %>" value="<%= oetcmeachul.FItemList(i).Ftotalsum %>">
	<input type="hidden" id="bizsection_cd_<%= i %>" value="<%= oetcmeachul.FItemList(i).Fbizsection_cd %>">
	<input type="hidden" id="selltype_<%= i %>" value="<%= oetcmeachul.FItemList(i).Fselltype %>">
	<td align=center><%= oetcmeachul.FItemList(i).Fidx %></td>
	<td align=center><%= oetcmeachul.FItemList(i).FYYYYMM %></td>
	<td align=center><%= oetcmeachul.FItemList(i).FDiffKey %></td>
	<td align=center><%= oetcmeachul.FItemList(i).GetBrandDivName %></td>
	<td align=center><a href="javascript:popMasterEdit('<%= oetcmeachul.FItemList(i).Fidx %>');"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
	<td align=center><%= oetcmeachul.FItemList(i).getShopDivName() %></td>
	<td align=center><font color="<%= oetcmeachul.FItemList(i).GetDivCodeColor %>"><%= oetcmeachul.FItemList(i).GetDivCodeName %></font></td>
	<td><%= oetcmeachul.FItemList(i).Ftitle %></td>

	<td align=right><%= (oetcmeachul.FItemList(i).Ftotalsum) %></td>

	<td align=center>
		<% if Not IsNull(oetcmeachul.FItemList(i).Fpapertype) then %>
			<font color="<%= oetcmeachul.FItemList(i).GetPaperTypeColor %>"><%= oetcmeachul.FItemList(i).GetPaperTypeName %></font>
		<% end if %>
	</td>
	<td align=center>
		<%= oetcmeachul.FItemList(i).Fbizsection_nm %>
	</td>
	<td align=center>
		<%= oetcmeachul.FItemList(i).Fselltypenm %>
	</td>
</tr>

<% next %>

<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan="14" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
</table>
<% end if %>
</form>
<%
set oetcmeachul = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
