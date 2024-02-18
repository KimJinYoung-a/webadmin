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
dim i, totalsellsum, totalsum, totalsuply, totalerr, totalbuy , fromDate , toDate ,shopdiv
dim bankinoutidx, ipkumstate
dim chulgoinfoyn, paperinfoyn, depositinfoyn, datetype
dim research, selltype, sellBizCd, excTPL, tplgubun
dim searchtype, searchstring
	research 	= RequestCheckvar(request("research"),10)
	ipkumstate 	= RequestCheckvar(request("ipkumstate"),1)
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

	searchtype 		= RequestCheckvar(request("searchtype"),32)
	searchstring 	= RequestCheckvar(request("searchstring"),32)

    chulgoinfoyn = RequestCheckvar(request("chulgoinfoyn"),32)
    paperinfoyn = RequestCheckvar(request("paperinfoyn"),32)
    depositinfoyn = RequestCheckvar(request("depositinfoyn"),32)
	datetype = RequestCheckvar(request("datetype"),32)
    selltype = RequestCheckvar(request("selltype"),32)
    sellBizCd= RequestCheckvar(request("sellBizCd"),32)

	excTPL= RequestCheckvar(request("excTPL"),32)
	tplgubun= RequestCheckvar(request("tplgubun"),32)

if (yyyy1="") then yyyy1 = Cstr(Year(Dateadd("d",now(),-30)))
if (mm1="") then mm1 = Cstr(Month(Dateadd("d",now(),-30)))
''if (dd1="") then dd1 = Cstr(day(Dateadd("d",now(),-30)))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))

if (searchtype = "taxidx") and (searchstring <> "") then
	if Not IsNumeric(searchstring) then
		response.write "<script>alert('���ݰ�꼭 �����ȣ�� ���ڸ� �����մϴ�.');</script>"
		searchstring = ""
	end if
end if


fromDate = yyyy1+"-"+format00(2,mm1)
toDate = yyyy2+"-"+format00(2,mm2)

page = request("page")
if page="" then page=1

if (research = "") then
	chulgoinfoyn = "Y"
	paperinfoyn = "Y"
	depositinfoyn = "Y"
	datetype = "yyyymm"
	if (C_InspectorUser) THEN  excTPL="Y"
	if (C_InspectorUser) THEN  tplgubun="3X"
end if

if (C_InspectorUser) THEN  datetype = "issuedate"

'// ===========================================================================
dim oetcmeachul
	set oetcmeachul = new CEtcMeachul
	oetcmeachul.FPageSize=50
	oetcmeachul.FCurrpage = page
	oetcmeachul.FRectshopid = shopid
	oetcmeachul.FRectdivcode = divcode
	oetcmeachul.FRectStateCd = statecd
	oetcmeachul.FRectDateType = datetype

	if (bankinoutidx = "") then
		'// �����IDX �˻��� ��¥ ����
		oetcmeachul.FRectStartDate = fromDate
		oetcmeachul.FRectendDate = toDate
	else
		oetcmeachul.FRectBankInOutIdx = bankinoutidx
	end if

	oetcmeachul.FRectShopDiv = shopdiv
    oetcmeachul.FRectSelltype   = selltype
    oetcmeachul.FRectSellBizCd  = sellBizCd

	oetcmeachul.FRectSearchType  = searchtype
	oetcmeachul.FRectSearchString  = searchstring

	oetcmeachul.FRectExcTPL  = excTPL
	oetcmeachul.FtplGubun  = tplgubun
	oetcmeachul.frectipkumstate = ipkumstate
	oetcmeachul.getEtcMeachulList()


'// ===========================================================================
dim chulgoinforows		: chulgoinforows = 3
dim paperinforows		: paperinforows = 3
dim depositinforows		: depositinforows = 2
dim otherinforows		: otherinforows = 16

if (chulgoinfoyn <> "Y") then
	chulgoinforows = 0
end if

if (paperinfoyn <> "Y") then
	paperinforows = 0
end if

if (depositinfoyn <> "Y") then
	depositinforows = 0
end if

dim curryyyy1, currmm1, curryyyy2, currmm2
dim currstartday, currendday

Dim IsTaxExist

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
	var popwin = window.open("/cscenter/taxsheet/tax_view.asp?issuetype=etcmeachul&idx=" + idx,"registerOffShopTax","width=1024 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function modifyInvoice(shopid, idx, workidx, invoiceidx){
	if (workidx == "") {
		alert("���� �۾��� �����ϼ���");
		return;
	}

	var popwin = window.open("/admin/fran/offinvoice_modify.asp?shopid=" + shopid + "&jungsanidx=" + idx + "&workidx=" + workidx + "&invoiceidx=" + invoiceidx,"modifyInvoice","width=1024 height=768 scrollbars=yes resizable=yes");
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
	window.open("http://www.bill36524.com/popupBillTax.jsp?NO_TAX=" + tax_no + "&NO_BIZ_NO="+b_biz_no,"view","width=1280,height=960,scrollbars=yes,resizable=yes");
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
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
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
			<option value="WS" <% if divcode="WS" then response.write "selected" %> > �Ǹź�����(��ü��Ź)
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
		<% if (NOT C_InspectorUser) THEN  %>
		&nbsp;
		����ι� : <%= fndrawSaleBizSecCombo(true,"sellBizCd",sellBizCd,"") %>
	    <% end if %>
		&nbsp;
		������� : <% CALL drawPartnerCommCodeBox(true,"sellacccd","selltype",selltype,"") %>

	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�˻����� :
		<select class="select" name="searchtype">
			<option value="">
			<option value="groupcode" <% if searchtype="groupcode" then response.write "selected" %> > �׷��ڵ�
			<option value="taxidx" <% if searchtype="taxidx" then response.write "selected" %> > ���ݰ�꼭�����ȣ
		</select>
		<input type="text" class="text" name="searchstring" value="<%= searchstring %>">
		&nbsp;
		�˻��Ⱓ :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		<% if (NOT C_InspectorUser) then %>
		&nbsp;
		<select class="select" name="datetype">
			<option value="yyyymm" <% if datetype="yyyymm" then response.write "selected" %> >������</option>
			<option value="issuedate" <% if datetype="issuedate" then response.write "selected" %> >���������(�ۼ���)</option>
		</select>
	    <% end if %>
		&nbsp;
		�����IDX :
		<input type="text" class="text" name="bankinoutidx" value="<%= bankinoutidx %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="chulgoinfoyn" value="Y" <%if (chulgoinfoyn = "Y") then %>checked<% end if %> > ������� ǥ��
		<input type="checkbox" name="paperinfoyn" value="Y" <%if (paperinfoyn = "Y") then %>checked<% end if %> > ������������ ǥ��
		<input type="checkbox" name="depositinfoyn" value="Y" <%if (depositinfoyn = "Y") then %>checked<% end if %> > �Ա����� ǥ��
		&nbsp;&nbsp;
		3PL ���� : <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
		<!--
		<input type="checkbox" name="excTPL" value="Y" <%if (excTPL = "Y") then %>checked<% end if %> > 3PL���� ����
		-->
		&nbsp;&nbsp;
		�Աݻ��� : <% drawSelectBoxIpkumState "ipkumstate", ipkumstate, "" %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<% 'if (NOT C_InspectorUser) then %>
<!-- �׼� ���� -->
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" >
<tr>
	<td align="left">
		<input type="button" class="button" value="��Ÿ������" onClick="javascript:popEtcMeachul();">
		<input type="button" class="button" value="��Ÿ�����ۼ�(���ͺ��)" onClick="javascript:popMasterAdd();">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<% 'end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=35>IDX</td>
	<td width=50>������</td>
	<td width=30>����<br>����</td>
	<!--
	<td>�귣��<br>����</td>
	-->
	<td>���óID</td>
	<td width=30>����</td>
	<!--
	<td>����</td>
	-->
	<td>����</td>

	<% if (chulgoinfoyn = "Y") then %>
		<td width=80>�ǸŰ���</td>
		<td width=80><b>�����</b></td>
		<td width=30>����<br>�ݾ�</td>
		<!--
		<td width=70>���԰���</td>
		<td width=40>������</td>
		-->
	<% end if %>

	<td width=80>�ۼ�����</td>
	<td width=1 style="padding: 1px;"></td>

	<td width=80><b>����ݾ�</b></td>
	<td width=80>���ް���</td>
	<td width=80>����</td>

	<% if (paperinfoyn = "Y") then %>
		<td width=65>���������<br>(���ݹ���)</td>
		<td width=30>����<br>����</td>
		<td width=50>����</td>
		<!--
		<td>����μ�</td>
		<td>��������</td>
		-->
	<% end if %>

	<td width=50>�������</td>
	<td width=1 style="padding: 1px;"></td>

	<% if (depositinfoyn = "Y") then %>
		<td width=80><b>�Ա�Ȯ�ξ�</b></td>
		<td width=65>�Ա���</td>
	<% end if %>

	<td width=50>�Աݻ���</td>
	<td width=1 style="padding: 1px;"></td>

	<td width=25>����</td>
</tr>
<% if oetcmeachul.FResultCount >0 then %>
<% for i=0 to oetcmeachul.FResultCount-1 %>
<%

if IsNull(oetcmeachul.FItemList(i).Ftotalsellcash) then
	oetcmeachul.FItemList(i).Ftotalsellcash = 0
end if

if IsNull(oetcmeachul.FItemList(i).Ftotalsuplycash) then
	oetcmeachul.FItemList(i).Ftotalsuplycash = 0
end if

if IsNull(oetcmeachul.FItemList(i).Ftotalsum) then
	oetcmeachul.FItemList(i).Ftotalsum = 0
end if

if IsNull(oetcmeachul.FItemList(i).Ftotalbuycash) then
	oetcmeachul.FItemList(i).Ftotalbuycash = 0
end if

totalsellsum = totalsellsum + oetcmeachul.FItemList(i).Ftotalsellcash
totalsum = totalsum + oetcmeachul.FItemList(i).Ftotalsum
totalsuply  = totalsuply + oetcmeachul.FItemList(i).Ftotalsuplycash
totalerr = totalerr  + oetcmeachul.FItemList(i).Ftotalsum -  oetcmeachul.FItemList(i).Ftotalsuplycash
totalbuy = totalbuy + oetcmeachul.FItemList(i).Ftotalbuycash

if IsNull(oetcmeachul.FItemList(i).Ftotmatchedipkumsum) then
	oetcmeachul.FItemList(i).Ftotmatchedipkumsum = 0
end If

IsTaxExist = False
If Not IsNull(oetcmeachul.FItemList(i).Fpapertype) Then
	If (oetcmeachul.FItemList(i).Fpapertype = "100") Then
		IsTaxExist = True
	End If
End If

%>
<tr bgcolor="#FFFFFF" height="25">
	<td align=center><a href="javascript:popMasterEdit('<%= oetcmeachul.FItemList(i).Fidx %>');"><%= oetcmeachul.FItemList(i).Fidx %></a></td>
	<td align=center><%= oetcmeachul.FItemList(i).FYYYYMM %></td>
	<td align=center><%= oetcmeachul.FItemList(i).FDiffKey %></td>
	<!--
	<td align=center><%= oetcmeachul.FItemList(i).GetBrandDivName %></td>
	-->
	<td align=center><a href="javascript:popMasterEdit('<%= oetcmeachul.FItemList(i).Fidx %>');"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
	<td align=center><%= oetcmeachul.FItemList(i).getShopDivName() %></td>
	<!--
	<td align=center><font color="<%= oetcmeachul.FItemList(i).GetDivCodeColor %>"><%= oetcmeachul.FItemList(i).GetDivCodeName %></font></td>
	-->
	<td><a href="javascript:popSubmasterEdit('<%= oetcmeachul.FItemList(i).Fidx %>');"><%= oetcmeachul.FItemList(i).Ftitle %></a></td>

	<% if (chulgoinfoyn = "Y") then %>
		<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><b><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsuplycash,0) %></b></td>
		<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum-oetcmeachul.FItemList(i).Ftotalsuplycash,0) %></td>
		<!--
		<% if FALSE then %>
		<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=right>
			<% if oetcmeachul.FItemList(i).Ftotalsum<>0 then %>
			<%= CLng(10000-(oetcmeachul.FItemList(i).Ftotalbuycash/oetcmeachul.FItemList(i).Ftotalsum*100*100))/100 %>%
			<% end if %>
		</td>
		<% end if %>
		-->
	<% end if %>

	<td align=center><font color="<%= oetcmeachul.FItemList(i).GetStateColor %>"><%= oetcmeachul.FItemList(i).GetStateName %></font></td>
	<td style="padding: 1px;"></td>

	<td align=right><b><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum,0) %></b></td>
	<% If IsTaxExist Then %>
	<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum-Round((oetcmeachul.FItemList(i).Ftotalsum/11),0),0) %></td>
	<td align=right><%= formatNumber(Round((oetcmeachul.FItemList(i).Ftotalsum/11),0),0) %></td>
	<% Else %>
	<td align=right><%= formatNumber(oetcmeachul.FItemList(i).Ftotalsum,0) %></td>
	<td align=right>0</td>
	<% End If %>

	<% if (paperinfoyn = "Y") then %>
		<td align=center><%= Left(oetcmeachul.FItemList(i).Ftaxdate,10) %></td>
		<td align=center>
			<% if Not IsNull(oetcmeachul.FItemList(i).Fpapertype) then %>
				<font color="<%= oetcmeachul.FItemList(i).GetPaperTypeColor %>"><%= oetcmeachul.FItemList(i).GetPaperTypeName %></font>
			<% end if %>
		</td>
		<td align=center>
			<%

			if oetcmeachul.FItemList(i).Fpapertype = "200" then

				'// ����Ű�����
				if (oetcmeachul.FItemList(i).Finvoiceidx <> "") and (Not IsNull(oetcmeachul.FItemList(i).Finvoiceidx)) then
					%>
					<a href="javascript:modifyInvoice('<%= oetcmeachul.FItemList(i).Fshopid %>', '<%= oetcmeachul.FItemList(i).Fidx %>', '<%= oetcmeachul.FItemList(i).Fworkidx %>', '<%= oetcmeachul.FItemList(i).Finvoiceidx %>');"><%= oetcmeachul.FItemList(i).Finvoiceidx %></a>
					<%
				else
					%>
					<input type="button" class="button" value="�ۼ�" onclick="modifyInvoice('<%= oetcmeachul.FItemList(i).Fshopid %>', '<%= oetcmeachul.FItemList(i).Fidx %>', '<%= oetcmeachul.FItemList(i).Fworkidx %>', '<%= oetcmeachul.FItemList(i).Finvoiceidx %>');">
					<%
				end if

			elseif (oetcmeachul.FItemList(i).Fpapertype = "100") or (oetcmeachul.FItemList(i).Fpapertype = "101") or (oetcmeachul.FItemList(i).Fpapertype = "102") then

				if oetcmeachul.FItemList(i).Fpaperissuetype = "1" then

					'���ݰ�꼭
					if IsNull(oetcmeachul.FItemList(i).FtaxNo) and IsNull(oetcmeachul.FItemList(i).Fissuestatecd) then
						%>
						<input type="button" class="button" value="����" onclick="registerOffShopTax('<%= oetcmeachul.FItemList(i).Fidx %>');" <% if oetcmeachul.FItemList(i).Fstatecd = 0 then %>disabled<% end if %> >
						<%
					else
						if (IsNull(oetcmeachul.FItemList(i).FtaxNo) or oetcmeachul.FItemList(i).FtaxNo = "") and Not IsNull(oetcmeachul.FItemList(i).Fissuestatecd) then
							%>
							<input type="button" class="button" value="����" onclick="registerOffShopTax('<%= oetcmeachul.FItemList(i).Fidx %>');" disabled>
							<%
						elseif Not IsNull(oetcmeachul.FItemList(i).FtaxNo) and Not IsNull(oetcmeachul.FItemList(i).Fissuestatecd) then
							if (Left(oetcmeachul.FItemList(i).FtaxNo,2) = "TX") then
								%>
								<a href="javascript:goView_Bill36524('<%=oetcmeachul.FItemList(i).FtaxNo%>','2118700620');"><%=oetcmeachul.FItemList(i).FtaxLinkidx %></a>
								<%
							else
								%>
								<a href="javascript:popTaxPrint('<%=oetcmeachul.FItemList(i).FtaxNo%>','<%=oetcmeachul.FItemList(i).FbizNo%>');"><img src="/images/icon_print02.gif" border="0"></a>
								<%
							end if
						else
							%>����.<%
						end if

					end if

				end if

			end if
			%>
		</td>
		<!--
		<td align=center>
			<%= oetcmeachul.FItemList(i).Fbizsection_nm %>
		</td>
		<td align=center>
			<%= oetcmeachul.FItemList(i).Fselltypenm %>
		</td>
		-->
	<% end if %>

	<td align=center>
		<%= oetcmeachul.FItemList(i).GetIssueStateName() %>
	</td>
	<td style="padding: 1px;"></td>

	<% if (depositinfoyn = "Y") then %>
		<td align=center>
			<% if (IsNull(oetcmeachul.FItemList(i).Ftotmatchedipkumsum) or (oetcmeachul.FItemList(i).Ftotmatchedipkumsum = 0)) then %>
				<% if (Not IsNull(oetcmeachul.FItemList(i).Fmaymatchedipkumsum)) then %>
					<font color=gray><%= FormatNumber(oetcmeachul.FItemList(i).Ftotalsum,0) %></font>
				<% end if %>
			<% else %>
				<a href="javascript:popIpkumList(<%= oetcmeachul.FItemList(i).Fidx %>)">
					<% if (oetcmeachul.FItemList(i).Ftotalsum = oetcmeachul.FItemList(i).Ftotmatchedipkumsum) then %>
						<b><%= formatNumber(oetcmeachul.FItemList(i).Ftotmatchedipkumsum,0) %></b>
					<% elseif (oetcmeachul.FItemList(i).Ftotalsum < oetcmeachul.FItemList(i).Ftotmatchedipkumsum) then %>
						<b><font color=blue><%= formatNumber(oetcmeachul.FItemList(i).Ftotmatchedipkumsum,0) %></font></b>
					<% else %>
						<b><font color=red><%= formatNumber(oetcmeachul.FItemList(i).Ftotmatchedipkumsum,0) %></font></b>
					<% end if %>
				</a>
			<% end if %>
		</td>
		<td align=center>
			<% if (oetcmeachul.FItemList(i).FStateCd >= "1") then %>
				<%
				if not(isnull(oetcmeachul.FItemList(i).FYYYYMM)) then
				currstartday = DateSerial(Left(oetcmeachul.FItemList(i).FYYYYMM, 4), (Right(oetcmeachul.FItemList(i).FYYYYMM, 2) - 3), 1)
				currendday = DateSerial(Left(oetcmeachul.FItemList(i).FYYYYMM, 4), (Right(oetcmeachul.FItemList(i).FYYYYMM, 2) + 6), 1)
				end if

				curryyyy1 = Year(currstartday)
				currmm1 = Month(currstartday)
				curryyyy2 = Year(currendday)
				currmm2 = Month(currendday)
				%>

				<% if (oetcmeachul.FItemList(i).Fipkumdate = "") or IsNull(oetcmeachul.FItemList(i).Fipkumdate) then %>
					<input type="button" class="button" value="ã��" onClick="popIpkumSearch(<%= oetcmeachul.FItemList(i).Fidx %>, 'txammount', <%= oetcmeachul.FItemList(i).Ftotalsum - oetcmeachul.FItemList(i).Ftotmatchedipkumsum %>, '<%= curryyyy1 %>', '<%= currmm1 %>', '<%= curryyyy2 %>', '<%= currmm2 %>')">
				<% else %>
					<a href="javascript:popIpkumSearch(<%= oetcmeachul.FItemList(i).Fidx %>, 'txammount', <%= oetcmeachul.FItemList(i).Ftotalsum - oetcmeachul.FItemList(i).Ftotmatchedipkumsum %>, '<%= curryyyy1 %>', '<%= currmm1 %>', '<%= curryyyy2 %>', '<%= currmm2 %>')"><%= oetcmeachul.FItemList(i).Fipkumdate %></a>
				<% end if %>
			<% else %>
				<%= oetcmeachul.FItemList(i).Fipkumdate %>
			<% end if %>
		</td>
	<% end if %>

	<td align=center>
		<%= oetcmeachul.FItemList(i).GetIpkumStateName() %>
	</td>
	<td style="padding: 1px;"></td>

	<td align=center>
		<% if oetcmeachul.FItemList(i).FStateCd="0" then %>
		<a href="javascript:DelThis('<%= oetcmeachul.FItemList(i).Fidx %>');">X</a>
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>�Ѱ�</td>
	<td></td>
	<td></td>
	<!--
	<td></td>
	-->
	<td></td>
	<td></td>
	<!--
	<td></td>
	-->
	<td></td>
	<% if (chulgoinfoyn = "Y") then %>
		<td align=right><%= formatNumber(totalsellsum,0) %></td>
		<td align=right><%= formatNumber(totalsuply,0) %></td>
		<td align=right><%= formatNumber(totalerr,0) %></td>
		<!--
		<td align=right><%= formatNumber(totalbuy,0) %></td>
		<td></td>
		-->
	<% end if %>

		<td></td>
		<td style="padding: 1px;"></td>

	<td align=right><%= formatNumber(totalsum,0) %></td>

	<% if (paperinfoyn = "Y") then %>
		<td></td>
		<td></td>
		<td></td>
		<!--
		<td></td>
		<td></td>
		-->
	<% end if %>

	<td></td>
	<td style="padding: 1px;"></td>

	<% if (depositinfoyn = "Y") then %>
		<td></td>
		<td></td>
	<% end if %>

	<td></td>
	<td style="padding: 1px;"></td>

	<td></td>
	<td></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height=20>
	<td colspan="<%= (otherinforows + chulgoinforows + paperinforows + depositinforows) %>" align=center>
	<% if oetcmeachul.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oetcmeachul.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oetcmeachul.StarScrollPage to oetcmeachul.FScrollCount + oetcmeachul.StarScrollPage - 1 %>
		<% if i>oetcmeachul.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oetcmeachul.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" >
	<td colspan="<%= (otherinforows + chulgoinforows + paperinforows + depositinforows) %>" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
</table>
<% end if %>

<%
set oetcmeachul = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
