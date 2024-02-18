<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����Ű���������
' History : 2015.05.27 ���ʻ����� ��
'			2016.03.18 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->

<%
dim page, shopid, research, i, reportdate, reportno, masteridx, excnoreport,  yyyy1,mm1 ,dd1,yyyy2,mm2,dd2, fromDate ,toDate, dateFlag
	menupos = request("menupos")
	page = request("page")
	shopid = request("shopid")
	research = request("research")
	reportdate = request("reportdate")
	reportno = request("reportno")
	masteridx = request("masteridx")
	excnoreport = request("excnoreport")
	yyyy1 		= request("yyyy1")
	mm1 		= request("mm1")
	dd1 		= request("dd1")
	yyyy2 		= request("yyyy2")
	mm2 		= request("mm2")
	dd2 		= request("dd2")
	dateFlag 	= request("dateFlag")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))-1
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = Left(DateSerial(yyyy1, mm1, dd1), 10)
toDate = Left(DateSerial(yyyy2, mm2, dd2+1), 10)

if (masteridx <> "") and Not IsNumeric(masteridx) then
	masteridx = ""
	response.write "<script>alert('�ε����� ���ڸ� �Է°����մϴ�.');</script>"
end if

if (page = "") then
	page = 1
end if

if (research = "") then
	excnoreport = "Y"
end if

dim ocoffinvoice

set ocoffinvoice = new COffInvoice
	ocoffinvoice.FRectShopid = shopid
	ocoffinvoice.FCurrPage = page
	ocoffinvoice.Fpagesize = 50
	ocoffinvoice.FRectReportDate = reportdate
	ocoffinvoice.FRectReportNo = reportno
	ocoffinvoice.FRectMasterIDX = masteridx
	ocoffinvoice.FRectExcNoReport = excnoreport
	ocoffinvoice.FRectDateFlag = dateFlag
	ocoffinvoice.FRectFromDate = fromDate
	ocoffinvoice.FRectToDate = toDate
	ocoffinvoice.GetMasterList

%>

<script language='javascript'>

function PopDownloadExportDeclareFile(masteridx,ino) {
	var popwin;

	popwin = window.open('<%= uploadImgUrl %>/linkweb/offinvoice/offinvoice_download.asp?idx=' + masteridx+'&ino='+ino,'PopDownloadExportDeclareFile','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function GotoPage(frm, pageno) {
	frm.page.value = pageno;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF">
	<td width="50" height="25" bgcolor="<%= adminColor("gray") %>" rowspan="2">�˻�<br>����</td>
	<td align="left">
		��¥���� :
		<select class="select" name="dateFlag">
			<option value="">-����-</option>
			<option value="regdate" <%if (dateFlag = "regdate") then %>selected<% end if %> >�����</option>
			<option value="reportdate" <%if (dateFlag = "reportdate") then %>selected<% end if %> >�Ű�����</option>
		</select>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		�Ű��ȣ : <input type="text" class="text" name="reportno" value="<%= reportno %>" size=20>
		&nbsp;
		IDX : <input type="text" class="text" name="masteridx" value="<%= masteridx %>" size=20>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF" >
	<td align="left">
		ShopID : 
		<% 'drawSelectBoxOffShop "shopid",shopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		<input type="checkbox" name="excnoreport" value="Y" <% if (excnoreport = "Y") then %>checked<% end if %> > ���� �̵�� �κ��̽� ����
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		�˻���� : <b><%= ocoffinvoice.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2" width="40" height="50">IDX</td>
	<td rowspan="2" width="70">�����</td>
	<td rowspan="2" width="120">�Ű��ȣ</td>
	<td rowspan="2" width="70">�Ű�����</td>
	<td colspan="2" height="25">����������</td>
	<td colspan="7">�����ݾ�</td>
	<td rowspan="2"  width="200">�������ε�</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">���óID</td>
	<td>���ó��</td>
	<td width="50">��ȭ�ڵ�</td>
	<td width="60">ȯ��</td>
	<td width="80">�Ű�ݾ�</td>
	<td width="80">��ȭ</td>
	<td width="80">��ǰ(��)</td>
	<td width="80">����(��)</td>
	<td width="80">����(��)</td>
</tr>
<% if ocoffinvoice.FResultCount >0 then %>
	<%
	dim tot_reportforeigntotalprice, tot_reporttotalprice, tot_totalgoodsprice, tot_totalboxprice, tot_errorno

	for i=0 to ocoffinvoice.FResultcount-1

	tot_reportforeigntotalprice = tot_reportforeigntotalprice + ocoffinvoice.FItemList(i).freportforeigntotalprice
	tot_reporttotalprice = tot_reporttotalprice + ocoffinvoice.FItemList(i).freporttotalprice
	tot_totalgoodsprice = tot_totalgoodsprice + ocoffinvoice.FItemList(i).ftotalgoodsprice
	tot_totalboxprice = tot_totalboxprice + ocoffinvoice.FItemList(i).ftotalboxprice
	tot_errorno = tot_errorno + (ocoffinvoice.FItemList(i).Freporttotalprice - (ocoffinvoice.FItemList(i).Ftotalgoodsprice + ocoffinvoice.FItemList(i).Ftotalboxprice))
	%>
	<tr bgcolor="#FFFFFF">
		<td align="center" height="25"><a href="offinvoice_modify.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>"  target="_blank"><%= ocoffinvoice.FItemList(i).Fidx %></a></td>
		<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Freportno %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Freportdate %></td>
		<td align="center">
			<a href="offinvoice_modify.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>" target="_blank">
			<%= ocoffinvoice.FItemList(i).Fshopid %></a>
		</td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Fshopname %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Freportpriceunit %></td>
		<td align="right"><%= FormatNumber(ocoffinvoice.FItemList(i).Freportexchangerate, 2) %></td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Freportforeigntotalprice, 2) %>
		</td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Freporttotalprice, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 0) %>
		</td>
		<td align="right">
			<%= FormatNumber((ocoffinvoice.FItemList(i).Freporttotalprice - (ocoffinvoice.FItemList(i).Ftotalgoodsprice + ocoffinvoice.FItemList(i).Ftotalboxprice)), 0) %>
		</td>

		<td align="center">
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename <> "") then %>
			<input type="button" class="button" value="����1" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,1)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename2 <> "") then %>
			<input type="button" class="button" value="����2" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,2)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename3 <> "") then %>
			<input type="button" class="button" value="����3" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,3)">
			<% end if %>
		</td>
	</tr>
	<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="8">�Ѱ�</td>
		<td align="right"><%= CurrFormat(tot_reportforeigntotalprice) %></td>
		<td align="right"><%= CurrFormat(tot_reporttotalprice) %></td>
		<td align="right"><%= CurrFormat(tot_totalgoodsprice) %></td>
		<td align="right"><%= CurrFormat(tot_totalboxprice) %></td>
		<td align="right"><%= CurrFormat(tot_errorno) %></td>
		<td></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
			<% if ocoffinvoice.HasPreScroll then %>
				<a href="javascript:GotoPage(frm, <%= ocoffinvoice.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
				<% if i>ocoffinvoice.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:GotoPage(frm, <%= i %>)">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ocoffinvoice.HasNextScroll then %>
				<a href="javascript:GotoPage(frm, <%= i %>)">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=21 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<%
set ocoffinvoice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
