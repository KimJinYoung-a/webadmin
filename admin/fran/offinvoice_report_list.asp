<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  수출신고필증관리
' History : 2015.05.27 최초생성자 모름
'			2016.03.18 한용민 수정
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
	response.write "<script>alert('인덱스는 숫자만 입력가능합니다.');</script>"
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

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF">
	<td width="50" height="25" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색<br>조건</td>
	<td align="left">
		날짜기준 :
		<select class="select" name="dateFlag">
			<option value="">-선택-</option>
			<option value="regdate" <%if (dateFlag = "regdate") then %>selected<% end if %> >등록일</option>
			<option value="reportdate" <%if (dateFlag = "reportdate") then %>selected<% end if %> >신고일자</option>
		</select>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		신고번호 : <input type="text" class="text" name="reportno" value="<%= reportno %>" size=20>
		&nbsp;
		IDX : <input type="text" class="text" name="masteridx" value="<%= masteridx %>" size=20>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF" >
	<td align="left">
		ShopID : 
		<% 'drawSelectBoxOffShop "shopid",shopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		&nbsp;
		<input type="checkbox" name="excnoreport" value="Y" <% if (excnoreport = "Y") then %>checked<% end if %> > 서류 미등록 인보이스 제외
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		검색결과 : <b><%= ocoffinvoice.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2" width="40" height="50">IDX</td>
	<td rowspan="2" width="70">등록일</td>
	<td rowspan="2" width="120">신고번호</td>
	<td rowspan="2" width="70">신고일자</td>
	<td colspan="2" height="25">구매자정보</td>
	<td colspan="7">결제금액</td>
	<td rowspan="2"  width="200">서류업로드</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">출고처ID</td>
	<td>출고처명</td>
	<td width="50">통화코드</td>
	<td width="60">환율</td>
	<td width="80">신고금액</td>
	<td width="80">원화</td>
	<td width="80">상품(원)</td>
	<td width="80">운임(원)</td>
	<td width="80">오차(원)</td>
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
			<input type="button" class="button" value="필증1" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,1)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename2 <> "") then %>
			<input type="button" class="button" value="필증2" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,2)">
			<% end if %>
			<% if (ocoffinvoice.FItemList(i).Fexportdeclarefilename3 <> "") then %>
			<input type="button" class="button" value="필증3" onClick="PopDownloadExportDeclareFile(<%= ocoffinvoice.FItemList(i).Fidx %>,3)">
			<% end if %>
		</td>
	</tr>
	<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="8">총계</td>
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
		<td colspan=21 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<%
set ocoffinvoice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
