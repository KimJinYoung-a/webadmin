<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%

'// ============================================================================
dim makerid, yyyy1,mm1
makerid = session("ssBctID")
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

dim startDate, endDate
startDate = yyyy1 & "-" & mm1 & "-01"
endDate = Left(DateAdd("m", 1, DateSerial(yyyy1, mm1, 1)), 10)


'// ============================================================================
dim opartner, i, page, groupid
set opartner = new CPartnerUser
opartner.FCurrpage = 1
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser

groupid = opartner.FOneItem.FGroupid

dim ogroup
''set ogroup = new CPartnerGroup
''ogroup.FRectGroupid = groupid
''ogroup.GetOneGroupInfo


'// ============================================================================
page   = requestCheckvar(request("page"),10)

if (page = "") then
	page = "1"
end if


dim oTax
set oTax = new CTax
oTax.FCurrPage = page
oTax.FPageSize = 20
oTax.FRectSdate = startDate
oTax.FRectEdate = endDate
oTax.FRectSupplyGroupID = groupid			'// 그룹아이디 밑에 모든 브랜드 발행내역 표시
oTax.GetTaxListUpche

dim strIsue

%>
<script language='javascript'>

function goPage(pg)
{
	var frm = document.frm_search;

	frm.page.value= pg;
	frm.submit();
}

function fnGetXLList() {
	var yyyy1,mm1;

	yyyy1 = "<%= yyyy1 %>";
	mm1 = "<%= mm1 %>";

    if (confirm(yyyy1 + '-' + mm1 + ' 고객 세금계산서 발행내역을 저장하시겠습니까?')) {
        // var popwin = window.open("about:blank","fnGetXLList","width=200,scrollbars=yes,resizable");
	    xlfrm.target = "iiframeXL";
 		xlfrm.action = "custTaxList_XL.asp";
		xlfrm.submit();
    }
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		고객 세금계산서 발행년월 :&nbsp;<% DrawYMBox yyyy1,mm1 %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>

<p>

<% if (oTax.FResultCount < 1) then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td align="left"><strong>* 고객 세금계산서 발행내역</strong></td>
</tr>
<tr height="30">
    <td align="center" bgcolor="#FFFFFF"> 검색 결과가 존재 하지 않습니다.</td>
</tr>
</table>
<% else %>
<p>
<div align="right">
	<input type=button class="button" onclick="fnGetXLList()" value="엑셀저장">
</div>
<p>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="15" align="left"><strong>* 고객 세금계산서 발행내역</strong></td>
</tr>
<form name="frm_list" method="Post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oTax.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oTax.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
	<td width="50">IDX</td>
	<td width="50">상태</td>
	<td><b>공급받는자</b></td>
	<td width="95">사업자번호</td>
	<td width="80">주문번호</td>
	<td>상품명</td>
	<td width="75">발행일</td>
	<td width="30">과세<br>구분</td>
	<td width="65">공급가액</td>
	<td width="50">세액</td>
	<td width="75">합계</td>
	<td>비고</td>
</tr>
	<%
		for i=0 to oTax.FResultCount - 1
			'발급여부
			if oTax.FTaxList(i).FisueYn="Y" then
				strIsue = "<font color=darkblue>발급</font>"
			else
				strIsue = "<font color=darkred>미발급</font>"
			end if
	%>
	<tr height="25" align="center" bgcolor="#FFFFFF">
		<td><%= oTax.FTaxList(i).FtaxIdx %></td>
		<td><%= strIsue %></td>
		<td align="left">&nbsp;<%= oTax.FTaxList(i).FBusiName %></td>
		<td><b><%= oTax.FTaxList(i).FBusiNo %></b></td>
		<td>
			<% if (Trim(oTax.FTaxList(i).Forderserial) <> "") then %>
				<%=oTax.FTaxList(i).Forderserial%>
			<% else %>
				<% if (oTax.FTaxList(i).Forderidx <> 0) then %>
					<%=oTax.FTaxList(i).Forderidx %>
				<% else %>
					<%=oTax.FTaxList(i).GetMultiOrderIdxSUM %>
				<% end if %>
			<% end if %>
		</td>
		<td align="left">&nbsp;<%= db2html(oTax.FTaxList(i).Fitemname) %>&nbsp;</td>
		<td>
			<b><%= FormatDate(oTax.FTaxList(i).FisueDate,"0000-00-00") %></b>
		</td>
		<td><%= oTax.FTaxList(i).TaxTypeString %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(i).FtotalPrice - oTax.FTaxList(i).FtotalTax) %></td>
		<td align="right"><%= CurrFormat(oTax.FTaxList(i).FtotalTax) %></td>
		<td align="right"><b><%= CurrFormat(oTax.FTaxList(i).FtotalPrice) %></b></td>
		<td></td>
	</tr>
	<%
		next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<%
			if oTax.HasPreScroll then
				Response.Write "<a href='javascript:goPage(" & oTax.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
			else
				Response.Write "[pre] &nbsp;"
			end if

			for i=0 + oTax.StartScrollPage to oTax.FScrollCount + oTax.StartScrollPage - 1

				if i>oTax.FTotalpage then Exit for

				if CStr(page) = CStr(i) then
					Response.Write " <font color='red'>[" & i & "]</font> "
				else
					Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
				end if

			next

			if oTax.HasNextScroll then
				Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
			else
				Response.Write "&nbsp; [next]"
			end if
			%>
		</td>
	</tr>
	</form>
</table>
<% end if %>
<%
set oTax = Nothing
%>
<iframe name="iiframeXL" width="0" height="0" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>

<form name=xlfrm method=post action="">
<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
<input type="hidden" name="mm1" value="<%= mm1 %>">
</form>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
