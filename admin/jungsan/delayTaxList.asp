<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출
' History : 2011.12.27 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/delaytaxcls.asp"-->
<%
dim i, j
dim yyyy1, mm1, yyyy2, mm2
dim yyyymm1, yyymm2, makerid ,offgubun
dim startYYYYMM, endYYYYMM, tmpYYYYMM
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	makerid = request("makerid")
	offgubun = request("offgubun")

if offgubun = "" then offgubun = "ON"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

startYYYYMM = yyyy1 + "-" + mm1
endYYYYMM = yyyy2 + "-" + mm2

dim ocdelaytax
set ocdelaytax = new CDelayTax
	ocdelaytax.FRectStartYYYYMM = startYYYYMM
	ocdelaytax.FRectEndYYYYMM = endYYYYMM

	ocdelaytax.FRectGubun = offgubun

	''ocdelaytax.FRectMakerid = makerid

	ocdelaytax.GetDelayTaxList

dim monthCnt
monthCnt = DateDiff("m", startYYYYMM + "-01", yyyy2 + "-" + mm2 + "-01") + 1

%>

<script type="text/javascript">

function formSubmit(page) {
	frm.page.value=page;
	frm.submit();
}

function popDelayTaxDetail(yyyy1, mm1, yyyy3, mm3, offgubun, issuegubun) {
	var popwin = window.open("popDelayTaxDetailList.asp?yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&yyyy3=" + yyyy3 + "&mm3=" + mm3 + "&offgubun=" + offgubun + "&issuegubun=" + issuegubun,"popDelayTaxDetail","width=1280, height=960, scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		구분 :
		<select class="select" name="offgubun">
		<option value="ON" <% if (offgubun = "ON") then %>selected<% end if %> >온라인</option>
		<option value="OFF" <% if (offgubun = "OFF") then %>selected<% end if %> >오프라인</option>
		<option value="ETC" <% if (offgubun = "ETC") then %>selected<% end if %> >기타매출</option>
		</select>
		&nbsp;
		정산월 : <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="formSubmit('1');">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="40">
		검색결과 : <b><%= ocdelaytax.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2" width="80">정산월</td>
	<td colspan="2">전체내역</td>

	<%
	tmpYYYYMM = startYYYYMM
	for j = 0 to monthCnt - 1
		%>
		<td colspan="2"><%= tmpYYYYMM %></td>
		<%
		tmpYYYYMM = Left(CStr(dateserial(Left(tmpYYYYMM,4),Right(tmpYYYYMM,2)+1,1)), 7)
	next
	%>

	<td colspan="2">발행이전</td>
	<td colspan="2">기타(선발행)</td>
	<td rowspan="2"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">정산액</td>
	<td width="40">건수</td>

	<%
	tmpYYYYMM = startYYYYMM
	for j = 0 to monthCnt - 1
		%>
		<td width="90">정산액</td>
		<td width="40">건수</td>
		<%
		tmpYYYYMM = Left(CStr(dateserial(Left(tmpYYYYMM,4),Right(tmpYYYYMM,2)+1,1)), 7)
	next
	%>

	<td width="80">금액</td>
	<td width="50">건수</td>
	<td width="80">기타액</td>
	<td width="50">건수</td>
</tr>
<%
if ocdelaytax.FresultCount > 0 then
%>
	<%
	for i=0 to ocdelaytax.FresultCount-1
	%>
		<tr bgcolor="#FFFFFF" align="center">
			<td><%= ocdelaytax.FItemList(i).Fyyyymm %></td>
			<td align="right"><%= FormatNumber(ocdelaytax.FItemList(i).Fttl,0)  %></td>
			<td align="right"><%= FormatNumber(ocdelaytax.FItemList(i).FttlCnt,0)  %></td>

			<%
			tmpYYYYMM = startYYYYMM
			for j = 0 to monthCnt - 1
				%>
				<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '<%= Left(tmpYYYYMM, 4) %>', '<%= Right(tmpYYYYMM, 2) %>', '<%= offgubun %>', '1')"><%= FormatNumber(ocdelaytax.FItemList(i).FarrTrPrice(j),0)  %></a></td>
				<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '<%= Left(tmpYYYYMM, 4) %>', '<%= Right(tmpYYYYMM, 2) %>', '<%= offgubun %>', '1')"><%= FormatNumber(ocdelaytax.FItemList(i).FarrTrCnt(j),0)  %></a></td>
				<%
				tmpYYYYMM = Left(CStr(dateserial(Left(tmpYYYYMM,4),Right(tmpYYYYMM,2)+1,1)), 7)
			next
			%>

			<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '', '', '<%= offgubun %>', '2')"><%= FormatNumber(ocdelaytax.FItemList(i).FtrNullPrice,0)  %></a></td>
			<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '', '', '<%= offgubun %>', '2')"><%= FormatNumber(ocdelaytax.FItemList(i).FtrNullCnt,0)  %></a></td>
			<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '', '', '<%= offgubun %>', '9')"><%= FormatNumber(ocdelaytax.FItemList(i).FtrErrPrice,0)  %></a></td>
			<td align="right"><a href="javascript:popDelayTaxDetail('<%= Left(ocdelaytax.FItemList(i).Fyyyymm, 4) %>', '<%= Right(ocdelaytax.FItemList(i).Fyyyymm, 2) %>', '', '', '<%= offgubun %>', '9')"><%= FormatNumber(ocdelaytax.FItemList(i).FtrErrCnt,0)  %></a></td>
			<td></td>
		</tr>
	<% next %>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="40">검색결과가 없습니다.</td>
</tr>
<% end if %>
</table>

<%
set ocdelaytax = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
