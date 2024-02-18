<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandStaticCls.asp"-->
<%

dim yyyy1, mm1, dd1
Dim makerID, page, ordby
dim i

yyyy1	= req("yyyy1", Left(DateAdd("d", -0, Now()),4))
mm1		= req("mm1", Mid(DateAdd("d", -0, Now()),6,2))
dd1		= req("dd1", Mid(DateAdd("d", -0, Now()),9,2))
makerID = req("makerID", "")
page = req("page", "1")
ordby = req("ordby", "7sellprc")

yyyy1 = "2018"
mm1 = "04"
dd1 = "04"

dim rs
dim oCBrandService

set oCBrandService = new CBrandService
oCBrandService.FRectYYYYMMDD = yyyy1 & "-" & mm1 & "-" & dd1
oCBrandService.FRectMakerid = makerID
oCBrandService.FCurrPage = page
oCBrandService.FPageSize = 100
oCBrandService.FRectOrderBy = ordby
rs = oCBrandService.GetBrandServiceByMeachulList()

class CBrandServiceItem
	public Fyyyymm
	public Fmakerid
	public FoneDaySellItemCnt
	public FoneDaySelltotalPrice
	public FoneDaySellOrderCnt
	public FoneWeekSellItemCnt
	public FoneWeekSelltotalPrice
	public FoneWeekSellOrderCnt
	public FoneMonthSellItemCnt
	public FoneMonthSelltotalPrice
	public FoneMonthSellOrderCnt
	public FthreeMonthSellItemCnt
	public FthreeMonthSelltotalPrice
	public FthreeMonthSellOrderCnt
	public FoneYearSellItemCnt
	public FoneYearSelltotalPrice
	public FoneYearSellOrderCnt

	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

function toClass(rs, i)
	dim result
	'// yyyymmdd, makerid, oneDaySellItemCnt, oneDaySelltotalPrice, oneDaySellOrderCnt, oneWeekSellItemCnt, oneWeekSelltotalPrice, oneWeekSellOrderCnt, oneMonthSellItemCnt, oneMonthSelltotalPrice, oneMonthSellOrderCnt, threeMonthSellItemCnt, threeMonthSelltotalPrice, threeMonthSellOrderCnt, oneYearSellItemCnt, oneYearSelltotalPrice, oneYearSellOrderCnt, regdate, lastupdate
	set result = new CBrandServiceItem
	result.Fyyyymm 			= rs(1,i)
	result.Fmakerid 		= rs(2,i)
	result.FoneDaySellItemCnt 		= rs(3,i)
	result.FoneDaySelltotalPrice 	= rs(4,i)
	result.FoneDaySellOrderCnt 		= rs(5,i)
	result.FoneWeekSellItemCnt 		= rs(6,i)
	result.FoneWeekSelltotalPrice 	= rs(7,i)
	result.FoneWeekSellOrderCnt 		= rs(8,i)
	result.FoneMonthSellItemCnt 		= rs(9,i)
	result.FoneMonthSelltotalPrice 		= rs(10,i)
	result.FoneMonthSellOrderCnt 		= rs(11,i)
	result.FthreeMonthSellItemCnt 		= rs(12,i)
	result.FthreeMonthSelltotalPrice 	= rs(13,i)
	result.FthreeMonthSellOrderCnt 		= rs(14,i)
	result.FoneYearSellItemCnt 		= rs(15,i)
	result.FoneYearSelltotalPrice 	= rs(16,i)
	result.FoneYearSellOrderCnt 	= rs(17,i)

	set toClass = result
end function

dim rowCnt, item, val

function dispUpDnRate(currPrc, prevPrc, currDt, prevDt)
	dim val
	if (currPrc = 0 or prevPrc = 0) then
		dispUpDnRate = "-"
	elseif (1.0 * (currPrc * prevDt) / (prevPrc * currDt) * 100) > 500 then
		dispUpDnRate = "500%+"
	else
		val = (1.0 * (currPrc * prevDt) / (prevPrc * currDt) * 100)
		if (val > 100) then
			val = "<font color='red'>" & FormatNumber(val, 2) & "%" & "</font>"
		elseif (val < 100) then
			val = "<font color='blue'>" & FormatNumber(val, 2) & "%" & "</font>"
		else
			val = FormatNumber(val, 2) & "%"
		end if

		dispUpDnRate = val
	end if
end function

%>

<script language='javascript'>
function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function jsPopDashBoard(makerid) {
    var popwin = window.open("/admin/brandStatic/brandServicePointDashBoard.asp?menupos=4024&makerID=" + makerid,"jsPopDashBoard","width=1400 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
		<td align="left">
	       	기준일 :
			<% DrawOneDateBox yyyy1, mm1, dd1 %>
			&nbsp;
			브랜드ID :
			<input type="text" class="text" name="makerID" value="<%=makerID%>">
		</td>

		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	       	정렬순서 :
			<select class="select" name="ordby">
				<option value="30sellcnt" <%= CHKIIF(ordby="30sellcnt", "selected", "") %>>30일 판매수량</option>
				<option value="30sellprc" <%= CHKIIF(ordby="30sellprc", "selected", "") %>>30일 판매금액</option>
				<option value="7sellcnt" <%= CHKIIF(ordby="7sellcnt", "selected", "") %>>7일 판매수량</option>
				<option value="7sellprc" <%= CHKIIF(ordby="7sellprc", "selected", "") %>>7일 판매금액</option>
				<option value="1sellcnt" <%= CHKIIF(ordby="1sellcnt", "selected", "") %>>1일 판매수량</option>
				<option value="1sellprc" <%= CHKIIF(ordby="1sellprc", "selected", "") %>>1일 판매금액</option>
			</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60" rowspan="2">년월</td>
		<td width="250" rowspan="2">브랜드</td>
		<td width="320" colspan="4">1일 판매내역</td>
		<td width="320" colspan="4">7일 판매내역</td>
		<td width="320" colspan="4">30일 판매내역</td>
		<td width="320" colspan="4">90일 판매내역</td>
		<td width="240" colspan="3">360일 판매내역</td>
		<td rowspan="2">비고</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
		<td width="80">등락율<br />(금액)</td>
		<td width="80">판매수량</td>
		<td width="80">판매금액</td>
		<td width="80">주문건수</td>
	</tr>
	<%
	If IsArray(rs) Then
		rowCnt = UBound(rs,2) + 1
		For i = 0 To UBound(rs,2)
			set item = toClass(rs, i)
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= item.Fyyyymm %></td>
		<td><a href="javascript:jsPopDashBoard('<%= item.Fmakerid %>')"><%= item.Fmakerid %></a></td>
		<td><%= FormatNumber(item.FoneDaySellItemCnt,0) %></td>
		<td><%= FormatNumber(item.FoneDaySelltotalPrice,0) %></td>
		<td><%= FormatNumber(item.FoneDaySellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(item.FoneDaySelltotalPrice, item.FoneWeekSelltotalPrice, 1, 7) %></td>
		<td><%= FormatNumber(item.FoneWeekSellItemCnt,0) %></td>
		<td><%= FormatNumber(item.FoneWeekSelltotalPrice,0) %></td>
		<td><%= FormatNumber(item.FoneWeekSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(item.FoneWeekSelltotalPrice, item.FoneMonthSelltotalPrice, 7, 30) %></td>
		<td><%= FormatNumber(item.FoneMonthSellItemCnt,0) %></td>
		<td><%= FormatNumber(item.FoneMonthSelltotalPrice,0) %></td>
		<td><%= FormatNumber(item.FoneMonthSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(item.FoneMonthSelltotalPrice, item.FthreeMonthSelltotalPrice, 30, 90) %></td>
		<td><%= FormatNumber(item.FthreeMonthSellItemCnt,0) %></td>
		<td><%= FormatNumber(item.FthreeMonthSelltotalPrice,0) %></td>
		<td><%= FormatNumber(item.FthreeMonthSellOrderCnt,0) %></td>
		<td><%= dispUpDnRate(item.FthreeMonthSelltotalPrice, item.FoneYearSelltotalPrice, 90, 360) %></td>
		<td><%= FormatNumber(item.FoneYearSellItemCnt,0) %></td>
		<td><%= FormatNumber(item.FoneYearSelltotalPrice,0) %></td>
		<td><%= FormatNumber(item.FoneYearSellOrderCnt,0) %></td>
		<td></td>
	</tr>
	<%
		next
	end if
	%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19" align="center">
		<% if oCBrandService.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCBrandService.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCBrandService.StartScrollPage to oCBrandService.FScrollCount + oCBrandService.StartScrollPage - 1 %>
			<% if i>oCBrandService.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCBrandService.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
