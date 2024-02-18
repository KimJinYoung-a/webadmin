<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim yyyy1, mm1
Dim makerID, onlyChkUpcheOnly
dim i

yyyy1	= req("yyyy1", Left(DateAdd("m", -1, Now()),4))
mm1		= req("mm1", Mid(DateAdd("m", -1, Now()),6,2))
makerID = req("makerID", "")
onlyChkUpcheOnly = req("onlyChkUpcheOnly", "")


Dim strSql
strSql = " db_datamart.dbo.usp_Ten_DeliveryClaimList_List ('" & yyyy1 & "-" & mm1 & "', '" & makerID & "')"

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

Dim rs
If Not db3_rsget.EOF Then
	rs = db3_rsget.getRows()
End If
db3_rsget.close

class CBrandServiceItem
	public Fyyyymm
	public Fmakerid
	public FtotCnt
	public FtotSum
	public FdelayCnt
	public FdelaySum
	public FstockoutCnt
	public FstockoutSum
	public FerrdeliveryCnt
	public FerrdeliverySum
	public FitemregerrCnt
	public FitemregerrSum
	public FupcheerrCnt
	public FupcheerrSum
	public FetcupcheerrCnt
	public FetcupcheerrSum

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

function toClass(rs, i)
	dim result
	'// totCnt, totSum, delayCnt, delaySum, stockoutCnt, stockoutSum, errdeliveryCnt, errdeliverySum, itemregerrCnt, itemregerrSum, upcheerrCnt, upcheerrSum
	set result = new CBrandServiceItem
	result.Fyyyymm 			= rs(0,i)
	result.Fmakerid 		= rs(1,i)
	result.FtotCnt 			= rs(2,i)
	result.FtotSum 			= rs(3,i)
	result.FdelayCnt 		= rs(4,i)
	result.FdelaySum 		= rs(5,i)
	result.FstockoutCnt 	= rs(6,i)
	result.FstockoutSum 	= rs(7,i)
	result.FerrdeliveryCnt 	= rs(8,i)
	result.FerrdeliverySum 	= rs(9,i)
	result.FitemregerrCnt 	= rs(10,i)
	result.FitemregerrSum 	= rs(11,i)
	result.FupcheerrCnt 	= rs(12,i)
	result.FupcheerrSum 	= rs(13,i)
	result.FetcupcheerrCnt 	= rs(14,i)
	result.FetcupcheerrSum 	= rs(15,i)

	set toClass = result
end function

dim rowCnt, item, val
dim totbaljuCnt, totstockoutCnt, totdelayCnt, totbaditemCnt, toterrdeliveryCnt

%>

<script language='javascript'>
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
	       	년월 :
			<% DrawYMBox yyyy1,mm1 %>
			&nbsp;
			브랜드ID :
			<input type="text" class="text" name="makerID" value="<%=makerID%>">
		</td>

		<td rowspan="2" width="80" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="60">
			년월
		</td>
		<td rowspan="2" width="250">브랜드</td>
		<td width="80" rowspan="2">총건수<br>(업체배송)</td>
		<td width="80" rowspan="2">총비용<br>(업체배송)</td>
        <td colspan="6">클래임 건수</td>
		<td colspan="6">클래임 비용</td>
		<td rowspan="2">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">배송지연</td>
        <td width="80">품절</td>
		<td width="80">오배송</td>
		<td width="80">상품등록오류</td>
		<td width="80">업체대응불량</td>
		<td width="80">기타업체과실</td>
		<td width="80">배송지연</td>
        <td width="80">품절</td>
		<td width="80">오배송</td>
		<td width="80">상품등록오류</td>
		<td width="80">업체대응불량</td>
		<td width="80">기타업체과실</td>
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
		<td><%= item.FtotCnt %></td>
		<td><%= FormatNumber(item.FtotSum,0) %></td>
		<td><%= item.FdelayCnt %></td>
		<td><%= item.FstockoutCnt %></td>
		<td><%= item.FerrdeliveryCnt %></td>
		<td><%= item.FitemregerrCnt %></td>
		<td><%= item.FupcheerrCnt %></td>
		<td><%= item.FetcupcheerrCnt %></td>
		<td><%= FormatNumber(item.FdelaySum,0) %></td>
		<td><%= FormatNumber(item.FstockoutSum,0) %></td>
		<td><%= FormatNumber(item.FerrdeliverySum,0) %></td>
		<td><%= FormatNumber(item.FitemregerrSum,0) %></td>
		<td><%= FormatNumber(item.FupcheerrSum,0) %></td>
		<td><%= FormatNumber(item.FetcupcheerrSum,0) %></td>
		<td></td>
	</tr>
	<%
		next
	end if
	%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
