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
strSql = " db_datamart.dbo.usp_Ten_DeliveryDelayList_List ('" & yyyy1 & "-" & mm1 & "', '" & makerID & "', '" & onlyChkUpcheOnly & "')"

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
	public FbaljuCnt
	public FstockoutCnt
	public FdelayCnt
	public FbaditemCnt
	public FerrdeliveryCnt
	public FchulgoCnt
	public FchulgoNDaySum
	public FrealOverNDaySum
	public FfalsehoodSongjangCnt

	public function GetSUM
		GetSUM = (FstockoutCnt + FdelayCnt + FbaditemCnt + FerrdeliveryCnt)
	end function

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
end class

function toClass(rs, i)
	dim result
	'// yyyymm, makerid, baljuCnt, stockoutCnt, delayCnt, baditemCnt, errdeliveryCnt, regdate, lastupdate
	'// , chulgoCnt, chulgoNDaySum, realOverNDaySum, falsehoodSongjangCnt
	set result = new CBrandServiceItem
	result.Fyyyymm 			= rs(0,i)
	result.Fmakerid 		= rs(1,i)
	result.FbaljuCnt 		= rs(2,i)
	result.FstockoutCnt 	= rs(3,i)
	result.FdelayCnt 		= rs(4,i)
	result.FbaditemCnt 		= rs(5,i)
	result.FerrdeliveryCnt 	= rs(6,i)
	result.FchulgoCnt 		= rs(7,i)
	result.FchulgoNDaySum 	= rs(8,i)
	result.FrealOverNDaySum 		= rs(9,i)
	result.FfalsehoodSongjangCnt 	= rs(10,i)

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
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="onlyChkUpcheOnly" value="Y" <%= CHKIIF(onlyChkUpcheOnly="Y", "checked", "") %>> 확인대상 브랜드(고객불만 취소(반품) 비율 5% 이상 또는 허위송장 비율 5% 이상 있는 브랜드)만
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<pre>
* 발주건수는 브랜드별 상품종류수 입니다.(한 주문에 3가지 상품을 주문하면 3으로 카운트합니다.)
* 1월에 발주된 주문에 대해, 2월에 취소가 이루어지면 년월이 각각 분리됩니다.
* 평균배송소요일은 업체통보 이후 택배사집하가 되는 때까지의 배송일수 입니다.
* 허위송장은 결제일 이전 택배사집하가 이루어지거나, 배송일과 택배사집하일의 차이가 3일 이상 나거나, 5일간 배송조회가 안되는 건입니다.
</pre>

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="60">
			년월
		</td>
		<td rowspan="2" width="250">브랜드</td>
		<td width="80" rowspan="2">총발주건수<br>(업체배송)</td>
        <td colspan="6">고객불만 취소(반품)건수</td>
        <td colspan="4" width="80">평균배송소요일</td>
		<td rowspan="2" width="80"><b>서비스지수</b></td>
		<td rowspan="2">비고</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="80">품절</td>
		<td width="80">배송지연</td>
		<td width="80">상품불량</td>
		<td width="80">오배송</td>
		<td width="80">합계</td>
		<td width="80"><b>비율<br />(발주대비)</b></td>
		<td width="80">출고건수</td>
		<td width="80">배송일기준</td>
		<td width="80">송장조회기준</td>
		<td width="80">허위송장건수</td>
	</tr>
	<%
	If IsArray(rs) Then
		rowCnt = UBound(rs,2) + 1
		For i = 0 To UBound(rs,2)
			set item = toClass(rs, i)

			totbaljuCnt = totbaljuCnt + item.FbaljuCnt
			totstockoutCnt = totstockoutCnt + item.FstockoutCnt
			totdelayCnt = totdelayCnt + item.FdelayCnt
			totbaditemCnt = totbaditemCnt + item.FbaditemCnt
			toterrdeliveryCnt = toterrdeliveryCnt + item.FerrdeliveryCnt
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= item.Fyyyymm %></td>
		<td><a href="javascript:jsPopDashBoard('<%= item.Fmakerid %>')"><%= item.Fmakerid %></a></td>
		<td><%= item.FbaljuCnt %></td>
		<td><%= item.FstockoutCnt %></td>
		<td><%= item.FdelayCnt %></td>
		<td><%= item.FbaditemCnt %></td>
		<td><%= item.FerrdeliveryCnt %></td>
		<td><%= item.GetSUM %></td>
		<td>
			<%
			if item.FbaljuCnt > 0 then
				val = Round((1.0 * item.GetSUM / item.FbaljuCnt * 100), 1)
				if (val >= 5) then
					response.write "<font color='red'><b>" & val & "%</b></font>"
				else
					response.write val & "%"
				end if
			else
				response.write "-"
			end if
			%>
		</td>
		<td><%= item.FchulgoCnt %></td>
		<% if item.FchulgoCnt > 0 then %>
		<td><%= Round(1.0*item.FchulgoNDaySum/item.FchulgoCnt,1) %></td>
		<td><%= Round(1.0*(item.FchulgoNDaySum+item.FrealOverNDaySum)/item.FchulgoCnt,1) %></td>
		<td>
			<% if (item.FfalsehoodSongjangCnt > 0) then %>
			<font color="red"><b><%= item.FfalsehoodSongjangCnt %></b></font>
			<% else %>
			<%= item.FfalsehoodSongjangCnt %>
			<% end if %>
		</td>
		<% else %>
		<td>-</td>
		<td>-</td>
		<td>-</td>
		<% end if %>
		<td></td>
		<td></td>
	</tr>
	<%
		next
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td><%= totbaljuCnt %></td>
		<td><%= totstockoutCnt %></td>
		<td><%= totdelayCnt %></td>
		<td><%= totbaditemCnt %></td>
		<td><%= toterrdeliveryCnt %></td>
		<td>
			<%= (totstockoutCnt + totdelayCnt + totbaditemCnt + toterrdeliveryCnt) %>
		</td>
		<td>
			<%
			if totbaljuCnt > 0 then
				val = Round((1.0 * (totstockoutCnt + totdelayCnt + totbaditemCnt + toterrdeliveryCnt) / totbaljuCnt * 100), 1)
				if (val >= 5) then
					response.write "<font color='red'><b>" & val & "%</b></font>"
				else
					response.write val & "%"
				end if
			else
				response.write "-"
			end if
			%>
		</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<%
	end if
	%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
