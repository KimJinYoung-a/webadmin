<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/dayRate/dayRateCls.asp"-->
<%
'####################################################
' Description : 해외몰 환율별 주문리스트
' History : 2016-10-24 김진영 생성
'####################################################
%>
<%
Dim oRate, page, i, mallgubun, sDt, eDt, moneyUnit
Dim TotMallSumprice, TotMallMaxShipping, TotMallTotprice, TotKRsellPrice, TotKRshipping, TotKRtotPrice
page    	= request("page")
mallgubun	= request("mallgubun")
sDt			= request("sDt")
eDt			= request("eDt")

If page = "" Then page = 1
If mallgubun = "" Then mallgubun = "cnglob10x10"
If sDt = "" Then sDt = Left(Date(), 8) & "01"
If eDt = "" Then eDt = Date()

If CDate(sdt) < "2016-06-01" Then
	response.write "<script>alert('시작일은 2016-06-01 이후로 설정하세요');history.back(-1);</script>"
End If

If CDate(edt) < "2016-06-01" Then
	response.write "<script>alert('종료일은 2016-06-01 이후로 설정하세요');history.back(-1);</script>"
End If

Select Case mallgubun
	Case "cnglob10x10"		moneyUnit = "USD"
	Case "cnhigo"			moneyUnit = "CNY"
	Case "11stmy"			moneyUnit = "MYR"
	Case "etsy"				moneyUnit = "USD"
	Case "zilingo"			moneyUnit = "SGD"
End Select

Set oRate = new CRate
	oRate.FCurrPage			= page
	oRate.FPageSize			= 30
	oRate.FRectMallGubun	= mallgubun
	oRate.FRectSDt			= sDt
	oRate.FRectEDt			= eDt
	oRate.getdayRateOrderList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function goToMall(v){
    frm.mallgubun.value = v;
    frm.page.value = 1;
    frm.submit();
}
</script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%= page %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		Mall선택 : 
		<select class="select" name="mallgubun" onchange="goToMall(this.value);">
			<option value="cnglob10x10" <%= chkiif(mallgubun="cnglob10x10", "selected", "") %>>cnglob10x10</option>
			<option value="cnhigo" <%= chkiif(mallgubun="cnhigo", "selected", "") %>>cnhigo</option>
			<option value="11stmy" <%= chkiif(mallgubun="11stmy", "selected", "") %>>11stmy</option>
			<option value="etsy" <%= chkiif(mallgubun="etsy", "selected", "") %>>etsy</option>
			<option value="zilingo" <%= chkiif(mallgubun="zilingo", "selected", "") %>>zilingo</option>
		</select>&nbsp;
		출고일 기간 : 
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<br />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oRate.FTotalCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="10%">10x10 주문번호</td>
	<td width="10%">카드승인번호</td>
	<td width="10%">출고일자</td>
	<td width="7%">출고일환율</td>
	<td width="10%">상품금액(<%= moneyUnit %>)</td>
	<td width="5%">배송비(<%= moneyUnit %>)</td>
	<td width="10%">합계(<%= moneyUnit %>)</td>
	<td width="10%">상품금액(KRW)</td>
	<td width="10%">배송비(KRW)</td>
	<td width="10%">합계(KRW)</td>
	<td width="8%">송장번호</td>
	<td >상품무게</td>
</tr>
<% For i=0 to oRate.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" height="25">
	<td><%= oRate.FItemList(i).FOrderserial %></td>
	<td><%= oRate.FItemList(i).FAuthcode %></td>
	<td><%= Left(oRate.FItemList(i).FBeadaldate, 10) %></td>
	<td align="right">
		<%
			Select Case mallgubun
				Case "cnglob10x10"		response.write oRate.FItemList(i).FUSD
				Case "cnhigo"			response.write oRate.FItemList(i).FCNY
				Case "11stmy"			response.write oRate.FItemList(i).FMYR
				Case "etsy"				response.write oRate.FItemList(i).FUSD
				Case "zilingo"			response.write oRate.FItemList(i).FSGD
			End Select
		%>
	</td>
	<td align="right"><%= oRate.FItemList(i).FMallSumprice %></td>
	<td align="right"><%= oRate.FItemList(i).FMallMaxShipping %></td>
	<td align="right"><%= oRate.FItemList(i).FMallTotprice %></td>
	<td align="right">
		<%
			If isnull(oRate.FItemList(i).FKRsellPrice) Then
				response.write "미등록"
			Else
				response.write FormatNumber(oRate.FItemList(i).FKRsellPrice ,0)
			End If
		%>
	</td>
	<td align="right">
		<%
			If isnull(oRate.FItemList(i).FKRshipping) Then
				response.write "미등록"
			Else
				response.write FormatNumber(oRate.FItemList(i).FKRshipping ,0)
			End If
		%>
	</td>
	<td align="right">
		<%
			If isnull(oRate.FItemList(i).FKRtotPrice) Then
				response.write "미등록"
			Else
				response.write FormatNumber(oRate.FItemList(i).FKRtotPrice ,0)
			End If
		%>
	</td>
	<td><%= oRate.FItemList(i).FDeliverno %></td>
	<td><%= oRate.FItemList(i).FitemWeigth %></td>
</tr>
<%
	TotMallSumprice			= TotMallSumprice + oRate.FItemList(i).FMallSumprice
	TotMallMaxShipping		= TotMallMaxShipping + oRate.FItemList(i).FMallMaxShipping
	TotMallTotprice			= TotMallTotprice + oRate.FItemList(i).FMallTotprice
	TotKRsellPrice			= TotKRsellPrice + CDBL(FormatNumber(oRate.FItemList(i).FKRsellPrice, 0))
	TotKRshipping			= TotKRshipping + CDBL(FormatNumber(oRate.FItemList(i).FKRshipping ,0))
	TotKRtotPrice			= TotKRtotPrice + CDBL(FormatNumber(oRate.FItemList(i).FKRtotPrice ,0))
%>
<% Next %>
<tr align="center" bgcolor="#FFFFE4" height="25">
	<td>합계</td>
	<td colspan="3"></td>
	<td align="right"><%= TotMallSumprice %></td>
	<td align="right"><%= TotMallMaxShipping %></td>
	<td align="right"><%= TotMallTotprice %></td>
	<td align="right"><%= FormatNumber(TotKRsellPrice, 0) %></td>
	<td align="right"><%= FormatNumber(TotKRshipping, 0) %></td>
	<td align="right"><%= FormatNumber(TotKRtotPrice, 0) %></td>
	<td></td>
	<td></td>
</tr>
<tr height="20">
    <td colspan="12" align="center" bgcolor="#FFFFFF">
        <% if oRate.HasPreScroll then %>
		<a href="javascript:goPage('<%= oRate.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oRate.StartScrollPage to oRate.FScrollCount + oRate.StartScrollPage - 1 %>
    		<% if i>oRate.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oRate.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<script language="javascript">
	var CAL_Start = new Calendar({
		inputField : "sDt", trigger    : "sDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "eDt", trigger    : "eDt_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
</script>
<% Set oRate = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->