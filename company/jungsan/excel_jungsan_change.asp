<%@ language=vbscript %>
<% option explicit %>
<%

response.expires = -1
response.AddHeader "Pragma", "no-cache"
response.AddHeader "cache-control", "no-store"

Response.ContentType = "application/vnd.ms-excel"

%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/companyjungsancls.asp"-->
<%

'==============================================================================

dim masterid,ix
dim bufsum,deasangsum,amountsum
dim junsandate,junsandatearr,junsandatestr

masterid = request("masterid")
junsandate = request("junsandate")

dim ijungsan
set ijungsan = new CUpcheJungSan
ijungsan.getOldDefaultInfo masterid
ijungsan.FMasterid = masterid
ijungsan.PartnerXLOldJungSanDeasangList

junsandatearr = split(junsandate,"-")
if junsandatearr(1) < 10 then
junsandatearr(1) = "0" + junsandatearr(1)
end if

junsandatestr = junsandatearr(0) + junsandatearr(1)

Response.AddHeader "Content-Disposition","attachment;filename=" & junsandatestr & "_jungsan_list.xls"
'==============================================================================

%>
<!doctype html public "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<body>
<table border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100" align="center">주문번호</td>
		<td width="80" align="center">UserID</td>
		<td width="65" align="center">구매자</td>
		<td width="72" align="center">결제금액</td>
		<td width="72" align="center">포장.배송료</td>
		<td width="90" align="center">정산대상금액</td>
		<td width="90" align="center">정산금액</td>
	</tr>
<% if ijungsan.FresultCount<1 then %>
<% else %>
	<% for ix=0 to ijungsan.FresultCount-1 %>
	<tr class="a">
		<td align="center"><%= ijungsan.FJungSanList(ix).FOrderSerial %></td>
		<% if ijungsan.FJungSanList(ix).FUserID<>"" then %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FUserID %></td>
		<% else %>
		<td align="center">&nbsp;</td>
		<% end if %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FBuyName %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FSubTotalPrice %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FBeasongPay %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).FDeasangPay %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="center"><%= ijungsan.FJungSanList(ix).FDeasangPay %></td>
	</tr>
	<% next %>
<% end if %>
</table>
<%
set ijungsan = nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->


