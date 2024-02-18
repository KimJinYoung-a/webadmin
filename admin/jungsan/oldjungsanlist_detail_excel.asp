<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsancls.asp"-->
<%
'#############################################################
'	Description : 지난정상조회 상세 Excel파일 생성
'	History		: 2023.07.05 생성; 허진원
'#############################################################
%>
<%
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_OLDCALUCULATE_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<%

dim page
dim ijungsan
Dim masterid,extsitename

extsitename = request("extsitename")

masterid = request("masterid")

page = request("page")
if (page="") then page=1

set ijungsan = new CUpcheJungSan

ijungsan.FcurrPage = page
ijungsan.FPageSize=9000
ijungsan.getOldDefaultInfo masterid

ijungsan.FMasterid = masterid
ijungsan.FrectSiteName = extsitename
ijungsan.PartnerOldDetailJungSanDeasangList

dim ix
dim bufsum, deasangsum, amountsum
bufsum =0
deasangsum =0
amountsum =0
%>
<table border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="8">
		<table border="0" cellspacing="0" cellpadding="0" class="a">
		<tr>
			<td>* 커미션 : </td>
			<td><Font color="#3333FF"><%= CDbl(ijungsan.FCommission)*100 %> %</font></td>
		</tr>
		<tr>
			<td>* 총 건수 : </td>
			<td><Font color="#3333FF"><%= FormatNumber(ijungsan.FTotalCount,0) %></font></td>
		</tr>
		<tr>
			<td>* 정산대상 금액 : </td>
			<td ><% = FormatNumber(ijungsan.FTotalJungsan,0)  %></td>
		</tr>
		<tr>
			<td>* 정산예정 금액 : </td>
			<td ><% = FormatNumber(ijungsan.FTotalJungsansum,0)  %></td>
		</tr>
		<tr>
			<td>* 기타사항 : </td>
			<td ><% = ijungsan.FEtcStr  %></td>
		</tr>

		</table>
	</td>
</tr>
<tr >
	<td align="center">주문번호</td>
	<td align="center">UserID</td>
	<td align="center">구매자</td>
	<td align="center">결제금액</td>
	<td align="center">포장.배송료</td>
	<td align="center">정산대상금액</td>
	<td align="center">정산금액</td>
	<td align="center">업체주문번호</td>
</tr>
<% if ijungsan.FresultCount<1 then %>
<tr>
	<td colspan="8" align="center">[검색결과가 없습니다.]</td>
</tr>
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
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FSubTotalPrice,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FBeasongPay,0) %></td>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).FDeasangPay,0) %></td>
		<%
			bufsum = CDbl(ijungsan.FJungSanList(ix).FDeasangPay)
			deasangsum = deasangsum + bufsum
			amountsum = amountsum + bufsum* CDbl(ijungsan.FCommission)
		 %>
		<td align="center"><%= FormatNumber(ijungsan.FJungSanList(ix).Fjungsansum,0) %></td>
		<td align="center"><%= ijungsan.FJungSanList(ix).Fauthcode & ijungsan.FJungSanList(ix).Fpaygatetid %></td>
	</tr>
	<%
			if (ix mod 100)=0 then Response.Flush
		next
	%>
<% end if %>
</table>
<%
set ijungsan = nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
