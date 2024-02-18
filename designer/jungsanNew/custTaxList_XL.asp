<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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
oTax.FCurrPage = 1
oTax.FPageSize = 200						'// 최대 200개
oTax.FRectSdate = startDate
oTax.FRectEdate = endDate
oTax.FRectSupplyGroupID = groupid			'// 그룹아이디 밑에 모든 브랜드 발행내역 표시
oTax.GetTaxListUpche

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>고객 세금계산서 발행내역(<%= CStr(yyyy1) %>-<% CStr(mm1) %>)</title>
<table>
	<tr>
		<td>번호</td>
		<td>발행일자</td>
		<td>공급자사업자등록번호</td>
		<td>종사업장번호</td>
		<td>상호</td>
		<td>대표자명</td>
		<td>공급받는자사업자등록번호</td>
		<td>종사업장번호</td>
		<td>상호</td>
		<td>대표자명</td>
		<td>합계금액</td>
		<td>공급가액</td>
		<td>세액</td>
		<td>전자세금계산서분류</td>
		<td>전자세금계산서종류</td>
		<td>공급자 이메일</td>
		<td>공급받는자 이메일</td>
		<td>품목명</td>
	</tr>
	<% for i=0 to oTax.FResultCount - 1 %>
	<tr>
		<td><%= oTax.FTaxList(i).FtaxIdx %></td>
		<td><%= FormatDate(oTax.FTaxList(i).FisueDate,"0000-00-00") %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiNo %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiSubNo %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiName %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiCEOName %></td>
		<td><%= oTax.FTaxList(i).FBusiNo %></td>
		<td><%= oTax.FTaxList(i).FbusiSubNo %></td>
		<td><%= oTax.FTaxList(i).FBusiName %></td>
		<td><%= oTax.FTaxList(i).FbusiCEOName %></td>
		<td><%= oTax.FTaxList(i).FtotalPrice %></td>
		<td><%= oTax.FTaxList(i).FtotalPrice - oTax.FTaxList(i).FtotalTax %></td>
		<td><%= oTax.FTaxList(i).FtotalTax %></td>
		<td>세금계산서</td>
		<td>위수탁</td>
		<td><%= oTax.FTaxList(i).FsupplyRepEmail %></td>
		<td><%= oTax.FTaxList(i).FrepEmail %></td>
		<td><%= oTax.FTaxList(i).Fitemname %></td>
	</tr>
	<% next %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
