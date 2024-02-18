<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 원가정산리스트 세금계산서 발급 엑셀다운로드
' History : 2023.09.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<%
dim i, research, page, ExcDel, productidx, yyyy1, mm1, yyyy2, mm2, dt, makerid, purchasetype, groupid, company_name, ppGubun
dim reportIdx, selectfinishflag, itemid, arrLIst
	page = requestCheckVar(getNumeric(request("page")),8)
	productidx = requestCheckVar(trim(getNumeric(request("productidx"))),8)
	reportIdx = requestCheckVar(trim(getNumeric(request("reportIdx"))),8)
	ExcDel = requestCheckVar(request("ExcDel"),1)
	research = requestCheckVar(request("research"),1)
	yyyy1    = requestCheckVar(request("yyyy1"),4)
	mm1      = requestCheckVar(request("mm1"),2)
	yyyy2    = requestCheckVar(request("yyyy2"),4)
	mm2      = requestCheckVar(request("mm2"),2)
	makerid = requestCheckVar(trim(request("makerid")),32)
	purchasetype = requestCheckVar(request("purchasetype"),2)
	groupid  = requestCheckVar(trim(request("groupid")),6)
	company_name  = requestCheckVar(trim(request("company_name")),64)
	ppGubun = requestCheckVar(trim(request("ppGubun")),32)
	selectfinishflag = requestCheckVar(request("selectfinishflag"),10)
	itemid      = requestCheckvar(request("itemid"),1500)

if page = "" then page = "1"
if ExcDel = "" and research="" then ExcDel = "Y"
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if
if yyyy2="" then
	dt = dateserial(year(Now),month(now),1)
	yyyy2 = Left(CStr(dt),4)
	mm2 = Mid(CStr(dt),6,2)
end if
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oPurchasedJungsan
set oPurchasedJungsan = new CPurchasedJungsan
	oPurchasedJungsan.FCurrPage = page
	oPurchasedJungsan.Fpagesize = 1000000
    oPurchasedJungsan.FRectExcDel = ExcDel
	oPurchasedJungsan.FRectproductidx = productidx
	oPurchasedJungsan.FRectYYYYMM1 = yyyy1 + "-" + mm1
	oPurchasedJungsan.FRectYYYYMM2 = yyyy2 + "-" + mm2
	oPurchasedJungsan.FRectmakerid = makerid
	oPurchasedJungsan.FRectpurchasetype = purchasetype
	oPurchasedJungsan.FRectgroupid = groupid
	oPurchasedJungsan.FRectcompany_name = company_name
	oPurchasedJungsan.FRectppGubun = ppGubun
	oPurchasedJungsan.FRectreportIdx = reportIdx
	oPurchasedJungsan.FRectItemid       = itemid
	oPurchasedJungsan.FRectFinishFlag = selectfinishflag
	oPurchasedJungsan.GetPurchasedJungsanMasterListNotPaging

if oPurchasedJungsan.FTotalCount>0 then
    arrLIst=oPurchasedJungsan.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENPurchasedProductJungsan" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="13">
		검색결과 : <b><%= oPurchasedJungsan.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>원가상세idx</td>
    <td>정산월</td>
    <td>그룹코드</td>
	<td>브랜드ID</td>
    <td>사업자명</td>
    <td>비용구분</td>
    <td>원가총액</td>
    <td>관련품의IDX</td>
    <td>세금계산서상태</td>
    <td>세금계산서등록일</td>
	<td>발행일</td>
    <td>비고</td>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrLIst(0,i) %></td>
	<td><%= arrLIst(1,i) %></td>
	<td class="txt"><%= arrLIst(2,i) %></td>
	<td><%= arrLIst(3,i) %></td>
	<td class="txt"><%= arrLIst(18,i) %></td>
	<td><%= arrLIst(4,i) %></td>
	<td><%= arrLIst(5,i) %></td>
	<td><%= FormatNumber(arrLIst(6,i), 0) %></td>
	<td class="txt"><%= arrLIst(7,i) %></td>
    <td><%= GetStateName(arrLIst(10,i)) %></td>
	<td><%= arrLIst(13,i) %></td>
	<td><%= arrLIst(12,i) %></td>
	<td>
		<% if IsElecTaxExists(arrLIst(14,i),arrLIst(10,i)) then %>
		<% else %>
			<%= arrLIst(16,i) %>
		<% end if %>
	</td>

</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oPurchasedJungsan = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->