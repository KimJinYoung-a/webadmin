<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������긮��Ʈ ���ݰ�꼭 �߱� �����ٿ�ε�
' History : 2023.09.15 �ѿ�� ����
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
Response.Buffer = true    '���ۻ�뿩��
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
		�˻���� : <b><%= oPurchasedJungsan.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>������idx</td>
    <td>�����</td>
    <td>�׷��ڵ�</td>
	<td>�귣��ID</td>
    <td>����ڸ�</td>
    <td>��뱸��</td>
    <td>�����Ѿ�</td>
    <td>����ǰ��IDX</td>
    <td>���ݰ�꼭����</td>
    <td>���ݰ�꼭�����</td>
	<td>������</td>
    <td>���</td>
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
    Response.Flush		' ���۸��÷���
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oPurchasedJungsan = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->