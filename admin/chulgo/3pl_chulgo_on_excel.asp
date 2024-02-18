<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� 3pl �ֹ�
' History : 2017.03.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/db_TPLopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim shopid , yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2,i, reloading
dim fromDate,toDate, page, tplcompanyid
dim totsuplyprice , totprofit , totprofit2 , custa ,makerid
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)
	page = requestcheckvar(getNumeric(request("page")),10)
	tplcompanyid = requestcheckvar(request("tplcompanyid"),32)
	reloading = requestcheckvar(request("reloading"),2)

if page = "" then page = 1
if reloading="" and tplcompanyid = "" then tplcompanyid="tplithinkso"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

dim oorder
set oorder = new Cchulgoitemlist
	oorder.FPageSize = 3000
	oorder.FCurrPage = page
	oorder.FRectStartdate = fromDate
	oorder.FRectEnddate = toDate
	oorder.FRecttplcompanyid = tplcompanyid
if tplcompanyid="tpliconic" or tplcompanyid="tplmmmg" or tplcompanyid="tplparagon" then
	oorder.fETC3plculgolist
else
	oorder.fonline3plculgolist
end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN3PL_ONLINE_ORDER" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="gray" border=1>
<tr bgcolor="#FFFFFF">
	<td colspan="6">
		�˻���� : <b><%= oorder.FTotalCount %></b>
		<!--&nbsp;������ : <b><%= page %>/ <%= oorder.FTotalPage %></b>-->
		&nbsp;&nbsp;���ִ� 3õ�Ǳ��� ���� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��¥</td>
	<td>����Ʈ</td>	
	<td>���ֹ��Ǽ�</td>	
	<td>�ֹ��Ǽ�(+)</td>
	<td>�ֹ��Ǽ�(-)</td>
	<td>��ǰ����</td>
	<td>�����Ǽ�</td>
	<td>���</td>
</tr>
<% if oorder.FresultCount > 0 then %>
	<% for i=0 to oorder.FresultCount-1 %>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= oorder.FItemList(i).fyyyymmdd %>
		</td>
		<td class='txt'>
			<%= oorder.FItemList(i).fsitename %>
		</td>		
		<td align="right">
			<%= oorder.FItemList(i).fordercnt %>
		</td>	
		<td align="right">
			<%= oorder.FItemList(i).forderpluscnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).forderminuscnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).fitemcnt %>
		</td>
		<td align="right">
			<%= oorder.FItemList(i).fitemcnt2 %>
		</td>
		<td></td>
	</tr>
	<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>

</table>

</body>
</html>

<%
set oorder = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->