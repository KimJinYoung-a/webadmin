<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ������ ��꼭 �������
' History : 2012.09.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<%
dim i
dim yyyymm : yyyymm=request("YYYYMM")
dim stDate,edDate
dim arrList, intLoop

if (yyyymm="") then yyyymm=LEFT(now(),7)
stDate = yyyymm+"-01"
edDate = Left(DateAdd("m",1,stDate),10)

Dim otax
Set otax = new CEsero
otax.FSDate=stDate
otax.FEDate=edDate

arrList = otax.getMonthTaxList()

Response.Buffer = true    '���ۻ�뿩��
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_hometax" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="gray">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=otax.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="ffffff">
	<td>��꼭��ȣ</td>
	<td>������</td>
	<td>����ó<Br>����ڹ�ȣ</td>
	<td>�������</td>
	<td>����ó</td>
	<td>����ó<Br>��ǥ��</td>
	<td>����ó<Br>���EMAIL</td>
	<td>����ó</td>
	<td>����ó<Br>ȸ���</td>
	<td>����ó<Br>��ǥ��</td>
	<td>����ó<Br>�����EMAIL</td>
	<td>�հ�</td>
	<td>���ް�</td>
	<td>�ΰ���</td>
	<td>���Ա���</td>
	<td>���⿩��</td>
	<td>��������</td>
	<td>���</td>
	<td>ǰ��</td>
	<td>����ι�</td>
</tr>

<%
IF isArray(arrList) THEN

For intLoop = 0 To UBound(arrList,2)

%>
<tr align="center" bgcolor="#FFFFFF">
	<td class='txt'><%= arrList(0,intLoop) %></td>
	<td><%= arrList(1,intLoop) %></td>
	<td><%= arrList(2,intLoop) %></td>
	<td><%= arrList(3,intLoop) %></td>
	<td><%= arrList(4,intLoop) %></td>
	<td><%= arrList(5,intLoop) %></td>
	<td><%= arrList(6,intLoop) %></td>
	<td><%= arrList(7,intLoop) %></td>
	<td><%= arrList(8,intLoop) %></td>
	<td><%= arrList(9,intLoop) %></td>
	<td><%= arrList(10,intLoop) %></td>
	<td align="right"><%= FormatNumber(arrList(11,intLoop),0) %></td>
	<td align="right"><%= FormatNumber(arrList(12,intLoop),0) %></td>
	<td align="right"><%= FormatNumber(arrList(13,intLoop),0) %></td>
	<td><%= arrList(14,intLoop) %></td>
	<td><%= arrList(15,intLoop) %></td>
	<td><%= arrList(16,intLoop) %></td>
	<td><%= arrList(17,intLoop) %></td>
	<td><%= arrList(18,intLoop) %></td>
	<td><%= arrList(19,intLoop) %></td>
</tr>

<%
	if intLoop mod 50 = 0 then
		Response.Flush		' ���۸��÷���
	end if

Next
%>

<% ELSE %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

</body>
</html>

<%
Set otax = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->