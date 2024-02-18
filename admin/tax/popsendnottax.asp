<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ������ ��꼭 �������
' History : 2012.09.20 �ѿ�� ����
'###########################################################

Response.Expires=-1440
'Response.Buffer=true	
Response.ContentType = "application/vnd.ms-excel" 	
Response.AddHeader "Content-disposition","attachment;filename=TEN" & Left(CStr(now()),10) & "_�����۰�꼭.xls"
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->

<%
dim i

Dim otax
Set otax = new CEsero
	otax.getsendnottax()
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
if otax.FResultCount>0 then
	
For i = 0 To otax.FResultCount - 1

%>
<tr align="center" bgcolor="#FFFFFF">
	<td class='txt'><%= otax.FItemList(i).ftaxKey %></td>
	<td><%= otax.FItemList(i).fappDate %></td>
	<td><%= otax.FItemList(i).fsellCorpNo %></td>
	<td><%= otax.FItemList(i).fsellJongNo %></td>
	<td><%= otax.FItemList(i).fsellCorpName %></td>
	<td><%= otax.FItemList(i).fsellCeoName %></td>
	<td><%= otax.FItemList(i).fsellEmail %></td>
	<td><%= otax.FItemList(i).fbuyCorpNo %></td>
	<td><%= otax.FItemList(i).fBuyCorpName %></td>
	<td><%= otax.FItemList(i).fBuyCeoName %></td>
	<td><%= otax.FItemList(i).fbuyEmail %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).ftotSum,0) %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).fsuplySum,0) %></td>
	<td align="right"><%= FormatNumber(otax.FItemList(i).ftaxSum,0) %></td>
	<td><%= otax.FItemList(i).ftaxSellType %></td>
	<td><%= otax.FItemList(i).ftaxModiType %></td>
	<td><%= otax.FItemList(i).ftaxType %></td>
	<td><%= otax.FItemList(i).fBigo %></td>
	<td><%= otax.FItemList(i).fDtlName %></td>
	<td><%= otax.FItemList(i).fbizseccd %></td>
</tr>

<% Next %>

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