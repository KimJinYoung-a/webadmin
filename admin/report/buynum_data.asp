<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->

�������
<%

dbget.close()	:	response.End


dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, nextdateStr
Dim fromDate,toDate,oreport
dim buynum,reyyyy

buynum = request("buynum")
reyyyy = request("reyyyy")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())

if (buynum="") then buynum = "1"
if (reyyyy="") then reyyyy = "2003"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

set oreport = new CBuyNumReport
oreport.FRectRegStart = fromDate
oreport.FRectRegEnd = toDate
oreport.FRectBuyNum = buynum
oreport.FRectYYYY = reyyyy
oreport.FirstBuySellReport


dim buysellavg,buycntavg

if IsNull(oreport.Fsubtotalprice) then oreport.Fsubtotalprice=0
if IsNull(oreport.Fitemno) then oreport.Fitemno=0

if IsNull(oreport.Fsubtotalprice) or (oreport.Fitemno=0) then
	buysellavg = 0
else
	buysellavg = CLng(oreport.Fsubtotalprice / oreport.Fitemno)
end if

	buycntavg = Round((oreport.Ftotalcnt / oreport.Fcnt),2)

dim i
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		�Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;���� Ƚ�� : <input type="text" name="buynum" value="<% = buynum %>" size="4">
		&nbsp;�Ⱓ :
		<select name="reyyyy">
			<option value="2003" <% if reyyyy = "2003" then response.write "selected" %>>2003</option>
			<option value="2004" <% if reyyyy = "2004" then response.write "selected" %>>2004</option>
			<option value="2005" <% if reyyyy = "2005" then response.write "selected" %>>2005</option>
			<option value="2006" <% if reyyyy = "2006" then response.write "selected" %>>2006</option>
		</select>
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<br>
<div class="a">ȸ�� ���� ����</div>
<table width="800" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td align="center" height="25">���</td>
	<td align="center">�� ��</td>
	<td align="center">�� ��</td>
	<td align="center">���</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25"><% = buynum %>��° ����</td>
		<td align="center"><%= FormatNumber(CLng(oreport.Fitemno),0) %>��</td>
		<td align="center"><%= FormatNumber(oreport.Fsubtotalprice,0) %>��</td>
		<td align="center"><% = FormatNumber(buysellavg,0) %>��</td>
</tr>
<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25"><% = reyyyy %>�⵵</td>
		<td align="center"><%= FormatNumber(CLng(oreport.Fcnt),0) %>��</td>
		<td align="center"><%= FormatNumber(oreport.Ftotalcnt,0) %>��</td>
		<td align="center">���<% = buycntavg %>��</td>
</tr>
</table><br><br>
<table class="a" >
<tr>
	<td>
* ����Ƚ������հ��ܰ� : ����Ƚ���� ���ڸ� ������ ���� ��ȸ<br>
		�Ǽ�(�ѱ���Ƚ��), �Ѿ�(�ѱ��ž�), ���(��հ��ܰ�)
	</td>
</tr>
<tr>
	<td>
**	�Ⱓ ��ձ���Ƚ�� : �Ⱓ�� ���� �������� ��ȸ<br>
		�Ǽ�(�������Ѽ�), �Ѿ�(�ѱ���Ƚ��), ���(��ձ���Ƚ��)
	</td>
</tr>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->