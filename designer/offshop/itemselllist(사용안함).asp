<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2009.04.07 ������ ����
'			2010.06.08 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim page,shopid,jungsanid ,fromDate,toDate ,yyyymmdd1,yyymmdd2 ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 
dim i,totalsum ,datefg
	jungsanid = session("ssBctID")
	shopid = request("shopid")
	page = request("page")
	if page="" then page=1
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"
		
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
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

totalsum=0

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectDesigner = jungsanid
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.frectdatefg = datefg
	
	if shopid<>"" then
		ooffsell.GetDaylySellItemList
	end if
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		SHOP : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		������� :
		<% drawmaechul_datefg "datefg" ,datefg ,""%> <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>

<Br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		* �߰��ǸŸ� �ϴ� ����(��Ÿ�����)�� ��� ��������Ϸ� �˻��� �ϼž� ��Ȯ�� ������ ���� �˴ϴ�.
		<br>�Ǹų����� �� ������, �����ǸŸ���(����5�ð�), �ְ��ǸŸ���(���� 10�ð�) ������Ʈ �Ǹ�,
		<br>������ �ֹ��� �������� ���� �˴ϴ�.		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="86">���ڵ�</td>
	<td width="86">����<br>���ڵ�</td>
	<td width="90">�귣��</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td width="70">�Һ��ڰ�</td>
	<td width="70">�ǸŰ�</td>
	<td width="60">����</td>
	<td width="80">�ǸŰ��հ�</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<% totalsum = totalsum + ooffsell.FItemList(i).FSubTotal %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).GetBarCode %></td>
	<td><%= ooffsell.FItemList(i).fextbarcode %></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>
	<td align="left"><%= ooffsell.FItemList(i).FItemName %></td>
	<td><%= ooffsell.FItemList(i).FItemOptionName %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
	<td><%= ooffsell.FItemList(i).FItemNo %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSubTotal,0) %></td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td><b>�Ѱ�</b></td>
	<td colspan="10" align="right"><b><%= FormatNumber(totalsum,0) %></b></td>
</tr>
</table>

<% if shopid="" then %>
	<script language='javascript'>alert('���� ������ �ּ���');</script>
<% end if %>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->