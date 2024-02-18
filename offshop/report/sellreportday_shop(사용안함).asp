<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2009.04.07 ������ ����
'			2010.03.26 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim page,shopid, yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, fromDate,toDate, oldlist
dim datefg , i
	shopid = session("ssBctID")
	page = request("page")
	if page="" then page=1
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	oldlist = request("oldlist")
	datefg = request("datefg")
	if datefg = "" then datefg = "maechul"	


if (shopid="doota01") then shopid="streetshop014"

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

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FPageSize=20
	ooffsell.FCurrPage=page
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOnlyShop = "on"
	ooffsell.FRectOldData = oldlist
	ooffsell.frectdatefg = datefg	
	ooffsell.GetDaylySumList
	
Dim CurrencyUnit, CurrencyChar, ExchangeRate
Dim FmNum
Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON",0,2)
	
%>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >1����������
				&nbsp;
				<% drawmaechuldatefg "datefg" ,datefg ,""%> 
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>	
</form>
</table>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
    <tr valign="bottom">       
        <td align="left">   
        <% if shopid="streetshop014" or  shopid="cafe002" then %>
        * ��������� : ��Ÿ(���� 5�� ���� ���� ���� 5�� ����), ��ȭ��(���� 3�� ���� ���� ���� 3�ñ���)<br>
        * �ֹ��� : 0�� ���� ~ 0�ñ���
        <% end if %>
	    </td>
	          
	</tr>	
</table>
<!-- ǥ �߰��� ��-->
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<% if ooffsell.FresultCount>0 then %>
<tr bgcolor="#EEEEEE" align="center">
	<td width="80">����ID</td>
	<td width="80"><%= CHKIIF(datefg="maechul","���������","�ֹ���") %></td>
	<td width="80">����Ǽ�</td>
	<td width="80">�ѱݾ�</td>
	<td width="80">�� ������</td>
	<td width="80">���ϸ������</td>
	<td width="80">���ϸ�������</td>
	<td width="60">�����۸��</td>
	<td width="60">�ֹ������</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).FShopid %></td>
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="center"><%= ooffsell.FItemList(i).FCount %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).FSpendMile,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSum,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSpendMile,FmNum) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FGainMile,FmNum) %></td>
	<td align="center"><a href="todayselldetail.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&datefg=<%= datefg %>&oldlist=<%= oldlist %>">����</a></td>
	<td align="center"><a href="todaysellmaster.asp?menupos=<%= menupos %>&terms=<%= ooffsell.FItemList(i).FTerm %>&shopid=<%= ooffsell.FItemList(i).FShopid %>&datefg=<%= datefg %>&oldlist=<%= oldlist %>">����</a></td>
</tr>
<% next %>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="11">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>
<%
set ooffsell= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->