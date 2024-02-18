<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� (����Ÿ��Ʈ ��輭������ ������)
' History : 2010.05.10 ������ ����
'			2012.02.07 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , fromDate,toDate , shopid ,i , datelen, datelen2 ,makerid, menupos, page
dim datefg , tmpdate , maechultype ,totrealsellprice ,totitemno ,totprofit ,totsellprice, totsuplyprice
	makerid = requestCheckVar(request("makerid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page="" then page = 1
if datefg = "" then datefg = "maechul"
tmpdate = dateadd("m",-1,date)

if (yyyy1="") then yyyy1 = Cstr(Year(tmpdate))
if (mm1="") then mm1 = Cstr(Month(tmpdate))
if (dd1="") then dd1 = Cstr(day(tmpdate))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'C_IS_SHOP = TRUE
'C_IS_Maker_Upche = TRUE

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

dim oreport
set oreport = new COffShopSell
	oreport.FPageSize = 2000
	oreport.FCurrPage = page
	oreport.frectdatefg = datefg
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.FRectShopID = shopid
	oreport.frectmakerid = makerid

	'/����Ÿ��Ʈ
	oreport.GetNewBrandSell_item_datamart

	'/���ε�� �ǽð�
	'oreport.GetNewBrandSell_item

totrealsellprice = 0
totitemno =0
totprofit = 0
totsellprice = 0
totsuplyprice = 0
%>

<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="1" cellspacing="1" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			</td>
		</tr>
		</table>
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<Br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="3" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		�� �Ϸ� �������� �Ǹŵ� ���� ����̸�, �Ϸ翡 �ѹ� ������ ������Ʈ �˴ϴ�.
    </td>
    <td align="right">

    </td>
</tr>
</form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oreport.FResultCount %></b> �� �� 2000�Ǳ��� ���� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>�����</td>
	<td>��¥</td>
	<td>�����ڵ�</td>
	<td>��ǰ��<Br><font color="blue">[�ɼǸ�]</font></td>
	<td>�귣��</td>
	<td>�ǸŰ�</td>
	<td>�����</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>���԰�</td>
	<% end if %>

	<td>�Ǹ�<Br>����</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>����<Br>����</td>
	<% end if %>
</tr>
<%
if oreport.FResultCount > 0 then

for i=0 to oreport.FResultCount - 1

totsellprice = totsellprice + oreport.FItemList(i).fsellprice
totrealsellprice = totrealsellprice + oreport.FItemList(i).frealsellprice
totsuplyprice = totsuplyprice + oreport.FItemList(i).fsuplyprice
totitemno = totitemno + oreport.FItemList(i).fitemno
totprofit = totprofit + (oreport.FItemList(i).frealsellprice - oreport.FItemList(i).fsuplyprice)
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
	<td><%= oreport.FItemList(i).fshopname %><Br>(<%= oreport.FItemList(i).fshopid %>)</td>
	<td><%= oreport.FItemList(i).fIXyyyymmdd %></td>
	<td><%= oreport.FItemList(i).fitemgubun %><%= CHKIIF(oreport.FItemList(i).fitemid>=1000000,Format00(8,oreport.FItemList(i).fitemid),Format00(6,oreport.FItemList(i).fitemid)) %><%= oreport.FItemList(i).fitemoption %></td>
	<td>
		<%= oreport.FItemList(i).fitemname %>

		<% if oreport.FItemList(i).fitemoptionname <> "" then %>
			<BR><font color="blue">[<%= oreport.FItemList(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= oreport.FItemList(i).fmakerid %></td>
	<td align="right"><%= FormatNumber(oreport.FItemList(i).fsellprice,0) %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).frealsellprice,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(oreport.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<td><%= oreport.FItemList(i).fitemno %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<%= FormatNumber(oreport.FItemList(i).frealsellprice - oreport.FItemList(i).fsuplyprice,0) %>
		</td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=5>�հ�</td>
	<td align="right"><%= FormatNumber(totsellprice,0) %></td>
	<td align="right"><%= FormatNumber(totrealsellprice,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(totsuplyprice,0) %></td>
	<% end if %>

	<td><%= FormatNumber(totitemno,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(totprofit,0) %></td>
	<% end if %>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height=24>
	<td align="center" colspan=15>�˻� ����� �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->