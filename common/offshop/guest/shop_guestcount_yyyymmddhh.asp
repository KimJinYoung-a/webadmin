<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������湮ī��Ʈ
' History : 2012.05.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/guest/shop_guestcount_cls.asp"-->
<%
dim shopid , i ,yyyy1 ,mm1 ,dd1 ,yyyy2 ,mm2 ,dd2 ,page ,fromDate ,toDate
	shopid = request("shopid")	
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	page = request("page")

	if page = "" then page = 1

if yyyy1="" then
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

'/����
if (C_IS_SHOP) then
	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID		'"streetshop011"
	end if
end if

dim oguest
set oguest = new cguestcount_list
	oguest.FPageSize = 500
	oguest.FCurrPage = page
	oguest.FRectShopID = shopid	
	oguest.FRectStartDay = fromDate
	oguest.FRectEndDay = toDate
	oguest.fshopguestcount_yyyymmddhh

%>

<script language="javascript">

function frmsubmit(page){
	frm.page.value = page;
	frm.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>		
				���� : <% drawSelectBoxOffShopdiv_off "shopid",shopid,"1","","" %>
				��¥ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('');">
	</td>
</tr>	
</form>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>	
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=oguest.FresultCount%></b>&nbsp;&nbsp; �� �ִ� 500�� ���� ��ȸ����
	</td>
</tr>

<%
dim z1_in_sum ,z2_in_sum ,z1z2_in_sum ,tmpshopid
	z1_in_sum = 0
	z2_in_sum = 0
	z1z2_in_sum = 0

if oguest.FResultCount>0 then
	
For i = 0 To oguest.FResultCount - 1

	if tmpshopid <> oguest.FItemList(i).fshopid then
		if i <> 0 then
%>
			<tr align="center" bgcolor="#FFFFFF">
				<td colspan=2>���հ�</td>
				<td align="right"><% = FormatNumber(z1_in_sum,0) %></td>
				<td align="right"><% = FormatNumber(z2_in_sum,0) %></td>
				<td align="right"><% = FormatNumber(z1z2_in_sum,0) %></td>
				<td></td>
			</tr>
<% 
			z1_in_sum = 0
			z2_in_sum = 0
			z1z2_in_sum = 0
		end if
%>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����</td>
			<td>
				��¥
			</td>
			<td>
				<%= getzonegubun(oguest.FItemList(i).fshopid,"z1_in") %>
			</td>
			<td>
				<%= getzonegubun(oguest.FItemList(i).fshopid,"z2_in") %>
			</td>		
			<td>
				�հ�
			</td>
			<!--<td>
				z1_out
			</td>
			<td>
				z2_out
			</td>-->			
			<td>���</td>
		</tr>
<%				
	end if
	
	tmpshopid = oguest.FItemList(i).fshopid
	z1_in_sum = z1_in_sum + oguest.FItemList(i).fz1_in
	z2_in_sum = z2_in_sum + oguest.FItemList(i).fz2_in
	z1z2_in_sum = z1z2_in_sum + (oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in)

%>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= oguest.FItemList(i).fshopname %>
	</td>
	<td><%= oguest.FItemList(i).fyyyymmdd %></td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz1_in,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz2_in,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(oguest.FItemList(i).fz1_in + oguest.FItemList(i).fz2_in,0) %>
	</td>
	<!--<td align="right">
		<%'= FormatNumber(oguest.FItemList(i).fz1_out,0) %>
	</td>
	<td align="right">
		<%'= FormatNumber(oguest.FItemList(i).fz2_out,0) %>
	</td>-->
	<td>
	</td>
</tr>

<%
Next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=2>���հ�</td>
	<td align="right"><% = FormatNumber(z1_in_sum,0) %></td>
	<td align="right"><% = FormatNumber(z2_in_sum,0) %></td>
	<td align="right"><% = FormatNumber(z1z2_in_sum,0) %></td>
	<td></td>
</tr>

<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>

<%
set oguest= Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->