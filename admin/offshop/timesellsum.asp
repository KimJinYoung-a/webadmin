<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
'		[OFF]����_������>>���Ϻ�����м� /admin/offshop/weeklysellreport.asp ������ ���
' History : 2009.04.07 ������ ����
'			2010.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/guest/shop_guestcount_cls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 , shopid ,fromDate,toDate , yyyymmdd1,yyymmdd2 ,i ,datefg
dim weekdate ,oldlist ,offgubun ,FmNum ,CurrencyUnit, CurrencyChar, ExchangeRate , makerid
dim buyergubun, inc3pl
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	shopid = requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),16)
	weekdate = requestCheckVar(request("weekdate"),30)
	oldlist = requestCheckVar(request("oldlist"),10)
	offgubun = requestCheckVar(request("offgubun"),16)
	makerid = requestCheckVar(request("makerid"),32)
	buyergubun = requestCheckVar(request("buyergubun"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"		
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

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
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

''��Ÿ�� ������ȸ ���� 
Dim isFixShopView
IF (session("ssBctID")="doota01") then 
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim oreport
set oreport = new COffShopSellReport
	oreport.FRectStartDay = fromDate
	oreport.FRectEndDay = toDate
	oreport.FRectShopID = shopid
	oreport.frectdatefg = datefg
	oreport.frectweekdate = weekdate
	oreport.FRectOldJumun = oldlist
	oreport.FRectOffgubun = offgubun
	oreport.FRectbuyergubun = buyergubun
	oreport.FRectInc3pl = inc3pl
	
	'//���忡 ���湮ī��Ʈ ������ ������
	if existsguestcountshopid(shopid) then
		oreport.getshopguestcountandhoursellreport
	else
		oreport.SearchMallSellrePort5
	end if

Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW" or CurrencyUnit="",0,2)
%>

<script type='text/javascript'>
	
function frmsubmit(){
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">		
		*�Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="oldlist" onclick='frmsubmit();' <% if oldlist="on" then response.write "checked" %>>3������
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ����:<% drawweekday_select "weekdate" , weekdate ," onchange='frmsubmit();'" %>
		&nbsp;&nbsp;
		* ���� ���� : <% Call DrawShopDivCombo("offgubun",offgubun) %>
		&nbsp;&nbsp;
		* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
        &nbsp;&nbsp;
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    	�� ������ �ֹ��� �������� ���� �˴ϴ�.
    </td>
    <td align="right">	       
    </td>
</tr>	
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oreport.FResultCount %></b>
	</td>
</tr>

<% if existsguestcountshopid(shopid) then %>
	<tr bgcolor="#EEEEEE" align="center">
		<td rowspan=2>�ð�</td>
		<td rowspan=2>�����<Br>(���ϸ�������)</td>
		<td rowspan=2>�ֹ�<br>�Ǽ�</td>
		<td rowspan=2>�ֹ��Ǽ�<br>���</td>
		<td colspan=3>������湮��</td>
	</tr>
	<tr bgcolor="#EEEEEE" align="center">
		<td><%= getzonegubun(shopid,"z1_in") %></td>
		<td><%= getzonegubun(shopid,"z2_in") %></td>
		<td>�հ�</td>
	</tr>
<% else %>
	<tr bgcolor="#EEEEEE" align="center">
		<td>�ð�</td>
		<td>�����<Br>(���ϸ�������)</td>
		<td>�ֹ�<br>�Ǽ�</td>
	</tr>
<% end if %>

<% if oreport.FResultCount > 0 then %>
<% for i=0 to oreport.FResultCount-1 %>
<tr bgcolor="#FFFFFF" align="center">
	<td>
		<%= oreport.FItemList(i).Fgpart %> ��
	</td>
	<td align="right" bgcolor="#E6B9B8">
		<%= FormatNumber(oreport.FItemList(i).Fselltotal+oreport.FItemList(i).fspendmile,FmNum) %>&nbsp;<%= CurrencyChar %>
	</td>
	<td>
		<%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %>
	</td>
	
	<% if existsguestcountshopid(shopid) then %>
		<td>
		<%
			If oreport.FItemList(i).fz1_all + oreport.FItemList(i).fz2_all = 0 Then
				Response.Write 0 & " %"
			Else
				Response.Write ROUND((oreport.FItemList(i).Fsellcnt/(oreport.FItemList(i).fz1_all + oreport.FItemList(i).fz2_all))*100,2) & " %"
			End If
		%>
		</td>
		<td>
			<%= FormatNumber(oreport.FItemList(i).fz1_all,0) %>
		</td>
		<td>
			<%= FormatNumber(oreport.FItemList(i).fz2_all,0) %>
		</td>
		<td>
			<%= FormatNumber(oreport.FItemList(i).fz1_all + oreport.FItemList(i).fz2_all,0) %>
		</td>
	<% end if %>	
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center"  >[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->