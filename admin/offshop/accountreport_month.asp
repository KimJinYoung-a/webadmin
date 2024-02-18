<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� ��������������� �ǽð�
' History : 2012.11.05 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->

<%
Dim i, cStatistic, shopid, datefg, yyyy1, mm1, stdate, BanPum, fromDate, toDate, offgubun, reload, inc3pl
dim vTot_spendmile, vTot_TenGiftCardPaySum, vTot_giftcardPaysum, vTot_cardsum, vTot_cashsum, vTot_TotalSum, vTot_extPaySum
	shopid 	= requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1	  = requestCheckVar(request("mm1"),2)
	BanPum     = requestCheckVar(request("BanPum"),1)
	offgubun = requestCheckVar(request("offgubun"),10)
	reload = requestCheckVar(request("reload"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"
if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

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
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

Set cStatistic = New COffShopSellReport
	cStatistic.FRectdatefg = datefg
	cStatistic.FRectStartdate = yyyy1 + "-" + mm1 + "-" + "01"
	cStatistic.FRectEnddate = CStr(DateAdd("m",1,DateSerial(yyyy1,mm1,1)))
	cStatistic.FRectshopid = shopid
	cStatistic.FRectBanPum = BanPum
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectInc3pl = inc3pl
	cStatistic.FPageSize = 500
	cStatistic.FCurrPage = 1
	cStatistic.GetJumunMethodReportMonth()
%>

<script language="javascript">

function searchSubmit(){
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
	<input type="hidden" name="reload" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
	<td width="30" bgcolor="<%= adminColor("gray") %>">�˻�<Br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* �Ⱓ :&nbsp;
				<% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawYMBox yyyy1,mm1 %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% else %>
						<!--* ���� : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				&nbsp;&nbsp;
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
				<Br>
				* ��ǰ���� :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
		</tr>
	</table>
</form>
<!-- �˻� �� -->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">
	<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=cStatistic.FTotalCount%></b> �� �� 500�Ǳ��� �˻��˴ϴ�.
	</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">�Ⱓ</td>
    <td align="center" colspan="3"></td>
    <td align="center" colspan="3">�ǰ�����</td>
    <td align="center" width="150" rowspan="2">�����հ�</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">���ϸ���</td>
    <td align="center">����Ʈī��</td>
    <td align="center">��ǰ��</td>
    <td align="center">�ſ�ī��</td>
    <td align="center">����</td>
	<td align="center">��Ÿ</td>
	</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
	<td align="center">
		<%= getweekendcolor(cStatistic.fitemlist(i).FRegdate) %>
	</td>
	<td align="center"><%= getweekend(cStatistic.fitemlist(i).FRegdate) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fspendmile,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaySum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcardsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcashsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fextPaysum,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(cStatistic.FItemList(i).fselltotal,0) %></td>
	</tr>
<%
	vTot_spendmile	= vTot_spendmile + CLng(cStatistic.FItemList(i).fspendmile)
	vTot_TenGiftCardPaySum		= vTot_TenGiftCardPaySum + CLng(cStatistic.FItemList(i).fTenGiftCardPaySum)
	vTot_giftcardPaysum		= vTot_giftcardPaysum + CLng(cStatistic.FItemList(i).fgiftcardPaysum)
	vTot_extPaysum		= vTot_extPaysum + CLng(cStatistic.FItemList(i).fextPaysum)
	vTot_cardsum		= vTot_cardsum + CLng(cStatistic.FItemList(i).fcardsum)
	vTot_cashsum			= vTot_cashsum + CLng(cStatistic.FItemList(i).fcashsum)
	vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FItemList(i).fselltotal)

Next
%>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">�հ�</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_spendmile,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaySum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cardsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cashsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_extPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TotalSum,0) %></td>
	</tr>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
