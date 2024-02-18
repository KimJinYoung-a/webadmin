<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� ��������-������ĺ�
' History : 2012.11.05 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/statistic/statisticCls_datamart.asp" -->

<%
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,BanPum ,fromDate ,toDate, offgubun, reload
dim vTot_spendmile, vTot_TenGiftCardPaySum, vTot_giftcardPaysum, vTot_cardsum, vTot_cashsum, vTot_TotalSum
dim oldlist, vTot_spendmilecnt, vTot_TenGiftCardPaycount, vTot_giftcardPaycnt, vTot_cardcnt, vTot_cashcnt
dim inc3pl
dim onlyTenShop
	shopid 	= request("shopid")
	datefg = request("datefg")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	BanPum     = request("BanPum")
	offgubun = request("offgubun")
	reload = request("reload")
	oldlist = request("oldlist")
    inc3pl = request("inc3pl")
	onlyTenShop = request("onlyTenShop")

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
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
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

Set cStatistic = New cStaticdatamart_list
	cStatistic.FRectdatefg = datefg
	cStatistic.FRectStartdate = fromDate
	cStatistic.FRectEndDate = toDate
	cStatistic.FRectshopid = shopid
	cStatistic.FRectBanPum = BanPum
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectOldData = oldlist
	cStatistic.FRectInc3pl = inc3pl

	cStatistic.FRectOnlyTenShop = onlyTenShop

	cStatistic.fStatistic_checkmethod_datamart()

vTot_spendmile=0
vTot_TenGiftCardPaySum=0
vTot_giftcardPaysum=0
vTot_cardsum=0
vTot_cashsum=0
vTot_TotalSum=0
vTot_spendmilecnt=0
vTot_TenGiftCardPaycount=0
vTot_giftcardPaycnt=0
vTot_cardcnt=0
vTot_cashcnt=0
%>

<script language="javascript">

function searchSubmit()
{
	//��¥ ��
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

	if (diffDay > 1095 && frm.oldlist.checked == false){
		alert('3�� ���� �����ʹ� 3������������ȸ �� üũ�ϼž� �մϴ�');
		return;
	}

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
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������������ȸ
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", " onchange='searchSubmit();'" %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", " onchange='searchSubmit();'" %>
					<% else %>
						<!--* ���� : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				<br>
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
				&nbsp;&nbsp;
				* ��ǰ���� :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;
				<input type="checkbox" name="onlyTenShop" value="Y" <% if (onlyTenShop = "Y") then %>checked<% end if %> >
				�ٹ����� ���常(streetshop011, streetshop014, streetshop018)
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
		�� �˻� �Ⱓ�� ��� �������ϴ�. �˻� ��ư�� ������, �ƹ� ������ ����δٰ�, �ٽ� �˻���ư�� Ŭ������ ������.
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=cStatistic.FTotalCount%></b>
	</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td rowspan="3" colspan="2">�Ⱓ</td>
    <td colspan="6"></td>
    <td colspan="4">�ǰ�����</td>
    <td width="150" rowspan="3">�����հ�</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td colspan="2">���ϸ���</td>
    <td colspan="2">����Ʈī��</td>
    <td colspan="2">��ǰ��</td>
    <td colspan="2">�ſ�ī��</td>
    <td colspan="2">����</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>�ݾ�</td>
    <td>�Ǽ�</td>
    <td>�ݾ�</td>
    <td>�Ǽ�</td>
    <td>�ݾ�</td>
    <td>�Ǽ�</td>
    <td>�ݾ�</td>
    <td>�Ǽ�</td>
    <td>�ݾ�</td>
    <td>�Ǽ�</td>
	</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF" align="center">
	<td>
		<%= getweekendcolor(cStatistic.fitemlist(i).FRegdate) %>
	</td>
	<td align="center"><%= getweekend(cStatistic.fitemlist(i).FRegdate) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fspendmile,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fspendmilecnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaySum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaycount,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaysum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaycnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcardsum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fcardcnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcashsum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fcashcnt,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(cStatistic.FItemList(i).fselltotal,0) %></td>
	</tr>
<%
	vTot_spendmile	= vTot_spendmile + CLng(cStatistic.FItemList(i).fspendmile)
	vTot_TenGiftCardPaySum		= vTot_TenGiftCardPaySum + CLng(cStatistic.FItemList(i).fTenGiftCardPaySum)
	vTot_giftcardPaysum		= vTot_giftcardPaysum + CLng(cStatistic.FItemList(i).fgiftcardPaysum)
	vTot_cardsum		= vTot_cardsum + CLng(cStatistic.FItemList(i).fcardsum)
	vTot_cashsum			= vTot_cashsum + CLng(cStatistic.FItemList(i).fcashsum)
	vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FItemList(i).fselltotal)
	vTot_spendmilecnt		= vTot_spendmilecnt + CLng(cStatistic.FItemList(i).fspendmilecnt)
	vTot_TenGiftCardPaycount		= vTot_TenGiftCardPaycount + CLng(cStatistic.FItemList(i).fTenGiftCardPaycount)
	vTot_giftcardPaycnt		= vTot_giftcardPaycnt + CLng(cStatistic.FItemList(i).fgiftcardPaycnt)
	vTot_cardcnt		= vTot_cardcnt + CLng(cStatistic.FItemList(i).fcardcnt)
	vTot_cashcnt		= vTot_cashcnt + CLng(cStatistic.FItemList(i).fcashcnt)

Next
%>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan="2">�հ�</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_spendmile,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_spendmilecnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaySum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaycount,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaysum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaycnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cardsum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_cardcnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cashsum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_cashcnt,0) %></td>
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
<!-- #include virtual="/lib/db/db3close.asp" -->
