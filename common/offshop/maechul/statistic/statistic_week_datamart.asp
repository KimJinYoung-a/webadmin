<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� �ֺ�
' History : 2012.11.02 �ѿ�� ����
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
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2, offgubun, reload, oldlist
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice
dim vTot_Miletotalprice, vTot_MaechulCountSum, vTot_MaechulPriceSum ,fromDate ,toDate, inc3pl
dim xl
	shopid 	= request("shopid")
	datefg = request("datefg")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	offgubun = request("offgubun")
	reload = request("reload")
	oldlist = request("oldlist")
    inc3pl = request("inc3pl")
	xl 			= request("xl")

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"
if (yyyy1="") then yyyy1 = Cstr(Year(beforeWeeksunday(DateAdd("ww",-12,now()))))
if (mm1="") then mm1 = Cstr(Month(beforeWeeksunday(DateAdd("ww",-12,now()))))
if (dd1="") then dd1 = Cstr(day(beforeWeeksunday(DateAdd("ww",-12,now()))))
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
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectOldData = oldlist
	cStatistic.FRectInc3pl = inc3pl
	cStatistic.fStatistic_weeklist_datamart()

if (xl = "Y") then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=datamart_off_weekly_xl.xls"
else

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

function detailStatistic(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid, datefg, offgubun, oldlist)
{
	var detailStatistic = window.open("/common/offshop/maechul/statistic/statistic_daily_datamart.asp?yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2+"&shopid="+shopid+"&datefg="+datefg+"&offgubun="+offgubun+"&oldlist="+oldlist+"&inc3pl=<%=inc3pl%>&menupos=<%=menupos%>","detailStatistic","width=1024,height=780,scrollbars=yes,resizable=yes");
	detailStatistic.focus();
}

function popXL()
{
    frmXL.submit();
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
						* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='searchSubmit();'" %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='searchSubmit();'" %>
					<% else %>
						<!--* ���� : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				<br>
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
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

<p />

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#EEEEEE">
	<tr>
		<td align="left">
			�� �˻� �Ⱓ�� ��� �������ϴ�. �˻� ��ư�� ������, �ƹ� ������ ����δٰ�, �ٽ� �˻���ư�� Ŭ������ ������.
		</td>
		<td align="right">
			<input type="button" class="button" value="�����ޱ�" onClick="popXL()">
		</td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->

<p />

<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=cStatistic.FTotalCount%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">�Ⱓ</td>
    <td align="center" colspan="2">�����(+)</td>
    <td align="center" colspan="2">�����(-)</td>
    <td align="center" colspan="2">������հ�</td>
    <td align="center" width="150" rowspan="2">���ϸ���</td>
    <td align="center" width="150" rowspan="2">�����Ѿ�</td>
    <td align="center" width="50" rowspan="2">���</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
    <td align="center">�ֹ��Ǽ�</td>
    <td align="center">�ݾ�</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center">
		<%= getweekendcolor(cStatistic.fitemlist(i).FMinDate) %> ~ <%= getweekendcolor(cStatistic.fitemlist(i).FMaxDate) %>
	</td>
	<td align="center"><%= Year(cStatistic.fitemlist(i).FMinDate) %> - <%= cStatistic.fitemlist(i).FWeek %>��</td>
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FCountPlus,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.fitemlist(i).FMaechulPlus,0) %></td>
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FCountMinus,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.fitemlist(i).FMaechulMinus,0) %></td>
	<td align="center"><%= FormatNumber(CLng(cStatistic.fitemlist(i).FCountPlus)+CLng(cStatistic.fitemlist(i).FCountMinus),0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8">
		<b><%= FormatNumber(CLng(cStatistic.fitemlist(i).FMaechulPlus)+CLng(cStatistic.fitemlist(i).FMaechulMinus),0) %></b>
	</td>
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FMiletotalprice,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.fitemlist(i).FSubtotalprice,0) %></td>
	<td align="center" >
		[<a href="javascript:detailStatistic('<%=Year(cStatistic.fitemlist(i).FMinDate)%>','<%=Month(cStatistic.fitemlist(i).FMinDate)%>','<%=Day(cStatistic.fitemlist(i).FMinDate)%>','<%=Year(cStatistic.fitemlist(i).FMaxDate)%>','<%=Month(cStatistic.fitemlist(i).FMaxDate)%>','<%=Day(cStatistic.fitemlist(i).FMaxDate)%>','<%= shopid %>','<%= datefg %>','<%= offgubun %>','<%= oldlist %>')">
		��</a>]
	</td>
</tr>
<%
vTot_CountPlus			= vTot_CountPlus + CLng(cStatistic.fitemlist(i).FCountPlus)
vTot_MaechulPlus		= vTot_MaechulPlus + CLng(cStatistic.fitemlist(i).FMaechulPlus)
vTot_CountMinus			= vTot_CountMinus + CLng(cStatistic.fitemlist(i).FCountMinus)
vTot_MaechulMinus		= vTot_MaechulMinus + CLng(cStatistic.fitemlist(i).FMaechulMinus)
vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(CLng(cStatistic.fitemlist(i).FCountPlus)+CLng(cStatistic.fitemlist(i).FCountMinus))
vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(CLng(cStatistic.fitemlist(i).FMaechulPlus)+CLng(cStatistic.fitemlist(i).FMaechulMinus))
vTot_Miletotalprice		= vTot_Miletotalprice + CLng(cStatistic.fitemlist(i).FMiletotalprice)
vTot_Subtotalprice		= vTot_Subtotalprice + CLng(cStatistic.fitemlist(i).FSubtotalprice)

Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">�հ�</td>
	<td align="center"><%=FormatNumber(vTot_CountPlus,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulPlus,0)%></td>
	<td align="center"><%=FormatNumber(vTot_CountMinus,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulMinus,0)%></td>
	<td align="center"><%=FormatNumber(vTot_MaechulCountSum,0)%></td>
	<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulPriceSum,0)%></b></td>
	<td align="center"><%=FormatNumber(vTot_Miletotalprice,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_Subtotalprice,0)%></td>
	<td></td>
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

<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="datefg" value="<%= datefg %>">
	<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
	<input type="hidden" name="mm1" value="<%= mm1 %>">
	<input type="hidden" name="dd1" value="<%= dd1 %>">
	<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
	<input type="hidden" name="mm2" value="<%= mm2 %>">
	<input type="hidden" name="dd2" value="<%= dd2 %>">
	<input type="hidden" name="offgubun" value="<%= offgubun %>">
	<input type="hidden" name="reload" value="<%= reload %>">
	<input type="hidden" name="oldlist" value="<%= oldlist %>">
	<input type="hidden" name="inc3pl" value="<%= inc3pl %>">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->