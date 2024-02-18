<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� ��������-�Ǹ�ó��
' History : 2013.01.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/statistic/statisticCls.asp" -->

<%
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,BanPum ,fromDate ,toDate, offgubun, reload
dim vTot_ordercnt, vTot_orgsellprice, vTot_totsale, vTot_possale, vTot_realSum, vTot_bonuscouponprice
dim vTot_TotalSum, vTot_Maechul, inc3pl
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
    inc3pl = request("inc3pl")
    
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

Set cStatistic = New cStatic_list
	cStatistic.FRectdatefg = datefg
	cStatistic.FRectStartdate = fromDate
	cStatistic.FRectEndDate = toDate
	cStatistic.FRectshopid = shopid
	cStatistic.FRectBanPum = BanPum
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectInc3pl = inc3pl	
	cStatistic.fStatistic_shop()

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
                
	if (diffDay > 1095){
		alert('3���������� �ǽð��˻��� �����մϴ�.');
		return;
	}
	
	frm.submit();
}

function detailStatistic(yyyy1,mm1,dd1,yyyy2,mm2,dd2,shopid,datefg, offgubun)
{
	var detailStatistic = window.open("/common/offshop/maechul/statistic/statistic_daily.asp?yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2+"&shopid="+shopid+"&datefg="+datefg+"&offgubun="+offgubun+"&inc3pl=<%=inc3pl%>&menupos=<%=menupos%>","detailStatistic","width=1000,height=780,scrollbars=yes,resizable=yes");
	detailStatistic.focus();
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
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
		�� �ǽð� �����ʹ� �ֱ� 3�� �����͸� �˻� �����մϴ�.
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
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">�Ǹ�ó(����)</td>
    <td align="center">�Ǽ�</td>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">�Һ��ڰ�</td>
    <td align="center">�ǸŰ�<br>A</td>
    <td align="center">�������αݾ�<br>A-B</td>
    <td align="center">�ǸŰ�(���ΰ�)<br>B</td>
    <td align="center">��ǰ����<br>����</td>
    <td align="center">�����Ѿ�<br>B-C</td>
    <td align="center">���ʽ���������<br>C</td>
    <td align="center">��Ÿ����</td>
    <% end if %>
    <td align="center">�����</td>
    <td align="center">���</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then
	
For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FItemList(i).fshopname %></td>
	<td align="center"><%= FormatNumber(cStatistic.FItemList(i).fordercnt,0) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fIorgsellprice,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FTotalSum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FTotalSum-cStatistic.FItemList(i).FMaechul,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaechul,0) %></td>
	<td align="right" style="padding-right:5px;"></td>
	<td align="right" style="padding-right:5px;" bgcolor="#9DCFFF"><%= FormatNumber(cStatistic.FItemList(i).FMaechul,0) %></td>
	<td align="right" style="padding-right:5px;"><%'= FormatNumber(cStatistic.FItemList(i).fbonuscouponprice,0) %></td>
	<td align="right" style="padding-right:5px;"></td>
    <% end if %>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= FormatNumber(cStatistic.FItemList(i).FMaechul,0) %></b></td>
	<td align="center" >
		[<a href="javascript:detailStatistic('<%=yyyy1%>','<%=mm1%>','<%=dd1%>','<%=yyyy2%>','<%=mm2%>','<%=dd2%>','<%= cStatistic.FItemList(i).fshopid %>','<%= datefg %>','<%= offgubun %>')">�Ϻ�</a>]
	</td>
</tr>
<%
vTot_ordercnt = vTot_ordercnt + cStatistic.FItemList(i).fordercnt
vTot_orgsellprice = vTot_orgsellprice + cStatistic.FItemList(i).fIorgsellprice
vTot_totsale = vTot_totsale + cStatistic.FItemList(i).FTotalSum
vTot_possale = vTot_possale + cStatistic.FItemList(i).FTotalSum-cStatistic.FItemList(i).FMaechul
vTot_realSum = vTot_realSum + cStatistic.FItemList(i).FMaechul
'vTot_bonuscouponprice = vTot_bonuscouponprice + cStatistic.FItemList(i).fbonuscouponprice
vTot_TotalSum = vTot_TotalSum + cStatistic.FItemList(i).FMaechul
vTot_Maechul = vTot_Maechul + cStatistic.FItemList(i).fMaechul

Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">�հ�</td>
	<td align="center"><%=FormatNumber(vTot_ordercnt,0)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_orgsellprice,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_totsale,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_possale,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_realSum,0)%></td>
	<td align="right" style="padding-right:5px;"></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_TotalSum,0)%></td>
	<td align="right" style="padding-right:5px;"><%'=FormatNumber(vTot_bonuscouponprice,0)%></td>	
	<td align="right" style="padding-right:5px;"></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_Maechul,0)%></td>
	<td align="center"></td>
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