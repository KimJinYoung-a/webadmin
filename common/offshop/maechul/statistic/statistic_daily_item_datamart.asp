<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ������� ��ǰ������-�Ϻ�
' History : 2013.01.29 �ѿ�� ����
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
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,BanPum ,fromDate ,toDate, vPurchaseType, makerid
dim vTot_itemno, vTot_Iorgsellprice, vTot_sellprice, vTot_realsellprice, vTot_suplyprice, vTot_MaechulProfit, vTot_MaechulProfitPer
dim offgubun, reload, oldlist, inc3pl, chkShowGubun
	shopid 	= request("shopid")
	datefg = request("datefg")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	BanPum     = request("BanPum")
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	makerid = request("makerid")
	offgubun = request("offgubun")
	reload = request("reload")
	oldlist = request("oldlist")
    inc3pl = request("inc3pl")
	chkShowGubun = request("chkShowGubun")

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

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
	cStatistic.FRectPurchasetype = vPurchasetype
	cStatistic.frectmakerid = makerid
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectOldData = oldlist
	cStatistic.FRectInc3pl = inc3pl
	cStatistic.FRectChkShowGubun = chkShowGubun
	cStatistic.fStatistic_daily_item_datamart()

vTot_itemno = 0
vTot_Iorgsellprice = 0
vTot_sellprice     = 0
vTot_realsellprice = 0
vTot_suplyprice = 0
vTot_MaechulProfit = 0
vTot_MaechulProfitPer = 0
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

function popitemdetail(yyyy1,mm1,dd1,shopid, purchasetype, makerid, datefg, offgubun, oldlist, commCd){
	var popitemdetail = window.open('/admin/offshop/todayselldetail.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid+'&purchasetype='+purchasetype+'&makerid='+makerid+'&datefg='+datefg+'&offgubun='+offgubun+'&oldlist='+oldlist+'&inc3pl=<%=inc3pl%>&commCd='+commCd+'&menupos=<%= menupos %>','popitemdetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popitemdetail.focus();
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
				* �Ⱓ :
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
						* ���� : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>
					<% end if %>
				<% end if %>
				<p>
				* ���� ���� : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
				&nbsp;&nbsp;
				* ��ǰ���� :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
				&nbsp;&nbsp;
				* �������� : 
				<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
				&nbsp;&nbsp;
				* ��ǰ���� :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
				&nbsp;&nbsp;
				<% if (C_IS_Maker_Upche) then %>
					* �귣�� : <%= makerid %><br>
					<input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				<p>
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;
				<input type="checkbox" name="chkShowGubun" value="Y" <% if (chkShowGubun = "Y") then %>checked<% end if %> > ���Ա��� ǥ��
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
<tr bgcolor="<%= adminColor("tabletop") %>">
	<% if (chkShowGubun = "Y") then %>
	<td align="center">���Ա���</td>
	<% end if %>
	<td align="center" colspan="2">�Ⱓ</td>
    <td align="center">�Ǹż���</td>
    <% if (NOT C_InspectorUser) then %>
    <td align="center">�Һ��ڰ�[��ǰ]</td>
    <td align="center">�ǸŰ�</td>
     <% end if %>
    <td align="center">�����</td>

    <% if not(C_IS_SHOP) then %>
    	<td align="center">�����Ѿ�[��ǰ]</td>
    <% end if %>

    <% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
    	<td align="center"><b>�������</b></td>
    	<td align="center">������</td>
    <% end if %>

    <td align="center">���</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<% if (chkShowGubun = "Y") then %>
	<td align="center"><%= GetJungsanGubunName(cStatistic.FItemList(i).Fjcomm_cd) %></td>
	<% end if %>
	<td align="center">
		<%= getweekendcolor(cStatistic.fitemlist(i).FRegdate) %>
	</td>
	<td align="center"><%= getweekend(cStatistic.fitemlist(i).FRegdate) %></td>
	<td align="center"><%= FormatNumber(cStatistic.FItemList(i).fitemno,0) %></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fIorgsellprice,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fsellprice,0) %></td>
	<% end if %>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(cStatistic.FItemList(i).frealsellprice,0) %></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td align="right" style="padding-right:5px;"><b><%= FormatNumber(cStatistic.FItemList(i).FMaechulProfit,0) %></b></td>
		<td align="right" style="padding-right:5px;"><%= cStatistic.FItemList(i).FMaechulProfitPer %>%</td>
	<% end if %>

	<td align="center" >
		[<a href="javascript:popitemdetail('<%= left(cStatistic.FItemList(i).FRegdate,4) %>','<%= mid(cStatistic.FItemList(i).FRegdate,6,2) %>','<%= right(cStatistic.FItemList(i).FRegdate,2) %>','<%= Shopid %>','<%= vPurchaseType %>','<%= makerid %>','<%= datefg %>','<%= offgubun %>','<%= oldlist %>','<%=cStatistic.FItemList(i).Fjcomm_cd%>');">��</a>]
	</td>
</tr>
<%
vTot_itemno = vTot_itemno + cStatistic.FItemList(i).fitemno
vTot_Iorgsellprice = vTot_Iorgsellprice + cStatistic.FItemList(i).fIorgsellprice
vTot_sellprice     = vTot_sellprice + cStatistic.FItemList(i).fsellprice
vTot_realsellprice = vTot_realsellprice + cStatistic.FItemList(i).frealsellprice
vTot_suplyprice = vTot_suplyprice + cStatistic.FItemList(i).fsuplyprice
vTot_MaechulProfit = vTot_MaechulProfit + cStatistic.FItemList(i).FMaechulProfit

Next

vTot_MaechulProfitPer = Round(((vTot_realsellprice - vTot_suplyprice)/CHKIIF(vTot_realsellprice=0,1,vTot_realsellprice))*100,2)
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (chkShowGubun = "Y") then %>
	<td align="center"></td>
	<% end if %>
	<td align="center" colspan="2">�Ѱ�</td>
	<td align="center"><%=FormatNumber(vTot_itemno,0)%></td>
	<% if (NOT C_InspectorUser) then %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_Iorgsellprice,0)%></td>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_sellprice,0)%></td>
    <% end if %>
	<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_realsellprice,0)%></td>

	<% if not(C_IS_SHOP) then %>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_suplyprice,0)%></td>
	<% end if %>

	<% if not(C_IS_SHOP) and not(C_IS_Maker_Upche) then %>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulProfit,0)%></b></td>
		<td align="right" style="padding-right:5px;"><%=vTot_MaechulProfitPer%>%</td>
	<% end if %>

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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
