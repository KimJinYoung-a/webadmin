<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �ܱ��α������
' History : 2013.02.20 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/buyergubun_cls_off.asp"-->
<%
dim shopid ,oldlist ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2 ,fromDate,toDate ,datefg, buyergubun
dim i ,totrealsum, totcnt , totspendmile, totmaechul, olddatay ,offgubun , reload , parameter, page, inc3pl
	olddatay = RequestCheckVar(request("olddatay"),10)
	shopid = request("shopid")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	oldlist = request("oldlist")
	datefg = request("datefg")
	offgubun = request("offgubun")
	reload = request("reload")
	page = request("page")
	buyergubun = request("buyergubun")
    inc3pl = request("inc3pl")

if datefg = "" then datefg = "maechul"
if reload <> "on" and offgubun = "" then offgubun = "95"
if page = "" then page=1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now())))
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


dim obuyer
set obuyer = new cbuyerlist
	obuyer.FPageSize = 500
	obuyer.FCurrPage = page
	obuyer.FRectShopID = shopid
	obuyer.FRectStartDay = fromDate
	obuyer.FRectEndDay = toDate
	obuyer.FRectOldData = oldlist
	obuyer.frectdatefg = datefg
	obuyer.frectoffgubun = offgubun
	obuyer.frectbuyergubun = buyergubun
	obuyer.FRectInc3pl = inc3pl
	obuyer.getbuyergubun_list

totrealsum = 0
totcnt = 0
totspendmile = 0
totmaechul = 0

parameter = "oldlist="&oldlist&"&datefg="&datefg&"&offgubun="&offgubun&"&menupos="&menupos
%>

<script language='javascript'>

function frmsubmit(){

	frm.submit();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="reload" value="on">
<input type="hidden" name="page" value=1>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="olddatay" value="<%= olddatay %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<!--<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������������ȸ-->
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					* ���� : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
				<p>
				* ���屸��:<% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
	</td>
</tr>
</form>
</table>
<!-- ǥ ��ܹ� ��-->
<Br>
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
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= obuyer.FTotalCount %></b> �� ��500�� ������ �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>����</td>
	<td>����</td>
	<td>�ֹ�<Br>�Ǽ�</td>
	<td>�����</td>
	<td>���</td>
</tr>
<%
if obuyer.FResultCount > 0 then

for i=0 to obuyer.FResultCount -1

totcnt = totcnt + obuyer.FItemList(i).fcnt
totrealsum = totrealsum + obuyer.FItemList(i).frealsum
totspendmile = totspendmile + obuyer.FItemList(i).fspendmile
totmaechul = totmaechul + (obuyer.FItemList(i).frealsum+obuyer.FItemList(i).fspendmile)

%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
	<td><%= obuyer.FItemList(i).fshopname %> (<%= obuyer.FItemList(i).Fshopid %>)</td>
	<td><%= obuyer.FItemList(i).fcodename %></td>
	<td><%= FormatNumber(obuyer.FItemList(i).fcnt,0) %></td>
	<td bgcolor="#E6B9B8" align="right">
		<%= FormatNumber(obuyer.FItemList(i).frealsum+obuyer.FItemList(i).fspendmile,0) %>
	</td>
	<td>
	</td>
</tr>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=2>�հ�</td>
	<td><%= FormatNumber(totcnt,0) %></td>
	<td align="right"><%= FormatNumber(totmaechul,0) %></td>
	<td></td>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>

</table>

<%
set obuyer = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
