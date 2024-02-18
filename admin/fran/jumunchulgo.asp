<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim shopid
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate

shopid = request("shopid")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))-1, Cstr(1))
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(1))

	yyymmdd2 = Left(dateadd("d", -1, toDate), 10)

        yyyy1 = left(fromDate,4)
        mm1 = Mid(fromDate,6,2)
        dd1 = Mid(fromDate,9,2)

        yyyy2 = left(yyymmdd2,4)
        mm2 = Mid(yyymmdd2,6,2)
        dd2 = Mid(yyymmdd2,9,2)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	yyymmdd2 = DateSerial(yyyy2, mm2, dd2)
	toDate = Left(dateadd("d", +1, yyymmdd2), 10)
end if


dim osheet
set osheet = new COrderSheet
osheet.FRectBaljuid = shopid
osheet.FRectStartDate = fromDate
osheet.FRectEndDate = toDate

osheet.GetFranBaljuVSChulgo


dim osheetitem
set osheetitem = new COrderSheet
osheetitem.FRectBaljuid = shopid
osheetitem.FRectStartDate = fromDate
osheetitem.FRectEndDate = toDate

osheetitem.GetFranBaljuVSChulgoByItem


dim i, tmp

%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>������ �ֹ���� ��� ���</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>���ֹ��� ���� ������ ����Դϴ�.(���ް�����)
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	�� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
	        	�Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="document.frm.submit();"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="90">�����̵�</td>
    	<td>���̸�</td>
      	<td width="50">����</td>
      	<td width="80">�ֹ��ݾ�</td>
      	<td width="80">���ݾ�</td>
      	<td width="80">(��ǰ�ݾ�)</td>
      	<td width="50">����(%)</td>
      	<td width="70">�ֹ�����</td>
      	<td width="70">������</td>
      	<td width="70">(��ǰ����)</td>
      	<td width="50">����(%)</td>
    </tr>
<% for i=0 to osheet.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="center"><%= osheet.FItemList(i).Fbaljuid %></td>
    	<td align="left"><%= osheet.FItemList(i).Fbaljuname %></td>
    	<td><%= osheet.FItemList(i).Fitemgubun %></td>
    	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
    	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
    	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalreturnsuplycash,0) %></td>
        <%
        if (osheet.FItemList(i).Fjumunsuplycash = 0) then
                tmp = -1
        else
                tmp = CInt(100*osheet.FItemList(i).Ftotalsuplycash/osheet.FItemList(i).Fjumunsuplycash)
        end if
        %>
    	<td align="right">
    	<% if (tmp < 90) then %>
    	  <font color=red><b><%= tmp %></b></font>
    	<% else %>
    	  <%= tmp %>
    	<% end if %>
    	</td>
    	<td align="right"><%= osheet.FItemList(i).Fjumunitemno %></td>
    	<td align="right"><%= osheet.FItemList(i).Ftotalitemno %></td>
    	<td align="right"><%= osheet.FItemList(i).Ftotalreturnitemno %></td>
        <%
        if (osheet.FItemList(i).Fjumunitemno = 0) then
                tmp = -1
        else
                tmp = CInt(100*osheet.FItemList(i).Ftotalitemno/osheet.FItemList(i).Fjumunitemno)
        end if
        %>
    	<td align="right">
    	<% if (tmp < 90) then %>
    	  <font color=red><b><%= tmp %></b></font>
    	<% else %>
    	  <%= tmp %>
    	<% end if %>
    	</td>
    </tr>
<% next %>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        <br>
        * �ֹ� ��� ���̰� ������ǰ(���ް�����, �ְ����� ��ǰ 10���� ǥ��)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="90">�����̵�</td>
      	<td width="30">����</td>
      	<td width="50">��ǰID</td>
      	<td>��ǰ��</td>
      	<td>�ɼǸ�</td>
      	<td width="80">�ֹ��ݾ�</td>
      	<td width="80">���ݾ�</td>
      	<td width="50">����(%)</td>
    </tr>
<% for i=0 to osheetitem.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="center"><%= osheetitem.FItemList(i).Fbaljuid %></td>
    	<td><%= osheetitem.FItemList(i).FItemGubun %></td>
    	<td><%= osheetitem.FItemList(i).FItemId %></td>
    	<td align="left"><%= osheetitem.FItemList(i).FItemName %></td>
    	<td align="left"><%= osheetitem.FItemList(i).FItemOptionname %></td>
    	<td align="right"><%= FormatNumber(osheetitem.FItemList(i).Fjumunsuplycash,0) %></td>
    	<td align="right"><%= FormatNumber(osheetitem.FItemList(i).Ftotalsuplycash,0) %></td>
        <%
        if (osheetitem.FItemList(i).Fjumunsuplycash = 0) then
                tmp = -1
        else
                tmp = CInt(100*osheetitem.FItemList(i).Ftotalsuplycash/osheetitem.FItemList(i).Fjumunsuplycash)
        end if
        %>
    	<td align="right">
    	<% if (tmp < 90) then %>
    	  <font color=red><b><%= tmp %></b></font>
    	<% else %>
    	  <%= tmp %>
    	<% end if %>
    	</td>
    </tr>
<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<p>
<%

set osheet = Nothing
set osheetitem = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->