<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1, yyyy2, mm2, stplace, targetGbn, itemgubun
dim ipgoMWdiv, itemMWdiv, itemid, itemoption
dim startYYYYMMDD, endYYYYMMDD
dim addInfoType
dim lastmwdiv, lastmakerid
dim tmpDate


page       	= requestCheckvar(request("page"),10)
research	= requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
yyyy2       = requestCheckvar(request("yyyy2"),10)
mm2         = requestCheckvar(request("mm2"),10)
stplace     = requestCheckvar(request("stplace"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
itemid   	= requestCheckvar(request("itemid"),10)
itemoption  = requestCheckvar(request("itemoption"),10)
lastmwdiv	= requestCheckvar(request("lastmwdiv"),10)
lastmakerid	= requestCheckvar(request("lastmakerid"),32)


page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", -1, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
	yyyy2 = yyyy1
	mm2 = mm1

	yyyy1 = "2014"
	mm1 = "01"
end if

if (itemgubun = "") then
	itemgubun = "10"
end if


'// ============================================================================
dim ojaego
set ojaego = new CMonthlyStock

ojaego.FPageSize = 100
ojaego.FCurrPage = page
ojaego.FRectStartYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectEndYYYYMM = yyyy2 + "-" + mm2
ojaego.FRectPlaceGubun = stplace

ojaego.FRectItemGubun = itemgubun
ojaego.FRectItemid = itemid
ojaego.FRectItemOption = itemoption

ojaego.FRectMwDiv = lastmwdiv

if (itemid <> "") then
	ojaego.GetMonthlyAvgPriceLogics
end if
''startYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
''endYYYYMMDD = Left(DateAdd("d", -1, DateSerial(yyyy1, mm1 + 1, 1)), 10)

%>

<script language='javascript'>

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">��/�� :</font> <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %> �� ��ո��԰�
			&nbsp;
			<font color="#CC3333">�԰�ó:</font>
		    <select name="stplace" class="select">
        		<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >����
        	</select>
			&nbsp;
	    	<font color="#CC3333">���Ա���(����ڻ�):</font>
	        <select name="lastmwdiv" class="select">
				<option value="" <%= CHKIIF(lastmwdiv="","selected" ,"") %> >��ü</option>
				<option value="M" <%= CHKIIF(lastmwdiv="M","selected" ,"") %> >����</option>
				<option value="W" <%= CHKIIF(lastmwdiv="W","selected" ,"") %> >��Ź</option>
				<option value="X" <%= CHKIIF(lastmwdiv="X","selected" ,"") %> >��Ÿ</option>
        	</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
	    	<font color="#CC3333">��ǰ����:</font>
        	<select name="itemgubun" class="select">
				<option value="" <%= CHKIIF(itemgubun="","selected" ,"") %> >��ü
				<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >�Ϲ�(10)
				<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >����ǰ(70)
				<option value="85" <%= CHKIIF(itemgubun="85","selected" ,"") %> >����ǰ(85)
				<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >����ǰ(80)
				<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >��������(90)
        	</select>
			&nbsp;
			<font color="#CC3333">��ǰ�ڵ�:</font>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8">
			&nbsp;
			<font color="#CC3333">�ɼ�:</font>
			<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>

	<h5>�۾���...</h5>
	* �����԰� ��ǰ�� ���� �� ������ �ý������ �ջ��Ͽ� ����մϴ�.(���Ա����� ������ ���)<br>
	* ���Ա����� �ٸ��ų� �����԰� ��ǰ�� �ƴ� ��� ���庰�� ��ո��԰��� ���˴ϴ�.<br>

<p>

<% if (itemid = "") then %>
	<h5>���� ��ǰ�ڵ带 �����ϼ���.</h5>
<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60" rowspan="2">����</td>
		<td width="40" rowspan="2">�԰�ó</td>
		<td width="100" rowspan="2">����</td>
		<td width="30" rowspan="2">����</td>
		<td width=70 rowspan="2">��ǰ�ڵ�</td>
		<td width=40 rowspan="2">�ɼ�</td>

		<td colspan="2">�������(����)</td>

		<td colspan="2">�������(����)</td>

		<td colspan="2">����԰�(����)</td>

		<td colspan="2">��ո��԰�(����)</td>

		<td width=60 rowspan="2">���Ա���<br>(����)</td>
		<td width=120 rowspan="2">�귣��</td>

		<td rowspan="2">���</td>
	</tr>

    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width=40>����</td>
		<td width=60>�ݾ�</td>
		<td width=40>����</td>
		<td width=60>�ݾ�</td>
		<td width=40>����</td>
		<td width=60>�ݾ�</td>
		<td width=60>����</td>
		<td width=60>���</td>
	</tr>

	<% if ojaego.FResultCount >0 then %>
	<% for i=0 to ojaego.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaego.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaego.FItemList(i).GetStockPlaceName %></td>
		<td align=center><%= ojaego.FItemList(i).Fshopid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemgubun %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemoption %></td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockPrev, 0) %></td>
		<td align=right>
			<% if Not IsNull(ojaego.FItemList(i).FavgipgoPriceSumPrev) then %>
			<%= FormatNumber(ojaego.FItemList(i).FavgipgoPriceSumPrev, 0) %>
			<% end if %>
		</td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockShopPrev, 0) %></td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockBuySumShopPrev, 0) %></td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotItemNo, 0) %></td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotBuyCash, 0) %></td>

		<td align=right>
			<% if Not IsNull(ojaego.FItemList(i).FavgipgoPricePrev) then %>
			<%= FormatNumber(ojaego.FItemList(i).FavgipgoPricePrev, 0) %>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FavgipgoPrice, 0) %></td>

		<td align=center>
			<% if Not IsNull(ojaego.FItemList(i).FlastmwdivPrev) then %>
				<% if (ojaego.FItemList(i).FlastmwdivPrev <> ojaego.FItemList(i).Flastmwdiv) then %>
					<%= ojaego.FItemList(i).FlastmwdivPrev %> -&gt;
				<% end if %>
			<% end if %>
			<%= ojaego.FItemList(i).Flastmwdiv %>
		</td>
		<td align=center>
			<% if Not IsNull(ojaego.FItemList(i).FmakeridPrev) then %>
				<% if (ojaego.FItemList(i).FmakeridPrev <> ojaego.FItemList(i).Fmakerid) then %>
					<%= ojaego.FItemList(i).FmakeridPrev %> -&gt;
				<% end if %>
			<% end if %>
			<%= ojaego.FItemList(i).Fmakerid %>
		</td>

		<td>

	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan=17 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>

</table>
<% end if %>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
