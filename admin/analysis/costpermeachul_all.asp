<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/costPerMeachulCls.asp"-->
<%

dim i, t

dim yyyy1,mm1,isusing,sysorreal, research, shopid, designer
dim vatyn
dim prevyyyy1, prevmm1
dim yyyy2, mm2, dd2
dim targetGbn, stockPlace, defaultmargin, pageSize
dim showShopID, showMakerID, sortBy, ordrBy

yyyy1     = RequestCheckVar(request("yyyy1"),10)
mm1       = RequestCheckVar(request("mm1"),10)
isusing   = RequestCheckVar(request("isusing"),10)
sysorreal = RequestCheckVar(request("sysorreal"),10)
research  = RequestCheckVar(request("research"),10)
shopid    = RequestCheckVar(request("shopid"),32)
designer  = RequestCheckVar(request("designer"),32)
vatyn     = RequestCheckVar(request("vatyn"),10)
targetGbn 	= RequestCheckVar(request("targetGbn"),10)
stockPlace	= RequestCheckVar(request("stockPlace"),10)
defaultmargin	= RequestCheckVar(request("defaultmargin"),10)
pageSize	= RequestCheckVar(request("pageSize"),10)
showShopID	= RequestCheckVar(request("showShopID"),10)
showMakerID	= RequestCheckVar(request("showMakerID"),10)
sortBy	= RequestCheckVar(request("sortBy"),10)
ordrBy	= RequestCheckVar(request("ordrBy"),10)

if (defaultmargin <> "") then
	if Not IsNumeric(defaultmargin) then
		response.write "<script>alert('������ ���ڸ� �����մϴ�.[" & defaultmargin & "]')</script>"
		defaultmargin = ""
	end if
end if

sysorreal="sys"
vatyn="Y"

if (vatyn="") then vatyn="Y"
if (pageSize="") then pageSize="200"

if (sortBy = "") then
	sortBy = "sortBy1"
end if

if (ordrBy = "") then
	ordrBy = "ordrBy1"
end if


dim nowdate
if yyyy1="" then
	'// ����
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

'// ���� ����
nowdate = DateAdd("m", 1, (yyyy1 + "-" + mm1 + "-01"))
nowdate = DateAdd("d", -1, nowdate)
yyyy2 = Left(CStr(nowdate),4)
mm2 = Mid(CStr(nowdate),6,2)
dd2 = Right(CStr(nowdate),2)

'// ������
nowdate = DateAdd("m", -1, (yyyy1 + "-" + mm1 + "-01"))
prevyyyy1 = Left(CStr(nowdate),4)
prevmm1 = Mid(CStr(nowdate),6,2)



'// ===========================================================================
dim oCostPerMeachul
set oCostPerMeachul = new CCostPerMeachul

oCostPerMeachul.FRectShopID   = shopid
oCostPerMeachul.FPageSize   = pageSize
oCostPerMeachul.FRectMakerID   = designer
oCostPerMeachul.FRectYYYYMM   = yyyy1 + "-" + mm1
oCostPerMeachul.FRectTargetGbn = targetGbn
oCostPerMeachul.FRectStockPlace = stockPlace
oCostPerMeachul.FRectDefaultmargin = defaultmargin
oCostPerMeachul.FRectShowShopID = showShopID
oCostPerMeachul.FRectShowMakerID = showMakerID

oCostPerMeachul.FRectSortBy = sortBy
oCostPerMeachul.FRectOrdrBy = ordrBy

oCostPerMeachul.GetCostPerMeachulList


dim itemcost, itemcostpermeachul, itemcostpermeachulunit, pointprice
dim totbuysumprevmonth, totbuysumthismonth, totmeachul, totmeaip
dim totitemcost, totitemcostpermeachul
dim itemrotationrate, itemgainlossrate
dim shopbuysumdiff
dim avgshopbuy
%>
<script language='javascript'>

function jsSetMakerID(designer) {
	var frm = document.frm;

	frm.designer.value = designer;
	frm.submit();
}


</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ��
			&nbsp;&nbsp;
			���� :
			<select class="select" name="targetGbn">
				<option value="">��ü</option>
				<option value="ON" <% if (targetGbn = "ON") then %>selected<% end if %> >ON</option>
				<option value="OF" <% if (targetGbn = "OF") then %>selected<% end if %> >OF</option>
			</select>
			&nbsp;&nbsp;
			�⺻���� :
			<input type="text" class="text" name="defaultmargin" size="6" value="<%= defaultmargin %>">
			&nbsp;&nbsp;
			���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
			�귣�� :
			<% drawSelectBoxDesignerwithName "designer", designer %>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">����ڻ� ����:</font>
        	<input type="radio" name="sysorreal" value="sys" checked >�ý������
			&nbsp;&nbsp;
			<input type="checkbox" name="showShopID" value="Y" <% if (showShopID = "Y") then %>checked<% end if %> > ����ǥ��
			&nbsp;&nbsp;
			<input type="checkbox" name="showMakerID" value="Y" <% if (showMakerID = "Y") then %>checked<% end if %> > �귣��ǥ��
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">�ΰ���:</font>
        	<input type="radio" name="vatyn" value="Y" checked >����
        	&nbsp;&nbsp;
			<font color="#CC3333">���̳ʽ����:</font>
        	<input type="radio" name="incminusstockyn" value="Y" checked >����
			&nbsp;&nbsp;
			ǥ�ð��� :
			<select class="select" name="pageSize">
				<option value="50" <% if (pageSize = "50") then %>selected<% end if %> >50</option>
				<option value="200" <% if (pageSize = "200") then %>selected<% end if %> >200</option>
				<option value="500" <% if (pageSize = "500") then %>selected<% end if %> >500</option>
				<option value="1000" <% if (pageSize = "1000") then %>selected<% end if %> >1000</option>
			</select>
			&nbsp;&nbsp;
			ǥ�ü��� :
			<input type="radio" name="sortBy" value="sortBy1" <% if (sortBy = "sortBy1") then %>checked<% end if %> >���ͷ�
			<input type="radio" name="sortBy" value="sortBy2" <% if (sortBy = "sortBy2") then %>checked<% end if %> >���;�
			<input type="radio" name="sortBy" value="sortBy3" <% if (sortBy = "sortBy3") then %>checked<% end if %> disabled>���;�(�����)
			&nbsp;&nbsp;
			���ļ��� :
			<input type="radio" name="ordrBy" value="ordrBy1" <% if (ordrBy = "ordrBy1") then %>checked<% end if %> >��������
			<input type="radio" name="ordrBy" value="ordrBy2" <% if (ordrBy = "ordrBy2") then %>checked<% end if %> >��������
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>


<br><br><font size=5>�������Դϴ�.</font><br><br>

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">����</td>
    	<td width="30">����</td>
    	<td width="30">�����ġ</td>
		<td width="100">����</td>
		<td width="150">�귣��</td>
		<td width="30">��ǰ����</td>
		<td width="80">��ǰ�ڵ�</td>
		<td width="30">�⺻����</td>

		<td width="100">�������<br>(A)</td>
		<td width="100">�������<br>(B)</td>
		<td width="100">����<br>(C)</td>
		<td width="100">�⸻���<br>(D)</td>

		<td width="100">�������<br>(E=A+B-D)</td>
		<td width="60">������<br>(F=E/C)</td>
		<td width="60">����<br>(G=100-F)</td>
		<td width="100">������</td>
		<td width="100">���ȸ����</td>

    	<td>���</td>
    </tr>
    <%
    totbuysumprevmonth = 0
    totbuysumthismonth = 0
    totmeachul = 0
    totmeaip = 0
    %>
    <% for i=0 to oCostPerMeachul.FResultCount-1 %>

	<%

	totbuysumprevmonth = totbuysumprevmonth + oCostPerMeachul.FItemList(i).FbuySumPrevMonth
	totbuysumthismonth = totbuysumthismonth + oCostPerMeachul.FItemList(i).FbuySumThisMonth
	totmeachul = totmeachul + oCostPerMeachul.FItemList(i).FcustomerMeachul
	totmeaip = totmeaip + oCostPerMeachul.FItemList(i).FipgoMeaip

	'// �������
	itemcost = oCostPerMeachul.FItemList(i).FbuySumPrevMonth + oCostPerMeachul.FItemList(i).FipgoMeaip - oCostPerMeachul.FItemList(i).FbuySumThisMonth

	'// ������, ����
	if (oCostPerMeachul.FItemList(i).FcustomerMeachul = 0) then
		itemcostpermeachul = "--"
		itemgainlossrate = "--"
	else
		t = (itemcost / oCostPerMeachul.FItemList(i).FcustomerMeachul) * 100.0

		itemcostpermeachul = FormatNumber(t, 1)
		itemgainlossrate = FormatNumber((100.0 - t), 1)
	end if

	'// �������ڻ�
    avgshopbuy = (oCostPerMeachul.FItemList(i).FbuySumPrevMonth + oCostPerMeachul.FItemList(i).FbuySumThisMonth) / 2

	'���ȸ����
	if (avgshopbuy = 0) or isNULL(avgshopbuy) then
		itemrotationrate = "--"
	else
		t = (itemcost / avgshopbuy) * 100.0
		itemrotationrate = FormatNumber(t, 1)
	end if

	%>
    <tr align="center" bgcolor="#FFFFFF" hright=30>
		<td><%= oCostPerMeachul.FItemList(i).Fyyyymm %></td>
		<td><%= oCostPerMeachul.FItemList(i).FtargetGbn %></td>
		<td><%= oCostPerMeachul.FItemList(i).FstockPlace %></td>
		<td><%= oCostPerMeachul.FItemList(i).Fshopid %></td>
		<td>
			<% if (showMakerID = "Y") and (oCostPerMeachul.FItemList(i).Fmakerid = "") then %>
			<a href="javascript:jsSetMakerID('����')">����</a>
			<% else %>
			<a href="javascript:jsSetMakerID('<%= oCostPerMeachul.FItemList(i).Fmakerid %>')"><%= oCostPerMeachul.FItemList(i).Fmakerid %></a>
			<% end if %>
		</td>
		<td><%= oCostPerMeachul.FItemList(i).Fitemgubun %></td>
		<td><%= oCostPerMeachul.FItemList(i).Fitemid %></td>
		<td><%= oCostPerMeachul.FItemList(i).Fdefaultmargin %></td>
		<td align="right" style="padding-right: 8px"><%= FormatNumber(oCostPerMeachul.FItemList(i).FbuySumPrevMonth,0) %></td>
		<td align="right" style="padding-right: 8px"><%= FormatNumber(oCostPerMeachul.FItemList(i).FipgoMeaip,0) %></td>
		<td align="right" style="padding-right: 8px"><%= FormatNumber(oCostPerMeachul.FItemList(i).FcustomerMeachul,0) %></td>
		<td align="right" style="padding-right: 8px"><%= FormatNumber(oCostPerMeachul.FItemList(i).FbuySumThisMonth,0) %></td>

		<td align="right" style="padding-right: 8px"><%= FormatNumber(itemcost,0) %></td>
		<td align="right" style="padding-right: 8px"><%= itemcostpermeachul %></td>
		<td align="right" style="padding-right: 8px"><%= itemgainlossrate %></td>
		<td align="right" style="padding-right: 8px"><%= FormatNumber(avgshopbuy,0) %></td>
		<td align="right" style="padding-right: 8px"><%= itemrotationrate %></td>

		<td></td>
	</tr>
    <% next %>
    <%
	totitemcost = totbuysumprevmonth + totmeaip - totbuysumthismonth
    %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="8">�Ѱ�</td>
    	<td align="right" ><%= FormatNumber(totbuysumprevmonth, 0) %></td>
		<td align="right" ><%= FormatNumber(totmeaip, 0) %></td>
		<td align="right" ><%= FormatNumber(totmeachul, 0) %></td>
		<td align="right" ><%= FormatNumber(totbuysumthismonth, 0) %></td>
    	<td align="right" colspan="6"></td>
    </tr>
</table>
<%
set oCostPerMeachul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
