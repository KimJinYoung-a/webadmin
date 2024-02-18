<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->
<%
Const isShowIpgoPrice = FALSE

dim i, t

dim yyyy1,mm1,isusing,sysorreal, research, shopid, designer
dim mwgubun
dim vatyn
dim incminusstockyn
dim prevyyyy1, prevmm1
dim yyyy2, mm2, dd2

yyyy1     = RequestCheckVar(request("yyyy1"),10)
mm1       = RequestCheckVar(request("mm1"),10)
isusing   = RequestCheckVar(request("isusing"),10)
sysorreal = RequestCheckVar(request("sysorreal"),10)
research  = RequestCheckVar(request("research"),10)
shopid    = RequestCheckVar(request("shopid"),32)
designer  = RequestCheckVar(request("designer"),32)
mwgubun   = RequestCheckVar(request("mwgubun"),10)
vatyn     = RequestCheckVar(request("vatyn"),10)
incminusstockyn     = RequestCheckVar(request("incminusstockyn"),10)

if (sysorreal="") then sysorreal="sys" ''real
''sysorreal="sys"

if (vatyn="") then vatyn="Y"
if (incminusstockyn="") then incminusstockyn="Y"

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
dim oshopcostpermeachul
set oshopcostpermeachul = new COffShopCostPerMeachul

oshopcostpermeachul.FRectShopID   = shopid
''oshopcostpermeachul.FRectMakerID   = designer
oshopcostpermeachul.FRectYYYYMM   = yyyy1 + "-" + mm1
oshopcostpermeachul.FRectGubun = sysorreal
oshopcostpermeachul.FRectMWDiv    = mwgubun

if (shopid <> "") then
	oshopcostpermeachul.GetOffShopCostPerMeachulByBrandList
end if

dim itemcost, itemcostpermeachul, itemcostpermeachulunit, pointprice
dim totshopbuysumprevmonth, totshopbuysumthismonth, totshopmeachul, totshopmeaip, totshoperrorthismonth
dim totitemcost, totitemcostpermeachul, totshopminusprevmonth, totshopminusthismonth
dim itemrotationrate, itemgainlossrate
dim shopbuysumdiff
dim avgshopbuy
%>
<script language='javascript'>

function PopPrevMonthStock(makerid, mwgubun) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= prevyyyy1 %>";
	var mm1 = "<%= prevmm1 %>";
	var showminus = "<% if (incminusstockyn = "Y") then %>on<% end if %>";

	var popwin = window.open('/admin/newreport/monthlystockShop_detail.asp?research=on&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&shopid=' + shopid + '&makerid=' + makerid + '&sysorreal=<%=sysorreal%>&mwgubun=' + mwgubun + '&showminus=' + showminus,'PopPrevMonthStock','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthStock(makerid, mwgubun) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var showminus = "<% if (incminusstockyn = "Y") then %>on<% end if %>";

	var popwin = window.open('/admin/newreport/monthlystockShop_detail.asp?research=on&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&shopid=' + shopid + '&makerid=' + makerid + '&sysorreal=<%=sysorreal%>&mwgubun=' + mwgubun + '&showminus=' + showminus,'PopPrevMonthStock','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthError(makerid) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var dd1 = "01";
	var yyyy2 = "<%= yyyy2 %>";
	var mm2 = "<%= mm2 %>";
	var dd2 = "<%= dd2 %>";

	var popwin = window.open('/admin/stock/off_baditem_list.asp?shopid=' + shopid + '&makerid=' + makerid + '&itembarcode=&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2 + '&errType=D','PopThisMonthError','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopThisMonthMeachul(makerid) {
	var shopid = "<%= shopid %>";
	var yyyy1 = "<%= yyyy1 %>";
	var mm1 = "<%= mm1 %>";
	var dd1 = "01";
	var yyyy2 = "<%= yyyy2 %>";
	var mm2 = "<%= mm2 %>";
	var dd2 = "<%= dd2 %>";

	var popwin = window.open('/admin/offshop/brandselldetail.asp?yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2 + '&designer=' + makerid + '&datefg=maechul&shopid=' + shopid + '&offgubun=&oldlist=','PopThisMonthMeachul','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopShopMakerMonthlyMaeip(makerid, jungsangubun) {
	var shopid = "<%= shopid %>";
	var yyyymm = "<%= yyyy1 %>-<%= mm1 %>";

	var popwin = window.open('/admin/analysis/popShopMaeip.asp?shopid=' + shopid + '&yyyymm=' + yyyymm + '&makerid=' + makerid + '&jungsangubun=' + jungsangubun,'PopShopMakerMonthlyMaeip','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
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
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> �� ������
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
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
        	&nbsp;&nbsp;
        	<font color="#CC3333">��� ���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> > ��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> > ����(�������+������)
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> > ��Ź(��ü��Ź+��Ź�Ǹ�)
        	<input type="radio" name="mwgubun" value="B031" <% if mwgubun="B031" then response.write "checked" %> > ������
        	<input type="radio" name="mwgubun" value="B022" <% if mwgubun="B022" then response.write "checked" %> > �������
        	<input type="radio" name="mwgubun" value="B012" <% if mwgubun="B012" then response.write "checked" %> > ��ü��Ź
        	<input type="radio" name="mwgubun" value="B011" <% if mwgubun="B011" then response.write "checked" %> > ��Ź�Ǹ�
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> > ������
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">�ΰ���:</font>
        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >����
        	<!--
        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >����
        	-->
        	&nbsp;&nbsp;
			<font color="#CC3333">���̳ʽ����:</font>
        	<input type="radio" name="incminusstockyn" value="Y" <% if incminusstockyn="Y" then response.write "checked" %> >����
        	<!--
        	<input type="radio" name="incminusstockyn" value="N" <% if incminusstockyn="N" then response.write "checked" %> >����
        	-->
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>


<br><br><font size=5>�������Դϴ�.</font><br><br>

<p>

<% if (shopid = "") then %>
	<script>alert("���� ������ �����ϼ���.");</script>
	<% response.end %>
<% end if %>

* �������, �⸻��� �� ������Ծ��� <font color=red>���� ������԰�, ���� �귣��, ���� ���걸��</font>�� �������� �մϴ�.<br>
* �������, �⸻���� <font color=red>�ý������</font>�� �������� �մϴ�.<br>
* ����������� ���Ի�ǰ�� I = A - D(���̳ʽ���� ����), E = I + B, ��Ź��ǰ�� E = B ���� ���˴ϴ�.<br>
<!--<br>
* ������� �� �⸻���� <font color=red>���̳ʽ����</font>�� ������ �ݾ��Դϴ�.
-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>�귣��</td>
    	<td width="80">���걸��</td>
    	<td width="70">�Һ��ڸ���<br>(C)</td>
    	<td width="70"><%=sysorreal%>����<br>(A)</td>
    	<td width="70">�������<br>(B)</td>

    	<td width="70"><%=sysorreal%>�⸻<br>(D)</td>

    	<td width="70">���������<br>(E)</td>
    	<td width="70">������<br>(F=E/C)</td>
    	<td width="70">����<br>(1 - F)</td>

    	<td width="70">������<br>(G=(D+A)/2)</td>
    	<td width="70">���<br>���ȸ����<br>(E / G)</td>

    	<td width="70">�������(H)</td>
    	<td width="70">�����<br>(I)</td>
    	<td></td>
    </tr>
    <%
    totshopbuysumprevmonth = 0
    totshopbuysumthismonth = 0
    totshopmeachul = 0
    totshopmeaip = 0
    totshoperrorthismonth = 0
    totshopminusprevmonth = 0
    totshopminusthismonth = 0
    %>
    <% for i=0 to oshopcostpermeachul.FResultCount-1 %>

    	<% if (designer = "") or (designer = oshopcostpermeachul.FItemList(i).Fmakerid) then %>
    	<% if (mwgubun = "") or (mwgubun = "M") or (mwgubun = "W") or (mwgubun = oshopcostpermeachul.FItemList(i).Fmwdiv) then %>

    	<%

		if (vatyn <> "Y") and (oshopcostpermeachul.FItemList(i).Fvatyn <> "N") then

			'// �ΰ��� ����
			oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth = oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth = oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth * 10 / 11

			oshopcostpermeachul.FItemList(i).Fshopmeachul = oshopcostpermeachul.FItemList(i).Fshopmeachul * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopmeaip = oshopcostpermeachul.FItemList(i).Fshopmeaip * 10 / 11

			oshopcostpermeachul.FItemList(i).Fshoperrorthismonth = oshopcostpermeachul.FItemList(i).Fshoperrorthismonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopminusprevmonth = oshopcostpermeachul.FItemList(i).Fshopminusprevmonth * 10 / 11
			oshopcostpermeachul.FItemList(i).Fshopminusthismonth = oshopcostpermeachul.FItemList(i).Fshopminusthismonth * 10 / 11

		end if

		totshopbuysumprevmonth 	= totshopbuysumprevmonth + oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth
		totshopbuysumthismonth 	= totshopbuysumthismonth + oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth
		totshopmeachul 			= totshopmeachul + oshopcostpermeachul.FItemList(i).Fshopmeachul
		totshopmeaip 			= totshopmeaip + oshopcostpermeachul.FItemList(i).Fshopmeaip
		totshoperrorthismonth 	= totshoperrorthismonth + oshopcostpermeachul.FItemList(i).Fshoperrorthismonth
		totshopminusprevmonth 	= totshopminusprevmonth + oshopcostpermeachul.FItemList(i).Fshopminusprevmonth
		totshopminusthismonth 	= totshopminusthismonth + oshopcostpermeachul.FItemList(i).Fshopminusthismonth


		'���������(E = A + B - D or B)
		if (oshopcostpermeachul.FItemList(i).Fmwdiv = "B011") or (oshopcostpermeachul.FItemList(i).Fmwdiv = "B012") or (oshopcostpermeachul.FItemList(i).Fmwdiv = "B013") then
			'��Ź
			shopbuysumdiff = 0
			itemcost = oshopcostpermeachul.FItemList(i).Fshopmeaip
		else
			'����
			shopbuysumdiff = oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth
			itemcost = shopbuysumdiff + oshopcostpermeachul.FItemList(i).Fshopmeaip
			''((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth + oshopcostpermeachul.FItemList(i).Fshopmeaip) - oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth)
		end if

		itemcostpermeachul = 0
		itemgainlossrate = 0

		'������, ����
		if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then
			'// �������(������ ���Ұ�)
			itemcostpermeachul = "--"
			itemgainlossrate = "--"
		else
			t = (itemcost / oshopcostpermeachul.FItemList(i).Fshopmeachul) * 100.0

			itemcostpermeachul = FormatNumber(t, 1)
			itemgainlossrate = FormatNumber((100.0 - t), 1)
		end if

		' �������ڻ�(���̳ʽ� ��� ����)
		''avgshopbuy = (((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopminusprevmonth) + (oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth - oshopcostpermeachul.FItemList(i).Fshopminusthismonth)) / 2)
        avgshopbuy = (((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth ) + (oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth )) / 2)

		'���ȸ����
		if (avgshopbuy = 0) or isNULL(avgshopbuy) then
			'// ������(���Ұ�)
			itemrotationrate = "--"
		else
			''t = (oshopcostpermeachul.FItemList(i).Fshopmeachul / avgshopbuy) * 100.0
			t = (itemcost / avgshopbuy) * 100.0
			itemrotationrate = FormatNumber(t, 1)
		end if

    	%>
    <tr align="center" bgcolor="#FFFFFF" hright=30>
        <td><%= oshopcostpermeachul.FItemList(i).Fmakerid %></td>
        <td><font color="<%= oshopcostpermeachul.FItemList(i).GetDivcdColor %>"><%= oshopcostpermeachul.FItemList(i).Fmwname %></font>
        <%= oshopcostpermeachul.FItemList(i).Fdefaultmargin %>
        </td>
        <td align=right>
        	<acronym title="�Һ��ڸ���">
        	<a href="javascript:PopThisMonthMeachul('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>')">
        		<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshopmeachul, 0) %>
        	</a>
        	</acronym>
        </td>

        <td align=right>
        	<acronym title="�������">
        	<a href="javascript:PopPrevMonthStock('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
				<% if (Not IsNull(oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth)) then %>
					<% if (incminusstockyn = "Y") then %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth), 0) %>
					<% else %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumprevmonth - oshopcostpermeachul.FItemList(i).Fshopminusprevmonth), 0) %>
					<% end if %>
				<% else %>
					<font color="red">ERR</font>
				<% end if %>
       			</a>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="�������">
        		<a href="javascript:PopShopMakerMonthlyMaeip('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
        			<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshopmeaip, 0) %>
        		</a>
        	</acronym>
        </td>

        <td align=right>
        	<acronym title="�⸻���">
        		<a href="javascript:PopThisMonthStock('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>', '<%= oshopcostpermeachul.FItemList(i).Fmwdiv %>')">
					<% if (incminusstockyn = "Y") then %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth), 0) %>
					<% else %>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopbuysumthismonth - oshopcostpermeachul.FItemList(i).Fshopminusthismonth), 0) %>
					<% end if %>
        		</a>
        	</acronym>
        </td>

        <td align=right>
			<% if (Not IsNull(itemcost)) then %>
				<acronym title="���������">
				<%= FormatNumber(itemcost, 0) %>
				</acronym>
			<% end if %>
        </td>
        <!-- ���ͱݾ�
        <td align=right>
			<% if (Not IsNull(itemcost)) then %>
				<acronym title="����">
				<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul - itemcost) < 0 then %><font color=red><% end if %>
				<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fshopmeachul - itemcost), 0) %>
				</acronym>
			<% end if %>
        </td>
        -->
        <td align=right>
        	<acronym title="������">
        		<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then %>
        			--
        		<% else %>
	        		<% if (Abs(itemcostpermeachul) > 500.0) then %>
	        			<font color=red><b>ERR</b></font>
	        		<% else %>
	        			<%= itemcostpermeachul %> %
	        		<% end if %>
	        	<% end if %>
    		</acronym>
        </td>
        <td align=right>
        	<acronym title="����">
        		<% if (oshopcostpermeachul.FItemList(i).Fshopmeachul <= 0) then %>
        			--
        		<% else %>
	        		<% if (Abs(itemcostpermeachul) > 500.0) then %>
	        			<font color=red><b>ERR</b></font>
	        		<% else %>
		        		<% if ((itemgainlossrate - oshopcostpermeachul.FItemList(i).Fdefaultmargin) > 5.0) or ((itemgainlossrate - oshopcostpermeachul.FItemList(i).Fdefaultmargin) < -10.0) then %><font color=red><b><% end if %>
		        		<%= itemgainlossrate %> %
	        		<% end if %>
	        	<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="�������ڻ�">
        		<% if (FALSE) and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B031") and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B022") then %>
        			--
        		<% elseif (Not IsNull(avgshopbuy)) then %>
        			<%= FormatNumber(avgshopbuy, 0) %>
        		<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="������ȸ����">
        		<% if (avgshopbuy = 0) then %>
        			--
        		<% else %>
	        		<% if (FALSE) and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B031") and (oshopcostpermeachul.FItemList(i).Fmwdiv <> "B022") then %>
	        			--
	        		<% elseif (Not IsNull(itemrotationrate)) then %>
		        		<% if (Abs(itemrotationrate) >= 1000.0) then %>
		        			<font color=red><b>ERR</b></font> <%'=itemcostpermeachul%>
		        		<% else %>
		        			<% if (itemrotationrate < 30.0) then %><font color=red><b><% end if %>
		        			<% if (itemrotationrate > 60.0) then %><font color=green><% end if %>
		        			<%= itemrotationrate %> %
		        		<% end if %>
	        		<% end if %>
        		<% end if %>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="�������">
        	<a href="javascript:PopThisMonthError('<%= oshopcostpermeachul.FItemList(i).Fmakerid %>')">
        		<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fshoperrorthismonth, 0) %>
        	</a>
        	</acronym>
        </td>
        <td align=right>
        	<acronym title="�����(����-�⸻, ���̳ʽ� ��� ����)">
        		<% if (shopbuysumdiff = 0) then %>
        			--
        		<% else %>
        			<%= FormatNumber(shopbuysumdiff, 0) %>
        		<% end if %>
        	</acronym>
        </td>

    	<td></td>
    </tr>
    	<% end if %>
    	<% end if %>
    <% next %>
    <%
	totitemcost = totshopbuysumprevmonth + totshopmeaip - totshopbuysumthismonth
    %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td >�Ѱ�</td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totshopmeachul, 0) %></td>
    	<td align="right" ><%= FormatNumber((totshopbuysumprevmonth - totshopminusprevmonth), 0) %></td>
    	<td align="right" ><%= FormatNumber(totshopmeaip, 0) %></td>

    	<td align="right" ><%= FormatNumber((totshopbuysumthismonth ), 0) %></td>
    	<td align="right" ></td>
    	<!--
    	<td align="right" ><%= FormatNumber((totshopmeachul - totitemcost), 0) %></td>
    	-->
    	<td align="right" >
    		<!--
    		<% if (totshopmeachul <> 0) then %>
    			<%= FormatNumber((totitemcost / totshopmeachul) * 100, 1) %>%
    		<% else %>
    			--
    		<% end if %>
    		-->
    	</td>
    	<td align="right" ></td>
    	<td align="right" ></td>
    	<td align="right" ></td>
    	<td align="right" ><%= FormatNumber(totshoperrorthismonth, 0) %></td>
    	<td align="right" ><%= FormatNumber((totshopbuysumprevmonth - totshopbuysumthismonth), 0) %></td>
    	<td align="right" ></td>
    </tr>
</table>
<%
set oshopcostpermeachul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
