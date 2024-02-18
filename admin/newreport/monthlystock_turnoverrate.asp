<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
dim yyyy1,mm1,isusing,sysorreal,mwgubun,makerid
yyyy1 = request("yyyy1")
mm1 = request("mm1")
isusing = request("isusing")
sysorreal = request("sysorreal")
mwgubun = request("mwgubun")
makerid = request("makerid")

if sysorreal="" then sysorreal="real"
dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectIsUsing = isusing
ojaego.FRectGubun = sysorreal
ojaego.FRectMakerid = makerid
ojaego.FRectMwDiv = mwgubun

if makerid<>"" then
	ojaego.GetMonthlyRealJeagoDetailByMaker
else
	ojaego.GetMonthlyRealJeagoDetail
end if

dim i
dim totno, totbuy, totsell
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ������ ����ڻ�
			&nbsp;&nbsp;
			<font color="#CC3333">�귣�� :</font> <% drawSelectBoxDesignerwithName "makerid",makerid %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">�����:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >�ý������
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >�ǻ����
        	&nbsp;&nbsp;&nbsp;
        	<font color="#CC3333">��ǰ��뱸��:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >��ü
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >�����
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >������
        	&nbsp;&nbsp;&nbsp;
        	<font color="#CC3333">���Ա���:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >��ü
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >����
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >��Ź
        	<input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >��ü
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<% if makerid<>"" then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">��ǰ�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="60">���Ա���</td>
    	<td width="60">������</td>
    	<td width="70">�Һ��ڰ���</td>
    	<td width="60">��ո���</td>
    	<td width="70">���԰���</td>
    	<td>���</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <% if (ojaego.FItemList(i).FIsUsing="N") or (ojaego.FItemList(i).FOptionUsing="N") then %>
    <tr align="center" bgcolor="#CCCCCC">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
    	<td><a href="javascript:TnPopItemStock('<%= ojaego.FItemList(i).FItemID %>');"><%= ojaego.FItemList(i).FItemID %></a></td>
    	<td align="left"><%= ojaego.FItemList(i).FItemName %></td>
    	<td><%= ojaego.FItemList(i).FItemOptionName %></td>
    	<td><%= fncolor(ojaego.FItemList(i).FMaeIpGubun,"mw") %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td>
    		<% if ojaego.FItemList(i).FTotSellSum<>0 then %>
    		<%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %>
    		<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="4">�Ѱ�</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>

    	<td></td>
    </tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2">�귣��ID</td>
    	<td rowspan="2" width="50">�����<br>����</td>
    	<td rowspan="2" width="70">������<br>(���ް�)<br>(S)</td>
    	
    	<td colspan="4">����3���� ȸ����</td>
    	<td colspan="4"><%= yyyy1 %>�� <%= mm1 %>�� ȸ����</td>
    	
    	<td rowspan="2" width="80"><%= yyyy1 %>�� <%= mm1 %>��<br>����(�԰�)��<br>(���ް�)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="70">ON�����<br>(�ǸŰ�)(O)</td>
    	<td width="70">����<br>(�ǸŰ�)(C)</td>
    	<td width="90"><b>�Ǹ�/����Ѿ�<br>(O+C)</b></td>
    	<td width="90"><font color="blue"><b>ȸ����<br>(R=(O+C)/S)</b></font></td>
    	
    	<td width="70">ON�����<br>(�ǸŰ�)(O)</td>
    	<td width="70">����<br>(�ǸŰ�)(C)</td>
    	<td width="90"><b>�Ǹ�/����Ѿ�<br>(O+C)</b></td>
    	<td width="90"><font color="blue"><b>ȸ����<br>(R=(O+C)/S)</b></font></td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <% if ojaego.FItemList(i).FMakerUsing="Y" then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
    <% end if %>
    	<td align="left"><a href="monthlystock_detail.asp?menupos=<%= menupos %>&mwgubun=<%= ojaego.FItemList(i).FMaeIpGubun %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&sysorreal=<%= sysorreal %>&isusing=<%= isusing %>&makerid=<%= ojaego.FItemList(i).FMakerid %>" ><%= ojaego.FItemList(i).FMakerid %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><b><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %><b></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ѱ�</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
		<td></td>
		<td></td>
		<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    </tr>
</table>
<% end if %>

<%
set ojaego = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->