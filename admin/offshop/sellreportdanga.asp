<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �Ⱓ�����ܰ�
' History : 2009.04.07 ������ ����
'			2010.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid ,oldlist ,fromDate,toDate ,yyyymmdd1,yyymmdd2 ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,i
dim inc3pl
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	oldlist = requestCheckVar(request("oldlist"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if page="" then page=1
	
if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-14)
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

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopID = shopid
	ooffsell.FCurrPage=page
	ooffsell.FRectOldData = oldlist
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectInc3pl = inc3pl
	
	if shopid<>"" then
		ooffsell.GetReportByDanga
	else
		response.write "<script language='javascript'>"
		response.write "	alert('������ ������ �ּ���');"
		response.write "</script>"
	end if
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<!--
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3������������
			&nbsp;
		-->
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

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
		�˻���� : <b><%= ooffsell.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>����Ǽ�</td>
	<td>�ѰǼ����%</td>
	<td>�����<Br>(���ϸ�������)</td>
	<td>�Ѹ�����%</td>
</tr>
<% if ooffsell.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
<% 
for i=0 to ooffsell.FresultCount-1
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff'; align="center">
	<td><%= ooffsell.FItemList(i).FTerm %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FCount,0) %></td>
	<td align="right">
		<% if ooffsell.maxc<>0 then %>
			<%= CLng(ooffsell.FItemList(i).FCount/ooffsell.maxc*100) %> %
		<% end if %>
	</td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).fspendmile,0) %></td>
	<td align="right">
		<% if ooffsell.maxt<>0 then %>
			<%= CLng( (ooffsell.FItemList(i).FSum+ooffsell.FItemList(i).fspendmile) /ooffsell.maxt*100) %> %
		<% end if %>
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
	<td align="center">�Ѱ�</td>
	<td align="right"><%= FormatNumber(ooffsell.maxc,0) %></td>
	<td></td>
	<td align="right"><%= FormatNumber(ooffsell.maxt,0) %></td>
	<td></td>
</tr>
<% end if %>
</table>

<%
set ooffsell= Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->