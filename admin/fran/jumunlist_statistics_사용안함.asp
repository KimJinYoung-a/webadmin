<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �ֹ������� ���
' History : 2010.06.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, shopid ,designer, statecd, baljucode ,notipgo, minusjumun, shopdiv ,totaljumunsuply, totalfixsuply
dim yyyy1,mm1 ,dd1,yyyy2,mm2,dd2, totaljumunsellcash ,i ,fromDate ,toDate , datefg ,parameter , arridx
dim tot10totalsellcash ,tot90totalsellcash , tot70totalsellcash , tot10totalsuplycash , tot90totalsuplycash ,tot70totalsuplycash
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	designer = request("designer")
	statecd  = request("statecd")
	baljucode= request("baljucode")
	notipgo = request("notipgo")
	minusjumun = request("minusjumun")
	shopdiv = request("shopdiv")
	shopid = request("shopid")
	page = request("page")
	menupos = request("menupos")
	datefg = request("datefg")	

	if page="" then page=1
				
if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))-1
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

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
		designer = session("ssBctID")	'"7321"

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if	

dim osheet
set osheet = new COrderSheet
	osheet.FRectFromDate = fromDate
	osheet.FRectToDate = toDate
	osheet.frectdatefg = datefg
	osheet.FCurrPage = page
	osheet.Fpagesize=20
	
	if baljucode<>"" then
		osheet.FRectBaljuCode = baljucode
	else
		osheet.FRectBaljuid = shopid
		osheet.FRectStatecd = statecd
		osheet.FRectMakerid = designer
		osheet.FRectDivCodeArr = "('501','502','503')"
		osheet.FRectNotIpgoOnly = notipgo
		osheet.FRectMinusOnly = minusjumun
		osheet.FRectshopdiv = shopdiv
	end if
	
	osheet.GetOrderSheetList_statistics
	
parameter = "shopid="&shopid&"&notipgo="& notipgo &"&minusjumun="& minusjumun &"&shopdiv="& shopdiv &"&statecd="& statecd &"&designer="& designer &"&menupos="& menupos
parameter = parameter & "&yyyy1="&yyyy1&"&mm1="& mm1&"&dd1="& dd1&"&yyyy2="& yyyy2&"&mm2="& mm2&"&dd2="& dd2&"&baljucode="& baljucode&"&datefg="& datefg

tot10totalsellcash = 0
tot90totalsellcash = 0
tot70totalsellcash = 0
tot10totalsuplycash = 0
tot90totalsuplycash = 0
tot70totalsuplycash = 0
%>

<script language='javascript'>

function pop_detail(idx){
	var pop_detail = window.open('jumunlist_statistics_detail.asp?idx='+idx,'pop_detail','width=1024,height=768,resizable=yes,scrollbars=yes');
	pop_detail.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�귣������ : <% drawSelectBoxDesignerwithName "designer", designer %>
		<br>
		�ֹ��ڵ� : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
		�ֹ����� :
		<select name="statecd" class="select">
			<option value="">��ü
			<option value=" " <% if statecd=" " then response.write "selected" %> >�ۼ���
			<option value="0" <% if statecd="0" then response.write "selected" %> >�ֹ�����
			<option value="1" <% if statecd="1" then response.write "selected" %> >�ֹ�Ȯ��
			<option value="2" <% if statecd="2" then response.write "selected" %> >�Աݴ��
			<option value="5" <% if statecd="5" then response.write "selected" %> >����غ�
			<option value="6" <% if statecd="6" then response.write "selected" %> >�����
			<option value="7" <% if statecd="7" then response.write "selected" %> >���Ϸ�
			<option value="8" <% if statecd="8" then response.write "selected" %> >�԰���
			<option value="9" <% if statecd="9" then response.write "selected" %> >�԰�Ϸ�
		</select>
		<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >����ó����
     	&nbsp;
     	<input type="checkbox" name="minusjumun" <% if minusjumun="on" then response.write "checked" %> >���̳ʽ��ֹ���
     	&nbsp;     	
     	SHOP���� : 
     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >��ü
		<input type="radio" name="shopdiv" value="j" <% if shopdiv="j" then response.write "checked" %> >����
		<input type="radio" name="shopdiv" value="f" <% if shopdiv="f" then response.write "checked" %> >������			
	</td>		
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>					
		��¥���� : 
		<% drawipgo_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type=hidden name="idxarr">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= osheet.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= osheet.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�ֹ��ڵ�</td>
	<td>������</td>
	<td>���޹޴���<br>�ֹ���(SHOP)</td>
	<td>�ֹ�����</td>
	<td>�ֹ���/<br>�԰�(��û)��</td>
	<td>���ֹ���<br>(�Һ��ڰ�)</td>
	<td>���ֹ���<br>(���ް�)</td>
	<td>Ȯ���ݾ�<br>(���ް�)</td>
	<td>�����</td>
	<td>��ǰ���к�<br>Ȯ���ݾ�(�ǸŰ�)</td>
	<td>��ǰ���к�<br>Ȯ���ݾ�(���ް�)</td>	
	<td>���</td>
</tr>
<% if osheet.FResultCount >0 then %>
<% 
for i=0 to osheet.FResultcount-1

arridx = arridx & osheet.FItemList(i).fidx & ","

tot10totalsellcash = tot10totalsellcash + osheet.FItemList(i).f10totalsellcash
tot90totalsellcash = tot90totalsellcash + osheet.FItemList(i).F90totalsellcash
tot70totalsellcash = tot70totalsellcash + osheet.FItemList(i).F70totalsellcash
tot10totalsuplycash = tot10totalsuplycash + osheet.FItemList(i).F10totalsuplycash
tot90totalsuplycash = tot90totalsuplycash + osheet.FItemList(i).F90totalsuplycash
tot70totalsuplycash = tot70totalsuplycash + osheet.FItemList(i).F70totalsuplycash
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash

if osheet.FItemList(i).Ftargetid="10x10" then
	totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
	totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
else
	totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunbuycash
	totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalbuycash
end if
%>
<tr bgcolor="#FFFFFF" align="center">
	<td rowspan=2 >
		<%= osheet.FItemList(i).Fbaljucode %>
	</td>
	<% if osheet.FItemList(i).Ftargetid<>"10x10" then %>
	<td rowspan=2 ><b><%= osheet.FItemList(i).Ftargetid %></b><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% else %>
	<td rowspan=2 ><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<% end if %>
	<td rowspan=2 ><%= osheet.FItemList(i).Fbaljuid %><br>(<%= osheet.FItemList(i).Fbaljuname %>)</td>
	<td rowspan=2 >
		<font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font>
		<br><%= osheet.FItemList(i).FAlinkCode %>
	</td>
	<td ><font color="#777777"><%= Left(osheet.FItemList(i).FRegdate,10) %></font></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<% if osheet.FItemList(i).Ftargetid="10x10" then %>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<% else %>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunbuycash,0) %></td>
	<td align="right"><%= FormatNumber(osheet.FItemList(i).Ftotalbuycash,0) %></td>
	<% end if %>
	<td ><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
	<td rowspan=2>
		<%= FormatNumber(osheet.FItemList(i).f10totalsellcash,0) %> (10)
		<br><%= FormatNumber(osheet.FItemList(i).f90totalsellcash,0) %> (90)
		<br><%= FormatNumber(osheet.FItemList(i).f70totalsellcash,0) %> (70)
	</td>
	<td rowspan=2>
		<%= FormatNumber(osheet.FItemList(i).f10totalsuplycash,0) %> (10)
		<br><%= FormatNumber(osheet.FItemList(i).f90totalsuplycash,0) %> (90)
		<br><%= FormatNumber(osheet.FItemList(i).f70totalsuplycash,0) %> (70)
	</td>	
	<td rowspan=2>
		<input type="button" onclick="pop_detail('<%= osheet.FItemList(i).fidx %>');" class="button" value="��">
	</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>
	    <% if IsNULL(osheet.FItemList(i).FIpgodate) then %>
	    	<%= Left(osheet.FItemList(i).Fscheduledate,10) %>
	    <% else %>
	    	<%= Left(osheet.FItemList(i).FIpgodate,10) %>
	    <% end if %>
	</td>
    <td>    	

    </td>
	<td colspan=3><font color="#777777"><%= DdotFormat(osheet.FItemList(i).Fbrandlist,30) %></font></td>
	
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=6>�Ѱ�</td>
	<td align="right"><%= formatNumber(totaljumunsellcash,0) %></td>
	<td align="right"><%= formatNumber(totaljumunsuply,0) %></td>
	<td align="right"><%= formatNumber(totalfixsuply,0) %></td>
	<td align="center">
		<%= FormatNumber(tot10totalsellcash,0) %> (10)
		<br><%= FormatNumber(tot90totalsellcash,0) %> (90)
		<br><%= FormatNumber(tot70totalsellcash,0) %> (70)	
	</td>
	<td align="center">
		<%= FormatNumber(tot10totalsuplycash,0) %> (10)
		<br><%= FormatNumber(tot90totalsuplycash,0) %> (90)
		<br><%= FormatNumber(tot70totalsuplycash,0) %> (70)	
	</td>
	<td align="center">
		<input type="button" onclick="pop_detail('<%= arridx %>');" class="button" value="���ջ�">
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if osheet.HasPreScroll then %>
			<a href="?page=<%= osheet.StartScrollPage-1 %>&<%=parameter%>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
			<% if i>osheet.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&<%=parameter%>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osheet.HasNextScroll then %>
			<a href="?page=<%= i %>&<%=parameter%>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set osheet = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
