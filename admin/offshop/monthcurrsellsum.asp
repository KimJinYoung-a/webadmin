<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �������������
' History : 2010.06.15 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim opt_rect ,yyyy1,mm1, oldlist ,shopid ,research ,stdate ,i,p1,p2,p3,p4 ,maybe_monthcount
dim maybe_monthsum, dayno, currno, nowdate, nowyyyymm, inc3pl
	opt_rect = requestCheckVar(request("opt_rect"),32)
	research = requestCheckVar(request("research"),2)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1	  = requestCheckVar(request("mm1"),2)
	oldlist = requestCheckVar(request("oldlist"),10)
	shopid = requestCheckVar(request("shopid"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if research<>"on" then
	if opt_rect="" then opt_rect="all"
end if

if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

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
	
dim oreport
set oreport = new COffShopSell
	oreport.FRectSearchType = opt_rect
	oreport.FRectFromDate = yyyy1 + "-" + mm1 + "-01"
	oreport.FRectOldJumun = oldlist
	oreport.FRectShopID = shopid
	oreport.FRectInc3pl = inc3pl	
	oreport.Getoffmonthlysum

if opt_rect="all" then
	nowdate = CStr(date)
	nowyyyymm = left(nowdate,7)
	currno = CInt(right(nowdate,2))

	nowdate = dateserial(Left(nowdate,4),Mid(nowdate,6,2)+1,0)
	dayno = CInt(right(nowdate,2))
end if
%>

<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% DrawYMBox yyyy1,mm1 %>~����
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3��������
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
				<p>
				<input type="radio" name="opt_rect" value="curr" <% if opt_rect="curr" then response.write "checked" %> >�ſ���~���ó�¥
				&nbsp;&nbsp;
				<input type="radio" name="opt_rect" value="all" <% if opt_rect="all" then response.write "checked" %> >�ſ���~����
	            &nbsp;&nbsp;
	            <b>* ����ó����</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>					
			</td>
		</tr>
		</table> 
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>	
</form>
</table>
<!-- ǥ ��ܹ� ��-->
<br>
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

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120">�Ⱓ</td>
	<td></td>
	<td width="100">�ݾ�(��)</td>
	<td width="50">�Ǽ�</td>
</tr>
<% if oreport.FresultCount>0 then %>
<% 
for i=0 to oreport.FresultCount-1

if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then

	if oreport.FItemList(i).Fselltotal<>0 and oreport.FItemList(i).Fselltotal <> "" then
		maybe_monthsum	 = (oreport.FItemList(i).Fselltotal * dayno / currno)
	end if

	if oreport.FItemList(i).Fsellcnt<>0 and oreport.FItemList(i).Fsellcnt <> "" then
		maybe_monthcount = (oreport.FItemList(i).Fsellcnt * dayno / currno)
	end if
	
	if maybe_monthcount>oreport.maxc then
		oreport.maxc = maybe_monthcount
	end if

	if maybe_monthsum>oreport.maxt then
		oreport.maxt = maybe_monthsum
	end if
	
	if maybe_monthsum <> 0 and maybe_monthsum <> "" and oreport.maxt <> 0 and oreport.maxt <> "" then
		p3 = Clng(maybe_monthsum/oreport.maxt*100)
	else
		p3 = 0
	end if
	
	if maybe_monthcount <> 0 and maybe_monthcount <> "" and oreport.maxc <> 0 and oreport.maxc <> "" then
		p4 = Clng(maybe_monthcount/oreport.maxc*100)
	else
		p4 = 0
	end if	
end if

if oreport.FItemList(i).Fselltotal <> 0 and oreport.FItemList(i).Fselltotal <> "" and oreport.maxt <> 0 and oreport.maxt <> "" then
	p1 = Clng(oreport.FItemList(i).Fselltotal/oreport.maxt*100)
else
	p1 = 0
end if

if oreport.FItemList(i).Fsellcnt <> 0 and oreport.FItemList(i).Fsellcnt <> "" and oreport.maxc <> 0 and oreport.maxc <> "" then
	p2 = Clng(oreport.FItemList(i).Fsellcnt/oreport.maxc*100)
else
	p2 = 0
end if	
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF';>
	<td width="120">
		<%= oreport.FItemList(i).Fsitename %>
	</td>
	<td>
		<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
			<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p3 %>%"></div><br>
			<div align="left"> <img src="/images/dot4.gif" height="3" width="<%= p4 %>%"></div><br>
		<% end if %>
		<div align="left"><img src="/images/dot1.gif" height="3" width="<%= p1 %>%"></div><br>
		<div align="left"><img src="/images/dot2.gif" height="3" width="<%= p2 %>%"></div>
	</td>
	<td class="a" width="100" align="right">
		<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
			<font color="#AAAAAA"><%= FormatNumber(maybe_monthsum,0) %></font><br>
		<% end if %>
		<%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %><br>
	</td>
	<td class="a" width="50" align="right">
		<% if Left(oreport.FItemList(i).Fsitename,7)=nowyyyymm then %>
			<font color="#AAAAAA"><%= FormatNumber(maybe_monthcount,0) %></font><br>
		<% end if %>
		<%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %>
	</td>
</tr>   
<% 
next
else
%>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
