<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���Ϻ� ����м�
' History : 2010.06.15 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
dim fromDate,toDate ,oldlist ,offgubun ,shopid ,page,i,ix ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,datefg , weekdate ,parameter
dim selltotal_jj, sellcnt_jj, DpartCount_jj ,selltotal ,selltotal_jm, sellcnt_jm, DpartCount_jm ,avgsell,avgselltotal
dim inc3pl
	oldlist = requestCheckVar(request("oldlist"),2)
	offgubun = requestCheckVar(request("offgubun"),32)
	page = requestCheckVar(request("page"),10)
	if page="" then page=1
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	weekdate = requestCheckVar(request("weekdate"),30)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if datefg = "" then datefg = "maechul"	
if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("m",-1,now())))
if (mm1="") then mm1 = Cstr(Month(dateadd("m",-1,now())))
if (dd1="") then dd1 = Cstr(day(dateadd("m",-1,now())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
toDate = Left(CStr(DateSerial(yyyy2,mm2,dd2+1)),10)

selltotal =0

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
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.FRectOldJumun = oldlist
	oreport.FRectOffgubun = offgubun
	oreport.FRectShopID = shopid
	oreport.frectdatefg = datefg
	oreport.frectweekdate = weekdate
	oreport.FRectInc3pl = inc3pl	
	oreport.GetoffWeeklySellReport

parameter = "oldlist="&oldlist&"&offgubun="&offgubun&"&shopid="&shopid&"&datefg="&datefg&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&inc3pl="&inc3pl&"&menupos="&menupos

for i=0 to oreport.FResultCount -1
	selltotal 	= selltotal + oreport.FItemList(i).Fselltotal
	if oreport.FItemList(i).FDpartCount<>0 then
		avgsell		= CLng(oreport.FItemList(i).Fselltotal/oreport.FItemList(i).FDpartCount)
		avgselltotal = avgselltotal + avgsell
	end if

	if oreport.FItemList(i).Fdpart="1" or oreport.FItemList(i).Fdpart="7" then
 		selltotal_jm	= selltotal_jm + oreport.FItemList(i).Fselltotal
 		sellcnt_jm		= sellcnt_jm + oreport.FItemList(i).Fsellcnt
 		DpartCount_jm	= DpartCount_jm + oreport.FItemList(i).FDpartCount
 	else
 		selltotal_jj	= selltotal_jj + oreport.FItemList(i).Fselltotal
 		sellcnt_jj		= sellcnt_jj + oreport.FItemList(i).Fsellcnt
 		DpartCount_jj	= DpartCount_jj + oreport.FItemList(i).FDpartCount
 	end if
next
%>

<script type='text/javascript'>
	
function frmsubmit(){
	frm.submit();
}

function category_sum(weekdate){
	var category_sum = window.open('/admin/offshop/offshop_categorysellsum.asp?weekdate='+weekdate+'&<%=parameter%>','category_sum','width=1024,height=768,scrollbars=yes,resizable=yes');
	category_sum.focus();
}

function best_sum(weekdate){
	var best_sum = window.open('/admin/offshop/offshop_categorybestseller.asp?weekdate='+weekdate+'&<%=parameter%>','best_sum','width=1024,height=768,scrollbars=yes,resizable=yes');
	best_sum.focus();
}

function time_sum(weekdate){
	var time_sum = window.open('/admin/offshop/timesellsum.asp?weekdate='+weekdate+'&<%=parameter%>','time_sum','width=1024,height=768,scrollbars=yes,resizable=yes');
	time_sum.focus();
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" onchange='frmsubmit();' <% if oldlist="on" then response.write "checked" %>>3������
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
				* ���屸�� : <% Call DrawShopDivCombo("offgubun",offgubun) %>
				&nbsp;&nbsp;
				* ����:<% drawweekday_select "weekdate" , weekdate ," onchange='frmsubmit();'" %>
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
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oreport.FresultCount %></b>
		�� 10�� ���� �˻��˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>�Ѹ���</td>
	<td>�Ѹ���Ǽ�<br>(�ֹ��Ǽ�)</td>
	<td>�ϼ�</td>
	<td>��ո���</td>
	<td>��ձ��ŰǼ�</td>
	<td>��հ��ܰ�</td>
	<td>��ո���������</td>
	<td>���</td>
</tr>
<% if oreport.FresultCount>0 then %>
<% for i=0 to oreport.FResultCount -1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="silver" onmouseout=this.style.background="#FFFFFF">
	<td><%= oreport.FItemList(i).GetDpartName %></td>
	<td align="right" bgcolor="#E6B9B8"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %></td>
	<td><%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %></td>
	<td><%= oreport.FItemList(i).FDpartCount %></td>
	<td align="right">
		<% if oreport.FItemList(i).FDpartCount<>0 then %>
			<% avgsell = CLng(oreport.FItemList(i).Fselltotal/oreport.FItemList(i).FDpartCount) %>
			<%= FormatNumber(avgsell,0) %>
		<% end if %>
	</td>
	<td>
		<% if oreport.FItemList(i).FDpartCount<>0 then %>
			<%= FormatNumber(CLng(oreport.FItemList(i).Fsellcnt/oreport.FItemList(i).FDpartCount),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if oreport.FItemList(i).Fsellcnt<>0 then %>
			<%= FormatNumber(CLng(oreport.FItemList(i).Fselltotal/oreport.FItemList(i).Fsellcnt),0) %>
		<% end if %>
	</td>
	<td>
		<% if avgselltotal<>0 then %>
			<%= CLng(avgsell/avgselltotal*100*100)/100 %> %
		<% end if  %>
	</td>
	<td width=350>
		<input type="button" class="button" value="ī�װ���" onclick="category_sum('<%= oreport.FItemList(i).Fdpart %>');">
		<input type="button" class="button" value="����Ʈ��" onclick="best_sum('<%= oreport.FItemList(i).Fdpart %>');">
		<input type="button" class="button" value="�ð��뺰��" onclick="time_sum('<%= oreport.FItemList(i).Fdpart %>');">
	</td>		
</tr>   
<% 
next
%>
<tr bgcolor="#FFFFFF" height="20">
	<td colspan="9"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="silver" onmouseout=this.style.background="#FFFFFF">
	<td align="center">����</td>
	<td align="right" ><%= FormatNumber(selltotal_jj,0) %></td>
	<td align="center"><%= FormatNumber(sellcnt_jj,0) %></td>
	<td align="center"><%= DpartCount_jj %></td>
	<td align="right">
		<% if DpartCount_jj<>0 then %>
			<%= FormatNumber(CLng(selltotal_jj/DpartCount_jj),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if DpartCount_jj<>0 then %>
			<%= FormatNumber(CLng(sellcnt_jj/DpartCount_jj),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if sellcnt_jj<>0 then %>
			<%= FormatNumber(CLng(selltotal_jj/sellcnt_jj),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if selltotal<>0 then %>
			<%= CLng(selltotal_jj/selltotal*100*100)/100 %> %
		<% end if  %>
	</td>
	<td></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="silver" onmouseout=this.style.background="#FFFFFF">
	<td align="center">�ָ�</td>
	<td align="right" ><%= FormatNumber(selltotal_jm,0) %></td>
	<td align="center"><%= FormatNumber(sellcnt_jm,0) %></td>
	<td align="center"><%= DpartCount_jm %></td>
	<td align="right">
		<% if DpartCount_jm<>0 then %>
			<%= FormatNumber(CLng(selltotal_jm/DpartCount_jm),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if DpartCount_jm<>0 then %>
			<%= FormatNumber(CLng(sellcnt_jm/DpartCount_jm),0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if sellcnt_jm<>0 then %>
			<%= FormatNumber(CLng(selltotal_jm/sellcnt_jm),0) %>
		<% end if %>
	</td>
	<td align="center">
		<% if selltotal<>0 then %>
			<%= CLng(selltotal_jm/selltotal*100*100)/100 %> %
		<% end if  %>
	</td>
	<td></td>
</tr>	
<%
else 
%>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>

<%
set oreport= Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->