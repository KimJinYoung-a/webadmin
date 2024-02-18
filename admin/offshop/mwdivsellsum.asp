<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2010.05.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/newoffshopsellcls.asp"-->
<%
const Maxlines = 10

dim totalpage, totalnum, q, i , yyyy1,mm1,dd1,yyyy2,mm2,dd2 , yyyymmdd1,yyymmdd2
dim ojumun ,fromDate,toDate,jnx,tmpStr,siteId ,searchId ,ck_nopoint 
dim TTLselltotal,TTLbuytotal,TTLsellcnt ,TTLminustotal,TTLminusbuytotal,TTLminuscount
dim shopid , offgubun , totprofit , totmagin, inc3pl
	ck_nopoint  = requestCheckVar(request("ck_nopoint")	,10)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	shopid = requestCheckVar(request("shopid"),32)
	offgubun = requestCheckVar(request("offgubun"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

	if (yyyy1="") then yyyy1 = Cstr(Year(now()))
	if (mm1="") then mm1 = Cstr(Month(now()))
	if (dd1="") then dd1 = Cstr(day(now()))
	if (yyyy2="") then yyyy2 = Cstr(Year(now()))
	if (mm2="") then mm2 = Cstr(Month(now()))
	if (dd2="") then dd2 = Cstr(day(now()))
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

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if	

set ojumun = new COffShopSell
	ojumun.FRectFromDate = fromDate
	ojumun.FRectToDate = toDate
	ojumun.FRectOffgubun = offgubun
	ojumun.FRectShopID = shopid
	ojumun.FRectInc3pl = inc3pl	
	ojumun.getmwdivsellsum()
%>

<script language='javascript'>

function Check(){
   document.frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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
		<input type="button" class="button_s" value="�˻�" onClick="Check();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ���� ���� : <% Call DrawShopDivCombo("offgubun",offgubun) %>
        &nbsp;&nbsp;
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="left">
		�ѱݾ�:&nbsp;<%= FormatNumber(ojumun.FMtotalmoney,0) %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		�ѰǼ�:&nbsp;<%= FormatNumber(ojumun.FMtotalsellcnt,0) %>
	</td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td width="50">���Ա���</td>
  	<td width="50">�Ǹ�<br>(+��)</td>

  	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
  		<td width="50">���԰�<br>(+��)</td>
  	<% end if %>

  	<td width="50">�Ǽ�<br>(+��)</td>
  	<td width="50">�Ǹ�<br>(-��)</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	  	<td width="50">���԰�<br>(-��)</td>
  	<% end if %>

  	<td width="50">�Ǽ�<br>(-��)</td>	

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
  		<td width="50">����</td>
  		<td width="50">������</td>
  	<% end if %>

  	<td>�׷���</td>  	
</tr>
<% if ojumun.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center"  >[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
<% 
for i=0 to ojumun.FResultCount - 1

TTLselltotal = TTLselltotal + ojumun.FItemList(i).Fselltotal
TTLbuytotal = TTLbuytotal + ojumun.FItemList(i).Fbuytotal
TTLsellcnt = TTLsellcnt + ojumun.FItemList(i).Fsellcnt

TTLminustotal = TTLminustotal + ojumun.FItemList(i).Fminustotal
TTLminusbuytotal = TTLminusbuytotal + ojumun.FItemList(i).Fminusbuytotal
TTLminuscount = TTLminuscount + ojumun.FItemList(i).Fminuscount
totmagin = totmagin + ojumun.FItemList(i).fmagin
totprofit = totprofit + ojumun.FItemList(i).fprofit
%>
<tr bgcolor="#FFFFFF" height=24>
	<td height="10">
      		<%= ojumun.FItemList(i).getcomm_cdname %>
  	</td>  	
	<td align="right">
		<% if Not (IsNull(ojumun.FItemList(i).Fselltotal)) then %>
			<%= FormatNumber(ojumun.FItemList(i).Fselltotal,0) %>
	  	<% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<% if Not (IsNull(ojumun.FItemList(i).Fbuytotal)) then %>
				<%= FormatNumber(ojumun.FItemList(i).Fbuytotal,0) %>
		  	<% end if %>
		</td>
  	<% end if %>

	<td align="right">
	  	<% if Not (IsNull(ojumun.FItemList(i).Fsellcnt)) then %>
	  		<%= ojumun.FItemList(i).Fsellcnt %>
	  	<% end if %>
	</td>
	<td align="right">
		<% if Not (IsNull(ojumun.FItemList(i).Fminustotal)) then %>
			<%= FormatNumber(ojumun.FItemList(i).Fminustotal,0) %>
	  	<% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<% if Not (IsNull(ojumun.FItemList(i).Fminusbuytotal)) then %>
				<%= FormatNumber(ojumun.FItemList(i).Fminusbuytotal,0) %>
		  	<% end if %>
		</td>
  	<% end if %>

	<td align="right">
	  	<% if Not (IsNull(ojumun.FItemList(i).Fminuscount)) then %>
	  		<%= ojumun.FItemList(i).Fminuscount %>
	  	<% end if %>
	</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<%= FormatNumber(ojumun.FItemList(i).fprofit,0) %>
		</td>
		<td align="right">
			<%= round(ojumun.FItemList(i).fmagin,1) %>%
		</td>
  	<% end if %>

  	<td align="left" height="35">
  		<% if ojumun.FItemList(i).Fselltotal <> 0 and ojumun.FItemList(i).Fselltotal <> "" and ojumun.maxt <> 0 and ojumun.maxt <> "" then %>
  			<img src="/images/dot1.gif" height="3" width="<%= CLng((ojumun.FItemList(i).Fselltotal/ojumun.maxt)*500) %>"><br>
  			<img src="/images/dot2.gif" height="3" width="<%= CLng((ojumun.FItemList(i).Fsellcnt/ojumun.maxc)*500) %>">
  		<% end if %>
  	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height=24>
    <td>�Ѱ�</td>
    <td align="right"><%= FormatNumber(TTLselltotal,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(TTLbuytotal,0) %></td>
  	<% end if %>

    <td align="right"><%= FormatNumber(TTLsellcnt,0) %></td>
    <td align="right"><%= FormatNumber(TTLminustotal,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(TTLminusbuytotal,0) %></td>
  	<% end if %>

    <td align="right"><%= FormatNumber(TTLminuscount,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td align="right"><%= FormatNumber(totprofit,0) %></td>
	    <td align="right">
	    	<% if totmagin <> 0 and totmagin <> "" and ojumun.FResultCount <> 0 then %>
	    		<%= round(totmagin / ojumun.FResultCount,1) %>%
	    	<% else %>
	    		0 %
	    	<% end if %>
	    </td>
  	<% end if %>

    <td></td>
</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
