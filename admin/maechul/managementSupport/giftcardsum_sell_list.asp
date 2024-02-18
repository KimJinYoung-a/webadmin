<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ʈ ī�� ���� ���
' History : 2012.11.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcardsum_cls.asp" -->

<%
Dim i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate, csell, csellcancel, accountdiv, onoffgubun
dim subtotalPrice
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	accountdiv     = request("accountdiv")
	onoffgubun     = request("onoffgubun")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = Cstr(day( dateadd("m",-1,date()) ))	
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2 +1)

Set csell = New cgiftcardsum_list
	csell.FRectStartdate = fromDate
	csell.FRectEndDate = toDate
	csell.FPageSize = 1000
	csell.FCurrPage	= 1
	csell.frectaccountdiv = accountdiv
	csell.FRectonoffgubun = onoffgubun	
	csell.fgiftcardsum_sell_list()

Set csellcancel = New cgiftcardsum_list
	csellcancel.FRectStartdate = fromDate
	csellcancel.FRectEndDate = toDate
	csellcancel.FPageSize = 1000
	csellcancel.FCurrPage	= 1
	csellcancel.frectaccountdiv = accountdiv
	csellcancel.FRectonoffgubun = onoffgubun
	csellcancel.frectcancelyn="Y"
	csellcancel.fgiftcardsum_sell_list()
	
subtotalPrice = 0
%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				�Ⱓ : <% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
				��뱸�� : <% drawonoffgubun "onoffgubun",onoffgubun," onchange='javascript:searchSubmit();'" %>
				�ǸŰ������� : <% drawgiftcardaccountdiv "accountdiv",accountdiv," onchange='searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>	
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�� �̴Ͻý�(�Ǹų���)
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= csell.FresultCount %></b> �� �� 1000�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>����Ʈī��<br>�ֹ���ȣ</td>
    <td>��¥</td>
    <td>�ݾ�</td>    
    <td>������<Br>���̵�</td>
    <td>������<Br>�̸�</td>
    <td>�Ǹ�<Br>��������</td>
    <td>��������</td>
    <td>��ҿ���</td>        
</tr>

<% if csell.FresultCount > 0 then %>
<%
For i = 0 To csell.FresultCount -1

subtotalPrice = subtotalPrice + csell.fitemlist(i).fsubtotalPrice
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= csell.fitemlist(i).fgiftOrderSerial %>
	</td>
	<td>
		<%= left(csell.fitemlist(i).fYYYYMMdd,10) %>
	</td>
	<td>
		<%= FormatNumber(csell.fitemlist(i).fsubtotalPrice,0) %>
	</td>
	<td>
		<%= csell.fitemlist(i).fuserid %>
	</td>	
	<td>
		<%= csell.fitemlist(i).fbuyname %>
	</td>
	<td>
		<%= csell.fitemlist(i).faccountname %>
	</td>	
	<td>
		<%= IpkumDivName(csell.fitemlist(i).fipkumdiv) %>
	</td>
	<td>
		<%= csell.fitemlist(i).fcancelyn %>
	</td>
</tr>	
<% next %>

<%
For i = 0 To csellcancel.FresultCount -1

'if csellcancel.fitemlist(i).fcancelyn="N" or (left(csellcancel.fitemlist(i).fYYYYMMdd,10)=left(csellcancel.fitemlist(i).fcanceldate,10) and csellcancel.fitemlist(i).fcancelyn="Y") then
	subtotalPrice = subtotalPrice + csellcancel.fitemlist(i).fsubtotalPrice
'end if
%>
<tr bgcolor="#c1c1c1" align="center" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='c1c1c1';>
	<td>
		<%= csellcancel.fitemlist(i).fgiftOrderSerial %>
	</td>
	<td>
		<%= left(csellcancel.fitemlist(i).fYYYYMMdd,10) %>
	</td>
	<td>
		<%= FormatNumber(csellcancel.fitemlist(i).fsubtotalPrice,0) %>
	</td>
	<td>
		<%= csellcancel.fitemlist(i).fuserid %>
	</td>	
	<td>
		<%= csellcancel.fitemlist(i).fbuyname %>
	</td>
	<td>
		<%= csellcancel.fitemlist(i).faccountname %>
	</td>	
	<td>
		<%= IpkumDivName(csellcancel.fitemlist(i).fipkumdiv) %>
	</td>
	<td>
		<%= csellcancel.fitemlist(i).fcancelyn %>
		
		<% if csellcancel.fitemlist(i).fcancelyn="Y" and csellcancel.fitemlist(i).fcanceldate<>"" then %>
			<BR><%= csellcancel.fitemlist(i).fcanceldate %>
		<% end if %>
	</td>
</tr>	
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=2>
		 �հ�
	</td>
	<td>
		<%= FormatNumber(subtotalPrice,0) %>
	</td>
		
	<td colspan=5></td>
</tr>	

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<% 
Set csell = Nothing
set csellcancel = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->