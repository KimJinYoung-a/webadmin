<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �귣���԰����Ǹ���
' History : 2011.06.21 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offshop_reportcls.asp"-->

<%
dim page , yyyymm ,osum ,shopid , yyyy1 , mm1 ,i ,commcd, inc3pl
dim totipgosum ,tottotsellsum , tottotipgocnt ,tottotsellcnt ,totremainSum ,totstsum ,totstno
	yyyy1    = requestCheckVar(request("yyyy1"),4)
	mm1    = requestCheckVar(request("mm1"),2)
	shopid    = requestCheckVar(request("shopid"),32)
	commcd    = requestCheckVar(request("commcd"),10)
	page    = requestCheckVar(request("page"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if page="" then page="1"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
				
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

if shopid = "" then shopid = "streetshop011"

set osum = new COffshopReport
	osum.FPageSize = 1000
	osum.FCurrPage = page
	osum.FRectShopID = shopid
	osum.frectcommcd = commcd
	osum.FRectyyyymm = yyyy1 & "-" & Format00(2,mm1)
	osum.FRectInc3pl = inc3pl	
	osum.getbrandipgomaechul
%>

<script language='javascript'>

	function ReSearch(page){
		frm.page.value=page;
		frm.submit();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ :
		<% DrawYMBox yyyy1,mm1 %>
		&nbsp;&nbsp;	
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if not C_IS_OWN_SHOP and shopid <> "" then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShopNot000 "shopid",shopid %>
			<% end if %>
		<% else %>
			* ���� : <% drawSelectBoxOffShopNot000 "shopid",shopid %>
		<% end if %>	    
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ���Ա��� : <% drawSelectBoxOFFJungsanCommCD "commcd",commcd %>
		&nbsp;&nbsp;
		<b>* ����ó����</b>
		<% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= osum.ftotalcount %></b> ���� 1000�� ���� �˻� �˴ϴ�
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����</td>
	<td>�귣��</td>
	<td>���԰��</td>
	<td>���Ǹž�</td>
	<td>��������</td>
	<td>�԰����</td>
	<td>�Ǹż���</td>
	<td>����������</td>
	<td>�Ǹ���</td>
	<td>�Ǹž�-�԰��</td>
	<td>���Ա���</td>
</tr>
<% if osum.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
<% 
for i=0 to osum.FResultCount - 1

totipgosum = totipgosum + osum.FItemList(i).fipgosum
tottotsellsum = tottotsellsum + osum.FItemList(i).ftotsellsum
tottotipgocnt = tottotipgocnt + osum.FItemList(i).ftotipgocnt
tottotsellcnt = tottotsellcnt + osum.FItemList(i).ftotsellcnt
totremainSum = totremainSum + osum.FItemList(i).fremainSum
totstsum = totstsum + osum.FItemList(i).fstsum
totstno = totstno + osum.FItemList(i).fstno
%>
<tr bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF"; align="center">
	<td><%= osum.FItemList(i).fshopid %></td>
	<td><%= osum.FItemList(i).fmakerid %></td>
	<td><%= FormatNumber(osum.FItemList(i).fipgosum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotsellsum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fstsum,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotipgocnt,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).ftotsellcnt,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fstno,0) %></td>
	<td><%= FormatNumber(osum.FItemList(i).fpro,0) %>%</td>
	<td><%= FormatNumber(osum.FItemList(i).fremainSum,0) %></td>
	<td><%= osum.FItemList(i).fcomm_name %> (<%= osum.FItemList(i).fcomm_cd %>)</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=2></td>
	<td><%= FormatNumber(totipgosum,0) %></td>
	<td><%= FormatNumber(tottotsellsum,0) %></td>
	<td><%= FormatNumber(totstsum,0) %></td>
	<td><%= FormatNumber(tottotipgocnt,0) %></td>
	<td><%= FormatNumber(tottotsellcnt,0) %></td>
	<td><%= FormatNumber(totstno,0) %></td>
	<td></td>
	<td><%= FormatNumber(totremainSum,0) %></td>
	<td></td>	
</tr>
<% end if %>
</table>

<%
	set osum = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
