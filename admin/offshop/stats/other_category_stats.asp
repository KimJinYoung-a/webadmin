<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ���۾� ī�װ��� ��� 
'				���������� �� ī�װ����� �����ϰ� ���۾����� �ۼ��Ǵ� ����Դϴ�
' Hieditor : 2011.11.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/stats/other_category_stats_cls.asp"-->

<%
Dim othercate,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer ,datefg , othercdl ,totsellcnt , totsellsum , totsuplysum ,othercheck, inc3pl
	designer = RequestCheckVar(request("designer"),32)
	othercdl = RequestCheckVar(request("othercdl"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
if datefg = "" then datefg = "maechul"			
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)
		
if page = "" then page = 1
if othercdl = "" then othercdl = "070"
if othercdl = "toms001" then
	designer = othercdl
	othercheck = "ON"
end if
			
'����/������
if (C_IS_SHOP) then
	
	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if


set othercate = new cothercate_list
	othercate.FPageSize = 500
	othercate.FCurrPage = page
	othercate.frectshopid = shopid
	othercate.FRectStartDay = fromDate
	othercate.FRectEndDay = toDate
	othercate.FRectmakerid = designer	
	othercate.frectdatefg = datefg
	othercate.frectothercdl = othercdl
	othercate.FRectInc3pl = inc3pl
	
	if shopid <> "" then
		othercate.getother_category
	end if
	
	if shopid = "" then response.write "<script>alert('������ �������ּ���');</script>"

totsellcnt = 0
totsellsum = 0
totsuplysum = 0

parameter = "shopid="&shopid&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&designer="&designer&"&datefg="&datefg&"&othercheck="&othercheck&"&inc3pl="&inc3pl
%>

<script language="javascript">
	
function frmsubmit(){
	frm.submit();
}

//����Ʈ��ǰ
function best_detail(catecdm,othercdl){
	var best_detail = window.open('other_category_stats_best.asp?catecdm='+catecdm+'&othercdl='+othercdl+'&<%=parameter%>','best_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	best_detail.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ⱓ : <% drawmaechul_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
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
		* �귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;
		* ��ī�װ��� : <% other_category "othercdl",othercdl," onchange='frmsubmit();'" %>
        &nbsp;&nbsp;
        <b>* ����ó����</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� ���������� ���� ī�װ����� �����ϰ� ���۾����� �ۼ��Ǵ� ����Դϴ�. ������ ���Ͻø� �ý������� ��û�ϼ���.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= othercate.FTotalCount %></b>
		�� 500�� ���� �˻��˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��<br>ī�װ���</td>
	<td>��<br>ī�װ���</td>
	<td>�Ǹż�</td>
	<td>�����</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>���Ծ�</td>
	<% end if %>
	
	<td>���</td>
</tr>
<% if othercate.FTotalCount>0 then %>
<% 
for i=0 to othercate.FTotalCount-1 

totsellcnt = totsellcnt + othercate.FItemList(i).fsellcnt
totsellsum = totsellsum + othercate.FItemList(i).fsellsum
totsuplysum = totsuplysum + othercate.FItemList(i).fsuplysum
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>	
	<td align="left">
		<%= othercate.FItemList(i).fcdlcode_nm %>
	</td>
	<td align="left">
		<%= othercate.FItemList(i).fcdmcode_nm %>
	</td>
	<td>
		<%= FormatNumber(othercate.FItemList(i).fsellcnt,0) %>
	</td>
	<td bgcolor="#E6B9B8">
		<%= FormatNumber(othercate.FItemList(i).fsellsum,0) %>
	</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			<%= FormatNumber(othercate.FItemList(i).fsuplysum,0) %>
		</td>
	<% end if %>
	
	<td width=100>
		<input type="button" class="button" value="����Ʈ��ǰ" onclick="best_detail('<%= othercate.FItemList(i).fcatecdm %>','<%= othercate.FItemList(i).fcatecdl %>');">
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">	
	<td colspan=2>�հ�</td>
	<td>
		<%= FormatNumber(totsellcnt,0) %>
	</td>
	<td>
		<%= FormatNumber(totsellsum,0) %>
	</td>
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			<%= FormatNumber(totsuplysum,0) %>
		</td>	
	<% end if %>	
	<td></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set othercate = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->