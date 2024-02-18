<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �𺰱�������
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone,i,page , parameter , shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,fromDate ,toDate
dim designer , sellgubun ,cdl ,cdm ,cds ,datefg
	designer = RequestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	sellgubun = requestCheckVar(request("sellgubun"),1)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
	datefg = requestCheckVar(request("datefg"),10)

	if datefg = "" then datefg = "maechul"			
	if sellgubun = "" then sellgubun = "S"
	if (yyyy1="") then yyyy1 = Cstr(Year(now()))
	if (mm1="") then mm1 = Cstr(Month(now()))
	if (dd1="") then dd1 = Cstr(day(now()))
	if (yyyy2="") then yyyy2 = Cstr(Year(now()))
	if (mm2="") then mm2 = Cstr(Month(now()))
	if (dd2="") then dd2 = Cstr(day(now()))
				
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
	
	if page = "" then page = 1

'����/������
if (C_IS_SHOP) then
	
	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if

set ozone = new czone_list
	ozone.FPageSize = 500
	ozone.FCurrPage = page
	ozone.frectshopid = shopid
	ozone.FRectStartDay = fromDate
	ozone.FRectEndDay = toDate
	ozone.FRectmakerid = designer	
	ozone.FRectCDL = cdl
	ozone.FRectCDM = cdm
	ozone.FRectCDN = cds	
	ozone.frectdatefg = datefg
	ozone.frectsellgubun = sellgubun

	if shopid <> "" then
		ozone.Getoffshopzonesum
	end if
	
	if shopid = "" then response.write "<script>alert('������ �������ּ���');</script>"
		
parameter = "shopid="&shopid&"&sellgubun="&sellgubun&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&menupos="&menupos&"&designer="&designer&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds
%>

<script language="javascript">

//�׷켳��
function zone_groupreg(){
	var zone_groupreg = window.open('/admin/offshop/zone/zone_common.asp?menupos=<%=menupos%>','zone_groupreg','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_groupreg.focus();
}

//���屸������
function zone_reg(){
	var zone_reg = window.open('/admin/offshop/zone/zone.asp?menupos=<%=menupos%>','zone_reg','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_reg.focus();
}

//��������ǰ���
function zone_item(){

	if (frm.shopid.value==''){
		alert('������ ������ �ּ���');
		return;
	}
		
	var zone_item = window.open('zone_item.asp?menupos=<%=menupos%>&shopid=<%=shopid%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>&datefg=<%=datefg%>','zone_item','width=1024,height=768,scrollbars=yes,resizable=yes');
	zone_item.focus();
}

//�󼼸���
function item_detail(idx,searchtype){

	var item_detail = window.open('zone_sum_detail.asp?idx='+idx+'&searchtype='+searchtype+'&<%=parameter%>','item_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
	item_detail.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		����:<% drawSelectBoxOffShop "shopid",shopid %>
		������� :
		<% drawmaechul_datefg "datefg" ,datefg ,""%> 
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<Br><!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�귣�� : <% drawSelectBoxDesignerwithName "designer",designer %>
		<input type="radio" name="sellgubun" value="S" <% if sellgubun="S" then response.write " checked" %>>������������
		<input type="radio" name="sellgubun" value="N" <% if sellgubun="N" then response.write " checked" %>>�����ϳ�������
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
		�� [����] "�׷켳��" ���� ���忡 �����ϴ� �׷��� ���� ������ , "����������" ���� �����庰�� ������ ���� �Ͻ���
		<br>"����ǰ��������" ���� ���������� ��ǰ�� �����ž� ������ ����˴ϴ�.
		<%
		'/������������
		if sellgubun="S" then
		%>
			<br>[����] �������� ������ �Ǹ�, �ش� ��ǰ�� ��ϵ� ���� ������ ���� �Ǹ�, �̸� �������� ��谡 ���� �˴ϴ�
			<br>�׷��Ƿ� ������ ��ǰ�� ������� ������, ��谡 ���� �ʽ��ϴ�.
		<%
		'/�����ϳ�������
		else
		%>
			<br>[����] ���� ������ ��ǰ ��Ͽ� ��ϵǾ��� �ִ� ���� �Դϴ�. ��ǰ�� ���� ������ ���� ���� �ʽ��ϴ�.
		<% end if %>
	</td>
	<td align="right">
		<input type="button" class="button" value="�׷켳��" onclick="zone_groupreg();">
		<input type="button" class="button" value="���屸������" onclick="zone_reg();">
		<input type="button" class="button" value="��������ǰ���" onclick="zone_item();">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ozone.FTotalCount %></b>
		�� 500�� ���� �˻��˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">�׷�</td>
	<td align="center">�Ŵ�Ÿ��</td>
	<td align="center">�󼼱�����</td>
	<td>��<br>�Ǹż���</td>
	<td>��<br>�Ǹ����</td>
	<td>������</td>
	<td>UNIT�� ����<br>UNIT</td>
	<td>���</td>
</tr>
<% if ozone.FTotalCount>0 then %>
<% for i=0 to ozone.FTotalCount-1 %>
<% if ozone.FItemList(i).fzonename <> "" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= getOffShopzonegroup(ozone.FItemList(i).fzonegroup) %>
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= getOffShopracktype(ozone.FItemList(i).fracktype) %>
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).fzonename = "" then %>
			-
		<% else %>
			<%= ozone.FItemList(i).fzonename %>
		<% end if %>
	</td>	
	<td>
		<%= ozone.FItemList(i).fitemcnt %>
	</td>
	<td>
		<%= FormatNumber(ozone.FItemList(i).fsellsum,0) %>
	</td>
	<td>
		<% if ozone.FSumTotal<>0 then %>
			<%= Clng( ((ozone.FItemList(i).fsellsum / ozone.FSumTotal) * 10000)) / 100 %> %
		<% end if %>
	</td>
	<td>
		<% if ozone.FItemList(i).funitvalue <> "" then %>
			<%= FormatNumber(ozone.FItemList(i).funitvalue,0) %>
		<% else %>
			-
		<% end if %>
		<% if ozone.FItemList(i).funit <> "" then %>
			<br><%= FormatNumber(ozone.FItemList(i).funit,0) %>
		<% else %>
			<br>-
		<% end if %>
	</td>
	<td width=200>
		<input type="button" class="button" value="��ǰ��" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','I');">
		<input type="button" class="button" value="ī�װ���" onclick="item_detail('<%= ozone.FItemList(i).fidx %>','C');">
	</td>	
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan=3>
		�հ�
	</td>		
	<td align="center">
		<%= FormatNumber(ozone.FCountTotal,0) %>
	</td>
		
	<td align="center">
		<%= FormatNumber(ozone.FSumTotal,0) %>
	</td>
	<td align="center" colspan=4></td>	
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->