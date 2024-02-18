<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_confirmitemcls.asp"-->
<!-- #include virtual="/admin/offshop/shopitemmodi_top.asp"-->
<%
dim designer
dim acttype, ckonlyusing, ckonlyoff, imageview, pricediff
designer 	= requestCheckVar(request("designer"),32)
acttype 	= requestCheckVar(request("acttype"),10)
ckonlyusing	= requestCheckVar(request("ckonlyusing"),2)
ckonlyoff	= requestCheckVar(request("ckonlyoff"),2)
imageview	= requestCheckVar(request("imageview"),2)
pricediff	= requestCheckVar(request("pricediff"),2)

dim oOffContractInfo
set oOffContractInfo = new COffContractInfo
	oOffContractInfo.FRectDesignerID = designer

	if designer<>"" then
		oOffContractInfo.GetPartnerOffContractInfo
	end if

dim i
%>
<table width="98%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		��ü:<% drawSelectBoxDesignerOffShopContract "designer",designer  %>

		<br>
		�ɼ�:
		<input type="checkbox" name="ckonlyusing" value="on" <% if ckonlyusing="on" then response.write "checked" %> >������λ�ǰ��
		<!--
		&nbsp;&nbsp;&nbsp;
		<input type="checkbox" name="ckonlyoff" value="on" <% if ckonlyoff="on" then response.write "checked" %> >������������
		&nbsp;&nbsp;<input type="checkbox" name="pricediff" value="on" <% if pricediff="on" then response.write "checked" %> >���ݻ��̸� ����
		-->
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<center><br><br>�������Դϴ�.</center>
<% if oOffContractInfo.FResultCount>0 then %>
<table width="98%" border="0" cellspacing="0" bgcolor="#CCCCCC" class="a">
	<tr >
	<td width="110" bgcolor="#FFDDDD" >��������(������)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=2 cellpadding=0 class=a>
		<tr>
			<td width=80><b>��������ǥ</b></td>
			<td ><b><%= oOffContractInfo.GetSpecialChargeDivName("streetshop000") %></b></td>
			<td width=80><b><%= oOffContractInfo.GetSpecialDefaultMargin("streetshop000") %> %</b></td>
		</tr>
		<% for i=0 to oOffContractInfo.FResultCount-1 %>
		<% if (oOffContractInfo.FItemList(i).Fshopdiv<>"3") and (oOffContractInfo.FItemList(i).Fshopid<>"streetshop000") then %>
		<tr>
			<td width=100><%= oOffContractInfo.FItemList(i).Fshopname %></td>
			<td ><%= oOffContractInfo.FItemList(i).GetChargeDivName() %></td>
			<td width=80><%= oOffContractInfo.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
	<td width="110" bgcolor="#FFDDDD" >��������(������)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=2 cellpadding=0 class=a>
		<tr>
			<td width=100><b>����������ǥ</b></td>
			<td ><b><%= oOffContractInfo.GetSpecialChargeDivName("streetshop800") %></b></td>
			<td width=80><b><%= oOffContractInfo.GetSpecialDefaultMargin("streetshop800") %> %</b></td>
		</tr>
		<% for i=0 to oOffContractInfo.FResultCount-1 %>
		<% if (oOffContractInfo.FItemList(i).Fshopdiv="3") and (oOffContractInfo.FItemList(i).Fshopid<>"streetshop800") then %>
		<tr>
			<td ><%= oOffContractInfo.FItemList(i).Fshopname %></td>
			<td ><%= oOffContractInfo.FItemList(i).GetChargeDivName() %></td>
			<td><%= oOffContractInfo.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
</table>
<% end if %>
<%
set oOffContractInfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->