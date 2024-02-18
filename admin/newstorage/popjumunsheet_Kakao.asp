<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ֹ�������
' History : �̻󱸻���
'			2018.09.13 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim webImgUrl
IF application("Svr_Info")="Dev" THEN
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'���̹���
else
	webImgUrl		= "/webimage"									'���̹���
end if

dim idx, itype, isFixed, i
	idx = requestCheckVar(getNumeric(request("idx")),10)
	itype = requestCheckVar(request("itype"),10)

dim oordersheetmaster, oordersheet
set oordersheetmaster = new COrderSheet
	oordersheetmaster.FRectIdx = idx
	oordersheetmaster.GetOneOrderSheetMaster

if oordersheetmaster.FtotalCount < 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�ش�Ǵ� �ֹ����� �����ϴ�.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if

isFixed = oordersheetmaster.FOneItem.IsFixed

set oordersheet = new COrderSheet
	oordersheet.FrectisFixed = isFixed
	oordersheet.FRectIdx = idx
	oordersheet.GetOrderSheetDetail

dim obrand
set obrand = new CBrandShopInfoItem
	obrand.FRectChargeId = oordersheetmaster.FOneItem.Ftargetid
	obrand.GetBrandShopInFo

dim scheduleorexedate
	scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")

dim ttlsellcash, ttlbuycash, ttlcount
ttlsellcash = 0
ttlbuycash  = 0
ttlcount    = 0

%>
<h1 align="center">�ֹ��� : <%= scheduleorexedate %></h2>

<h1 align="center">�ٹ�����(<%= oordersheetmaster.FOneItem.FBaljuCode %>)</h2>

<table align="center" width="500">
	<% for i=0 to oordersheet.FResultCount -1 %>
	<tr>
		<td align="center" colspan="2">
			<% if (oordersheet.FItemList(i).Fbasicimage <> "") then %>
			<img src="<%= oordersheet.FItemList(i).Fbasicimage %>">
			<% elseif (oordersheet.FItemList(i).Foffimgmain <> "") then %>
			<img src="<%= oordersheet.FItemList(i).Foffimgmain %>" width="450">
			<% end if %>
		</td>
	</tr>
	<tr>
		<td width="50"></td>
		<td>
			<h2>��ǰ�ڵ� : <%= oordersheet.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %>-<%= oordersheet.FItemList(i).FItemOption %></h2>
			<h2>��ǰ�� : <%= oordersheet.FItemList(i).FItemName %></h2>
			<h2>�ɼǸ� : <%= oordersheet.FItemList(i).FItemOptionName %></h2>
			<h2>���� : <%= oordersheet.FItemList(i).Fbaljuitemno %></h2>
		</td>
	</tr>
	<% next %>
</table>
<%
set obrand = Nothing
set oordersheetmaster = Nothing
set oordersheet = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
