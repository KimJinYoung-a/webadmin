<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문내역서
' History : 이상구생성
'			2018.09.13 한용민 생성
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
 	webImgUrl		= "http://testwebimage.10x10.co.kr"				'웹이미지
else
	webImgUrl		= "/webimage"									'웹이미지
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
	response.write "	alert('해당되는 주문건이 없습니다.');"
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
<h1 align="center">주문일 : <%= scheduleorexedate %></h2>

<h1 align="center">텐바이텐(<%= oordersheetmaster.FOneItem.FBaljuCode %>)</h2>

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
			<h2>상품코드 : <%= oordersheet.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %>-<%= oordersheet.FItemList(i).FItemOption %></h2>
			<h2>상품명 : <%= oordersheet.FItemList(i).FItemName %></h2>
			<h2>옵션명 : <%= oordersheet.FItemList(i).FItemOptionName %></h2>
			<h2>수량 : <%= oordersheet.FItemList(i).Fbaljuitemno %></h2>
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
