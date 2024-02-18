<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드 리스트
' History : 2012.08.21 서동석 생성
'			2012.08.22 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<%

dim masteridx, itemcount

masteridx = requestCheckVar(request("idx"), 32)

dim oPickupMaster
set oPickupMaster = new CAGVItems
	oPickupMaster.FRectMasterIdx = masteridx
	oPickupMaster.GetPickupMasterOne

dim oPickupDetail
set oPickupDetail = new CAGVItems
	oPickupDetail.FRectMasterIdx = masteridx
	oPickupDetail.FPageSize = 20000
	oPickupDetail.GetPickupAgvStockoutList

dim companyurl
IF application("Svr_Info")="Dev" THEN
	companyurl = "http://testcomp.10x10.co.kr"
else
	companyurl = "http://company.10x10.co.kr"
end if

dim i

%>
<STYLE TYPE="text/css">

<!-- .break {page-break-before: always;} -->

</STYLE>
<script language="javascript1.2" type="text/javascript" src="/js/barcode.js"></script>

<table width="100%" height=40 border="0" cellpadding="2" cellspacing="1">
<tr>
	<td width="150" class="a">
		<b>AGV결품내역</b>
	</td>
	<td width="250" class="a">
		IDX : <%= masteridx %>
	</td>
	<td width="400" class="a">
		작업지시코드 : <%= oPickupMaster.FOneItem.FrequestNo %>
	</td>
	<td align="center" class="a">
		<table border=0 cellspacing="0" cellpadding="2" class="a">
		<tr>
			<td align="center">

			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
if oPickupDetail.FResultCount>0 then
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<%
for i=0 to oPickupDetail.FResultCount-1
%>
<%
if itemcount >= 16 then
    itemcount = 0
%>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height=20 bgcolor="#FFFFFF">
	<td >입고완료시각:</td>
	<td width=250></td>
	<td >입고자:</td>
	<td width=250></td>
</tr>
</table>
<div class="break"></div>
<table width="100%" height=30 border="0" cellpadding="2" cellspacing="1">
<tr>
	<td width="150" class="a">
		<b>AGV결품내역</b>
	</td>
	<td width="250" class="a">
		IDX : <%= masteridx %>
	</td>
	<td width="400" class="a">
		작업지시코드 : <%= oPickupMaster.FOneItem.FrequestNo %>
	</td>
	<td align="center" class="a">
		<table border=0 cellspacing="0" cellpadding="2" class="a">
		<tr>
			<td align="center">

			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<% end if %>

<tr bgcolor="#FFFFFF">
	<td width="60" align="center">상품R</td>
	<td width="100" align="center">브랜드</td>
	<td width="20" align="center">gbn</td>
	<td width="60" align="center">itemid</td>
	<td width="40" align="center">option</td>
	<td width="30" align="center">수량</td>
	<td align="center" colspan="2">상품명</td>
	<td align="center">옵션</td>
	<td width="30" align="center">재고</td>
    <td width="30" align="center">피킹</td>
</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" RowSpan="2"><%= oPickupDetail.FItemList(i).GetItemRackCode %></td>
		<td align="left"><%= oPickupDetail.FItemList(i).Fmakerid %></td><% ''브랜드명은 이걸로 통일할것(발주서,인덱스바코드 등등), skyer9, 2018-03-13 %>
		<td align="center"><%= oPickupDetail.FItemList(i).FItemGubun %></td>
		<td align="center"><%= oPickupDetail.FItemList(i).FItemID %></td>
		<td align="center"><%= oPickupDetail.FItemList(i).FItemOption %></td>
		<td align="center" RowSpan="2">
			<font size=2><b><%= oPickupDetail.FItemList(i).Fshortageno %></b></font>
		</td>
		<td align="left" ColSpan="2" style="padding-left: 5px; padding-right: 5px;"><%= oPickupDetail.FItemList(i).FItemName %></td>
		<td align="center" RowSpan="2" style="padding-left: 5px; padding-right: 5px;"><%= oPickupDetail.FItemList(i).FItemOptionName %></td>
		<td align="center" rowspan="2"><%= oPickupDetail.FItemList(i).Frealstock %></td>
        <td align="center" rowspan="2"><%= oPickupDetail.FItemList(i).Fpickupno %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center" colspan="4">
		</td>
	    <td align="center" style="padding-left: 5px; padding-right: 5px;">
			<img src="<%=companyurl%>/barcode/barcode.asp?image=3&type=20&height=15&barwidth=1&TextAlign=2&data=<%= BF_MakeTenBarcode(oPickupDetail.FItemList(i).FItemGubun, oPickupDetail.FItemList(i).Fitemid, oPickupDetail.FItemList(i).Fitemoption) %>"><!--<br>
			<%= BF_MakeTenBarcode(oPickupDetail.FItemList(i).FItemGubun, oPickupDetail.FItemList(i).Fitemid, oPickupDetail.FItemList(i).Fitemoption) %>-->
		</td>
	    <td align="center" style="padding-left: 5px; padding-right: 5px;">
			<% if Not IsNull(oPickupDetail.FItemList(i).Fpublicbarcode) then %>
			<% if (oPickupDetail.FItemList(i).Fpublicbarcode <> "") then %>
			<% if Len(oPickupDetail.FItemList(i).Fpublicbarcode) > 4 then %>
			<%= Left(oPickupDetail.FItemList(i).Fpublicbarcode, (Len(oPickupDetail.FItemList(i).Fpublicbarcode) - 4)) %><b><%= Right(oPickupDetail.FItemList(i).Fpublicbarcode, 4) %></b>
			<% else %>
			<b><%= oPickupDetail.FItemList(i).Fpublicbarcode %></b>
			<% end if %>
			<% end if %>
			<% end if %>
		</td>
	</tr>
<%
itemcount = itemcount + 1

next
%>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height=20 bgcolor="#FFFFFF">
	<td >입고완료시각:</td>
	<td width=250></td>
	<td >입고자:</td>
	<td width=250></td>
</tr>
</table>

<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
