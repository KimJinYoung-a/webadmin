<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  온라인 출고지시
' History : 2009.03.28 서동석 생성
'			2023.07.11 한용민 수정(29cm 삭제)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljuipgocls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljucls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<%
dim baljukey ,section ,maxpage ,obaljupage ,oblaju, companyurl, sitebaljukey
dim research ,ckimage ,i , prepageno, maxsection
	baljukey = requestCheckVar(getNumeric(request("baljukey")),10)
	sitebaljukey = requestCheckVar(getNumeric(request("sitebaljukey")),10)
	section = requestCheckVar(request("section"),10)
	ckimage = requestCheckVar(request("ckimage"),2)
	research = requestCheckVar(request("research"),2)

if section="" then section=1

dim obaljukey
set obaljukey = new CBaljuIpgo
if (sitebaljukey <> "") and (baljukey = "") then
	obaljukey.FRectSiteBaljuKey = sitebaljukey
	baljukey = obaljukey.GetBaljuKeyWithSiteBaljuKey
end if

set oblaju = new CBalju
	oblaju.FRectBaljuKey = baljukey
	oblaju.GetOneBaljuMaster

set obaljupage = new CBaljuIpgo
	obaljupage.FRectBaljuKey = baljukey
	''page 작성 count(pageno=0)>0 인경우
	obaljupage.FRectMakePageSize = 16
	obaljupage.FRectPreMakeItemNo = 5
	obaljupage.MakeBaljuPage

	maxpage = obaljupage.GetMaxPage
	obaljupage.FPageNoStart = (section-1)*10+1
	obaljupage.FPageNoEnd = (section-1)*10+10
	obaljupage.GetBaljuIpgoByPageRect

ckimage="on"

maxsection =  CInt(maxpage\10)
if  (maxpage\10)<>(maxpage/10) then
	maxsection = maxsection +1
end if

IF application("Svr_Info")="Dev" THEN
	companyurl = "http://testcomp.10x10.co.kr"
else
	companyurl = "http://company.10x10.co.kr"
end if
%>

<STYLE TYPE="text/css">

<!-- .break {page-break-before: always;} -->

</STYLE>
<script language="javascript1.2" type="text/javascript" src="/js/barcode.js"></script>

<table width="100%" height=40 border="0" cellpadding="2" cellspacing="1">
<tr>
	<td width=110>
		<form name="frm" method="get" style="margin:0px;" >
		<input type="hidden" name="research" value="on">
		<input type="hidden" name="baljukey" value="<%= baljukey %>">
		<table border="0" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td align=right>
				Section : <input type=text name="section" value="<%= section %>" size="2" maxlength=3 >/<%= maxsection %>
			</td>
		</tr>
		</table>
		</form>
	</td>
	<td width="110" class="a">
		IDX : <%= baljukey %><br>
		출고지시코드 : <b><%= oblaju.FOneBaljumaster.FSiteBaljuID %></b>
	</td>
	<td width="400" class="a">
		출고지시일 : <b><%= Left(oblaju.FOneBaljumaster.FBaljudate,10) %></b> 차수 : <b><%= oblaju.FOneBaljumaster.Fdifferencekey %></b> 그룹 : <b><%= oblaju.FOneBaljumaster.Fworkgroup %></b>
		&nbsp;
		<% if obaljupage.FResultCount>0 then %>
		Page <b><%= obaljupage.FItemList(0).FPageNo %></b>/<%= maxpage  %>
		<% end if %>
	</td>
	<td align="center" class="a">
		<table border=0 cellspacing="0" cellpadding="2" class="a">
		<tr>
			<td align="center">
				<% if obaljupage.FResultCount>0 then %>
					<img src="<%=companyurl%>/barcode/barcode.asp?image=3&type=20&height=30&barwidth=1&caption=<%= format00(6,baljukey) %><%= format00(6,obaljupage.FItemList(0).FPageNo) %>&data=<%= format00(6,baljukey) %><%= format00(6,obaljupage.FItemList(0).FPageNo) %>">
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
if obaljupage.FResultCount>0 then
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<%
for i=0 to obaljupage.FResultCount-1
%>
<% if (prepageno<>"") and (prepageno<>obaljupage.FItemList(i).FPageNo) then %>
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
	<td width=110>&nbsp;</td>
	<td width="400" class="a">
		출고지시코드 : <b><%= baljukey %></b>
		&nbsp;
		출고지시일 : <b><%= Left(oblaju.FOneBaljumaster.FBaljudate,10) %></b> 차수 : <b><%= oblaju.FOneBaljumaster.Fdifferencekey %></b> 그룹 : <b><%= oblaju.FOneBaljumaster.Fworkgroup %></b>
	</td>
	<td align="right" class="a">
		Page <b><%= obaljupage.FItemList(i).FPageNo %></b>/<%= maxpage  %>
	</td>
	<td align="center" class="a">
		<table border=0 cellspacing="0" cellpadding="2" class="a">
		<tr>
			<td align="center">
				<img src="<%=companyurl%>/barcode/barcode.asp?image=3&type=20&height=30&barwidth=1&caption=<%= format00(6,baljukey) %><%= format00(6,obaljupage.FItemList(i).FPageNo) %>&data=<%= format00(6,baljukey) %><%= format00(6,obaljupage.FItemList(i).FPageNo) %>">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
<% end if %>

<% if (prepageno<>obaljupage.FItemList(i).FPageNo) then %>
<tr bgcolor="#FFFFFF">
	<td width="32" align="center">상품R</td>
	<td width="50" align="center">이미지</td>
	<td width="100" align="center">브랜드</td>
	<td width="40" align="center">itemid</td>
	<td width="40" align="center">option</td>
	<td align="center">상품명</td>
	<td align="center">옵션</td>
	<td width="50" align="center">가격</td>
	<td width="30" align="center">수량</td>
	<td width="40" align="center">비고</td>
</tr>
<% end if %>
	<tr bgcolor="#FFFFFF">
	  <td align="center" RowSpan="2"><%= obaljupage.FItemList(i).GetItemRackCode %></td>
	  <% if ckimage="on" then %>
  		<td width="50" align="center" RowSpan="2"><img src="<%= obaljupage.FItemList(i).FimageSmall %>" width=50 height=50></td>
  	  <% end if %>
	  <td align="left" RowSpan="2"><%= obaljupage.FItemList(i).FBrandName %></td>
	  <td align="center"><%= obaljupage.FItemList(i).FItemID %></td>
	  <td align="center"><%= obaljupage.FItemList(i).FItemOption %></td>
	  <td align="left" RowSpan="2">&nbsp;&nbsp;&nbsp;&nbsp;<%= obaljupage.FItemList(i).FItemName %><br>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<img src="<%=companyurl%>/barcode/barcode.asp?image=3&type=20&height=15&barwidth=1&TextAlign=2&data=<%= BF_MakeTenBarcode(obaljupage.FItemList(i).FSiteSeq, obaljupage.FItemList(i).Fitemid, obaljupage.FItemList(i).Fitemoption) %>">
	  </td>
	  <td align="center" RowSpan="2"><%= obaljupage.FItemList(i).FItemOptionName %>&nbsp;</td>
	  <td align="right" RowSpan="2"><%= FormatNumber(obaljupage.FItemList(i).FOrgSellcash,0) %></td>
	  <td align="center" RowSpan="2"><%= obaljupage.FItemList(i).FBaljuNo %></td>
	  <td RowSpan="2">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="center"></td>
	    <td align="center"></td>
	</tr>
<%
prepageno=obaljupage.FItemList(i).FPageNo

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

<%
set oblaju = Nothing
set obaljupage = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
