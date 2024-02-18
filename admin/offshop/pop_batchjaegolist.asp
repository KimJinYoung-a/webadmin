<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
dim shopid, idx
	shopid = requestCheckVar(request("shopid"),32)
	idx = requestCheckVar(request("idx"),10)

dim oshoporder
set oshoporder = new CShopOrder
oshoporder.FRectShopID = shopid
oshoporder.FRectIdx = idx
oshoporder.FPageSize = 2000
oshoporder.GetShopOrderDetail

dim i

%>



<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	&nbsp;
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="40">NO</td>
      <td width="60">브랜드ID</td>
      <td width="40">구분</td>
      <td width="60">상품ID</td>
      <td width="60">옵션코드</td>
      <td>상품명</td>
      <td width="100">옵션</td>
      <td width="50">판매가</td>
      <td width="50">매입가</td>
      <td width="50">갯수</td>
    </tr>
    <% for i=0 to oshoporder.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
      <td><%= (i + 1) %></td>
      <td align="left"><%= oshoporder.FItemList(i).FMakerid %></td>
      <td><%= oshoporder.FItemList(i).Fitemgubun %></td>
      <td><%= oshoporder.FItemList(i).Fitemid %></td>
      <td><%= oshoporder.FItemList(i).Fitemoption %></td>
      <td align="left"><%= oshoporder.FItemList(i).Fitemname %></td>
      <td align="left"><%= oshoporder.FItemList(i).Fitemoptionname %></td>
      <td align="right"><%= FormatNumber(oshoporder.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(oshoporder.FItemList(i).Fsuplyprice,0) %></td>
      <td><%= oshoporder.FItemList(i).Fitemno %></td>
    </tr>
	<% next %>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set oshoporder = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->