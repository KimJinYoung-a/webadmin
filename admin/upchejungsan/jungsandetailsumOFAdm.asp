<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim idx, gubuncd, shopid
dim makerid, groupid
makerid = requestCheckVar(request("makerid"),32)
idx     = requestCheckvar(request("idx"),10)
gubuncd = requestCheckvar(request("gubuncd"),32)
shopid  = requestCheckvar(request("shopid"),32)

''groupid = requestCheckvar(request("groupid"),10) ''getPartnerId2GroupID(makerid)


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
ooffjungsan.FRectGroupid = groupid
''if (groupid<>"") then  '' 어드민은 그룹ID 와 상관없이 조회
    ooffjungsan.GetOneOffJungsanMaster
''end if

if (ooffjungsan.FResultCount<1) then
    response.write "<script >alert('검색 결과가 없습니다.');</script>"
    dbget.close()	:	response.End
end if

Dim IsCommissionTax : IsCommissionTax = ooffjungsan.FOneItem.IsCommissionTax

dim ooffjungsandetail
set ooffjungsandetail = new COffJungsan
ooffjungsandetail.FPageSize   = 1000
ooffjungsandetail.FRectIDX = idx
ooffjungsandetail.FRectMakerid = ooffjungsan.FOneItem.FMakerid
ooffjungsandetail.GetOffJungsanDetailSummaryList

dim ooffjungsandetaillist
set ooffjungsandetaillist = new COffJungsan
ooffjungsandetaillist.FPageSize  = 3000
ooffjungsandetaillist.FRectIDX = idx
ooffjungsandetaillist.FRectGubunCD = gubuncd
ooffjungsandetaillist.FRectShopid  = shopid

if (idx<>"") and ((shopid<>"") or (gubuncd<>""))  then
ooffjungsandetaillist.GetOffJungsanDetailList
end if

dim i
dim ttlitemno, ttlorgsellprice, ttlrealsellprice, ttlsuplyprice, ttlcommission
ttlitemno       = 0
ttlorgsellprice = 0
ttlrealsellprice= 0
ttlsuplyprice   = 0
ttlcommission   = 0

dim subitemno, subtotal
subitemno       = 0
subtotal        = 0

dim orgsellmargin, realsellmargin, selecteddefaultmargin
orgsellmargin   = 0
realsellmargin  = 0

%>
<script language='javascript'>
function PopDetailList(idx,gubuncd,shopid){
    location.href = '?idx=' + idx + '&gubuncd=' + gubuncd + '&shopid=' + shopid + '&makerid=<%=makerid%>';
}

function ExcelDetailList(idx,gubuncd,shopid){
alert('..');
return;
    location.href = 'off_jungsandetailsum_excelAdm.asp?idx=' + idx + '&gubuncd=' + gubuncd + '&shopid=' + shopid  + '&makerid=<%=makerid%>';
}

</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
        	<b>
        	<%= ooffjungsan.FOneItem.FTitle %>&nbsp;<%= ooffjungsan.FOneItem.Fmakerid %>&nbsp;&nbsp;
        	&nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.Fdifferencekey %> 차 &nbsp;&nbsp;
             &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.getJGubunName %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.GetSimpleTaxtypeName %>&nbsp;&nbsp;
            </b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>

<!-- 표 상단바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">정산구분</td>
      <td width="100">매장 명</td>
      <td width="100">정산구분</td>
      <td width="80">과세<br>구분</td>
      <td width="80">수량</td>
      <% if (IsCommissionTax) then %>
        <td width="90">구매총액</td>
	    <td width="80">기본판매<br>수수료</td>
        <td width="50">&nbsp;</td>
        <td width="80">매장할인액<br>(텐바이텐부담)</td>
        <td width="80">고객실주문액<br>(협력사매출액)</td>
		<td width="100">수수료</td>
		<td width="100">지급대상액<br>(정산확정액)</td>
      <% else %>
        <td width="150">판매가총액</td>
        <td width="150">매출액총액</td>
	    <td width="150">공급가총액</td>
	    <td width="100">공급마진율</td>
      <% end if %>
      <td width="40">상세<br>내역</td>
    </tr>
    <% if ooffjungsandetail.FResultCount>0 then %>
    <% for i=0 to ooffjungsandetail.FResultCount - 1 %>
    <%
    ttlitemno           = ttlitemno + ooffjungsandetail.FItemList(i).Ftot_itemno
    ttlorgsellprice     = ttlorgsellprice + ooffjungsandetail.FItemList(i).Ftot_orgsellprice
    ttlrealsellprice    = ttlrealsellprice + ooffjungsandetail.FItemList(i).Ftot_realsellprice
    ttlsuplyprice       = ttlsuplyprice + ooffjungsandetail.FItemList(i).Ftot_jungsanprice
    ttlcommission       = ttlcommission + ooffjungsandetail.FItemList(i).Ftot_commission
    %>
    <% if (shopid=ooffjungsandetail.FItemList(i).Fshopid) and (gubuncd=ooffjungsandetail.FItemList(i).Fgubuncd) then %>
    <% selecteddefaultmargin = ooffjungsandetail.FItemList(i).Fdefaultmargin %>
    <tr align="center" bgcolor="#BBBBDD">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
      <td><%= ooffjungsandetail.FItemList(i).getJSummaryGugunName %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fshopname %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= ooffjungsandetail.FItemList(i).GetItemVatTypeName %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <% if (IsCommissionTax) then %>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice+ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
      <td align="center">
            <% if (ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0) then %>
            <%= CLNG((ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice+ooffjungsandetail.FItemList(i).Ftot_commission)/ooffjungsandetail.FItemList(i).Ftot_orgsellprice*100*100)/100 %> %
            <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
      <% else %>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>
      <td align="center">
        <% if ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0 then %>
        <%= CLng((1-ooffjungsandetail.FItemList(i).Ftot_jungsanprice/ooffjungsandetail.FItemList(i).Ftot_orgsellprice)*10000)/100 %> %
        <% end if %>
      </td>
      <% end if %>
      <td align="center">
      	<a href="javascript:PopDetailList('<%= idx %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>','<%= ooffjungsandetail.FItemList(i).FShopid %>')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
      	<a href="javascript:ExcelDetailList('<%= idx %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>','<%= ooffjungsandetail.FItemList(i).FShopid %>')"><img src="/images/iexcel.gif" width="16" border="0"></a>
     </td>
    </tr>
    <% next %>

    <tr bgcolor="#FFFFFF">
        <td align="center">합계</td>
        <td colspan="3"></td>
        <td align="center"><%= FormatNumber(ttlitemno,0) %></td>

        <% if (IsCommissionTax) then %>
            <td align="right"><%= FormatNumber(ttlorgsellprice,0) %></td>
            <td align="right"><%= FormatNumber(ttlorgsellprice-ttlrealsellprice+ttlcommission,0) %></td>
            <td align="center">
            <% if (ttlorgsellprice<>0) then %>
            <%= CLNG((ttlorgsellprice-ttlrealsellprice+ttlcommission)/ttlorgsellprice*100*100)/100 %> %
            <% end if %>
            </td>
            <td align="right"><%= FormatNumber(ttlorgsellprice-ttlrealsellprice,0) %></td>
            <td align="right"><%= FormatNumber(ttlrealsellprice,0) %></td>
            <td align="right"><%= FormatNumber(ttlcommission,0) %></td>
            <td align="right"><%= FormatNumber(ttlsuplyprice,0) %></td>
        <% else %>
            <td align="right"><%= FormatNumber(ttlorgsellprice,0) %></td>
            <td align="right"><%= FormatNumber(ttlrealsellprice,0) %></td>
            <td align="right"><%= FormatNumber(ttlsuplyprice,0) %></td>
            <td align="center">
                <% if ttlorgsellprice<>0 then %>
                   <%= CLng((1-ttlsuplyprice/ttlorgsellprice)*10000)/100 %> %
                <% end if %>
            </td>
        <% end if %>

        <td align="center">
        </td>
    </tr>

    <% else %>
    <tr bgcolor="#FFFFFF">
      <td colspan="13" align="center">[검색 결과가 없습니다.]</td>
    </tr>
    <% end if %>
</table>
<br>



<%
subitemno = 0
subtotal  = 0
%>
<br>
<% if ooffjungsandetaillist.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="<%=CHKIIF(IsCommissionTax,15,11)%>" align="left">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
			<b>상세리스트</b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="70">매출코드</td>
      <td width="40">구분</td>
      <td width="60">상품코드</td>
      <td width="50">옵션코드</td>
      <td width="100">상품명</td>
      <td width="80">옵션명</td>
      <td width="40">수량</td>
      <% if (IsCommissionTax) then %>
      <td width="60">구매총액</td>
      <td width="60">기본판매<br>수수료</td>
      <td width="50">&nbsp;</td>
      <td width="70">매장할인액<br>(텐바이텐부담)</td>
      <td width="80">고객실주문액<br>(협력사매출액)</td>
      <td width="60">수수료</td>
      <td width="60">정산액</td>
      <td width="80">정산합계<br>(수량*정산액)</td>
      <% else %>
      <td width="50">판매가</td>
      <td width="50">공급가</td>
      <td width="60">공급마진율</td>
      <td width="80">공급가계<br>(수량*공급가)</td>
      <% end if %>

    </tr>
    <% for i=0 to ooffjungsandetaillist.FResultCount-1 %>
    <%
        subitemno   = subitemno + ooffjungsandetaillist.FItemList(i).FItemNo
        subtotal    = subtotal + ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo

    %>
    <tr  bgcolor="#FFFFFF">
      <td ><%= ooffjungsandetaillist.FItemList(i).Forderno %></td>
      <td ><%= ooffjungsandetaillist.FItemList(i).Fitemgubun %></td>
      <td ><%= ooffjungsandetaillist.FItemList(i).Fitemid %></td>
      <td ><%= ooffjungsandetaillist.FItemList(i).Fitemoption %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
      <% if (IsCommissionTax) then %>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice+ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
      <td align="center">
            <% if (ooffjungsandetaillist.FItemList(i).Forgsellprice<>0) then %>
            <%= CLNG((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice+ooffjungsandetaillist.FItemList(i).Fcommission)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100 %> %
            <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="right">
          <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
          <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
          <% else %>
          <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
          <% end if %>
      </td>
      <% else %>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="center">
      <% if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then %>
      <%= 100-CLNG((ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100) %> %
      <% end if %>
      </td>
      <td align="right">
          <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
          <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
          <% else %>
          <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
          <% end if %>
      </td>
      <% end if %>

    </tr>
   <% next %>
    <tr bgcolor="#FFFFFF">
        <td align="center">합계</td>
        <td colspan="5"></td>
        <td align="center">
            <% if (ooffjungsan.FOneItem.Ftot_itemno<>subitemno) then %>
            <b><%= FormatNumber(subitemno,0) %></b>
            <% else %>
            <%= FormatNumber(subitemno,0) %>
            <% end if %>
        </td>
        <td colspan="<%=CHKIIF(IsCommissionTax,7,3)%>"></td>
        <td align="right">
            <% if (ooffjungsan.FOneItem.Ftot_jungsanprice<>subtotal) then %>
            <b><%= FormatNumber(subtotal,0) %></b>
            <% else %>
            <%= FormatNumber(subtotal,0) %>
            <% end if %>
        </td>
    </tr>
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
<% end if %>
<%
set ooffjungsan = Nothing
set ooffjungsandetail = Nothing
set ooffjungsandetaillist = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->