<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim idx, gubuncd, shopid
idx     = requestCheckvar(request("idx"),10)
gubuncd = requestCheckvar(request("gubuncd"),16)
shopid  = requestCheckvar(request("shopid"),32)

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
''ooffjungsan.FRectMakerid = 업체일경우 session 업체아이디
ooffjungsan.GetOneOffJungsanMaster

if (ooffjungsan.FResultCount<1) then
    response.write "<script >alert('검색 결과가 없습니다.');</script>"
    dbget.close()	:	response.End
end if

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

if (shopid<>"") or (gubuncd<>"")  then
    ooffjungsandetaillist.GetOffJungsanDetailList
end if

dim i
dim ttlitemno, ttlorgsellprice, ttlrealsellprice, ttlsuplyprice, ttlcommission
ttlitemno       = 0
ttlorgsellprice = 0
ttlrealsellprice= 0
ttlsuplyprice   = 0
ttlcommission  = 0
dim subitemno, subtotal
subitemno       = 0
subtotal        = 0

dim orgsellmargin, realsellmargin, selecteddefaultmargin
orgsellmargin   = 0
realsellmargin  = 0

%>
<script language='javascript'>
function PopDetailList(idx,gubuncd,shopid){
    location.href = '?idx=' + idx + '&gubuncd=' + gubuncd + '&shopid=' + shopid ;
}

function PopDetailEdit(idx,gubuncd,shopid){
    var popwin = window.open('off_jungsandetailedit.asp?idx=' + idx + '&gubuncd=' + gubuncd + '&shopid=' + shopid,'off_jungsandetailedit','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddEtcDetail(frm,idx){
    var shopid = frm.shopid.value;
    if (shopid.length<1){
        alert('매장을 선택 하세요.');
        frm.shopid.focus();
        return;
    }
    PopDetailEdit(idx,'B999',shopid);
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1"  class="a" bgcolor="#999999">
    <form name="frm" method="get" action="">
    <input type="hidden" name="idx" value="<%= idx %>">

    <tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="2" width="50" bgcolor="#EEEEEE"><!-- 검색<br>조건 --></td>
        <td align="left">
            <%= ooffjungsan.FOneItem.FTitle %>&nbsp;<%= ooffjungsan.FOneItem.Fmakerid %>&nbsp;&nbsp;
            <%= ooffjungsan.FOneItem.Fdifferencekey %> 차 &nbsp;&nbsp;
            <font color="<%= ooffjungsan.FOneItem.GetTaxtypeNameColor %>"><%= ooffjungsan.FOneItem.GetSimpleTaxtypeName %></font> &nbsp;&nbsp;
            총 정산액 : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_jungsanprice,0) %>&nbsp;&nbsp;
            총 판매상품수 : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_itemno,0) %>&nbsp;&nbsp;
            <% if (ooffjungsan.FOneItem.IsCommissionTax) then %>
            총 수수료 : <%= FormatNumber(ooffjungsan.FOneItem.Ftotalcommission,0) %>
            <% end if %>
        </td>
        <td rowspan="2" width="50" bgcolor="#EEEEEE">
            <!--
            <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
            -->
        </td>
    </tr>

    </form>
</table>
<!-- 표 상단바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmetc" >
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td align="right">
        기타내역 추가시 :   <% drawSelectBoxOffShopAll "shopid","" %>
        <input type="button" value="추가" onclick="AddEtcDetail(frmetc,'<%= idx %>')">
        </td>
    </tr>
    </form>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">가맹점코드</td>
      <td width="100">가맹점 명</td>
      <td width="100">기본정산구분</td>
      <td width="100">정산구분</td>
      <td width="80">총상품건수</td>
      <td width="90">판매가합</td>
      <td width="90">매출액</td>
      <td width="80">수수료</td>
      <td width="90">정산액</td>
      <td width="50">소비가<br>마진</td>
      <td width="50">매출액<br>마진</td>
      <td width="40">상세<br>내역</td>
      <td width="50">수정</td>
    </tr>
    <% if ooffjungsandetail.FResultCount>0 then %>
    <% for i=0 to ooffjungsandetail.FResultCount - 1 %>
    <%
    ttlitemno           = ttlitemno + ooffjungsandetail.FItemList(i).Ftot_itemno
    ttlorgsellprice     = ttlorgsellprice + ooffjungsandetail.FItemList(i).Ftot_orgsellprice
    ttlrealsellprice    = ttlrealsellprice + ooffjungsandetail.FItemList(i).Ftot_realsellprice
    ttlsuplyprice       = ttlsuplyprice + ooffjungsandetail.FItemList(i).Ftot_jungsanprice
    ttlcommission       = ttlcommission + ooffjungsandetail.FItemList(i).Ftot_commission
    if ooffjungsandetail.FItemList(i).Ftot_orgsellprice<>0 then
        orgsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_orgsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_orgsellprice*100*100)/100
    else
        orgsellmargin = 0
    end if

    if ooffjungsandetail.FItemList(i).Ftot_realsellprice<>0 then
        realsellmargin = CLng((ooffjungsandetail.FItemList(i).Ftot_realsellprice-ooffjungsandetail.FItemList(i).Ftot_jungsanprice)/ooffjungsandetail.FItemList(i).Ftot_realsellprice*100*100)/100
    else
        realsellmargin = 0
    end if

    %>
    <% if (shopid=ooffjungsandetail.FItemList(i).Fshopid) and (gubuncd=ooffjungsandetail.FItemList(i).Fgubuncd) then %>
    <% selecteddefaultmargin = ooffjungsandetail.FItemList(i).Fdefaultmargin %>
    <tr align="center" bgcolor="#BBBBDD">
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    <% end if %>
      <td><%= ooffjungsandetail.FItemList(i).Fshopid %></td>
      <td><%= ooffjungsandetail.FItemList(i).Fshopname %></td>
      <td>
        <%= ooffjungsandetail.FItemList(i).GetChargeDivName %>,
        <%= ooffjungsandetail.FItemList(i).Fdefaultmargin %>,
        <% if ooffjungsandetail.FItemList(i).Fautojungsan="N" then response.write "<font color='blue'>수</font>" else response.write "자" %>,
        <%= ooffjungsandetail.FItemList(i).Fautojungsandiv %>
      </td>
      <td><%= ooffjungsandetail.FItemList(i).Fcomm_name %></td>
      <td><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_itemno,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_orgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_realsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_commission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetail.FItemList(i).Ftot_jungsanprice,0) %></td>

      <td align="center">
      <% if ooffjungsandetail.FItemList(i).Fdefaultmargin<>orgsellmargin then %>
        <font color="red"><%= orgsellmargin %></font> %
      <% else %>
        <%= orgsellmargin %> %
      <% end if %>
      </td>
      <td align="center">
      <% if ooffjungsandetail.FItemList(i).Fdefaultmargin<>realsellmargin then %>
        <font color="blue"><%= realsellmargin %></font> %
      <% else %>
        <%= realsellmargin %> %
      <% end if %>

      </td>
      <td><a href="javascript:PopDetailList('<%= idx %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>','<%= ooffjungsandetail.FItemList(i).FShopid %>')"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
      <td><a href="javascript:PopDetailEdit('<%= idx %>','<%= ooffjungsandetail.FItemList(i).FGubuncd %>','<%= ooffjungsandetail.FItemList(i).FShopid %>')"><img src="/images/icon_modify.gif" border="0" width="45"></a></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td align="center">합계</td>
      <td colspan="3"></td>
      <td align="center"><%= FormatNumber(ttlitemno,0) %></td>
      <td align="right"><%= FormatNumber(ttlorgsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlrealsellprice,0) %></td>
      <td align="right"><%= FormatNumber(ttlcommission,0) %></td>
      <td align="right"><%= FormatNumber(ttlsuplyprice,0) %></td>

      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <% else %>
    <tr bgcolor="#FFFFFF">
      <td colspan="13" align="center">[검색 결과가 없습니다.]</td>
    </tr>
    <% end if %>
</table>
<br>

<% if ooffjungsandetaillist.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="70">매출코드</td>
      <td width="70">상품코드</td>
      <td width="100">상품명</td>
      <td width="80">옵션명</td>
      <td width="60">판매가</td>
      <td width="60">실판매가</td>
      <td width="60">수수료</td>
      <td width="60">정산액</td>
      <td width="40">소비<br>마진</td>
      <td width="40">매출<br>마진</td>
      <td width="40">갯수</td>
      <td width="36">센터<br>매입<br>구분</td>
      <td width="64">정산액</td>
    </tr>
    <% for i=0 to ooffjungsandetaillist.FResultCount-1 %>
    <%
        subitemno   = subitemno + ooffjungsandetaillist.FItemList(i).FItemNo
        subtotal    = subtotal + ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo

        if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then
            orgsellmargin = CLng((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100
        else
            orgsellmargin = 0
        end if

        if ooffjungsandetaillist.FItemList(i).Frealsellprice<>0 then
            realsellmargin = CLng((ooffjungsandetaillist.FItemList(i).Frealsellprice-ooffjungsandetaillist.FItemList(i).Fsuplyprice)/ooffjungsandetaillist.FItemList(i).Frealsellprice*100*100)/100
        else
            realsellmargin = 0
        end if
    %>
    <tr  bgcolor="#FFFFFF">
      <td><%= ooffjungsandetaillist.FItemList(i).Forderno %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).GetBarCode %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemName %></td>
      <td><%= ooffjungsandetaillist.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Forgsellprice,0) %></td>
      <td align="right">
        <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Frealsellprice,0) %>
        <% if (ooffjungsandetaillist.FItemList(i).Frealsellprice<>ooffjungsandetaillist.FItemList(i).Forgsellprice) then %>
            <% if ooffjungsandetaillist.FItemList(i).Forgsellprice<>0 then %>
                <br><font color="red"><%= Clng((ooffjungsandetaillist.FItemList(i).Forgsellprice-ooffjungsandetaillist.FItemList(i).Frealsellprice)/ooffjungsandetaillist.FItemList(i).Forgsellprice*100*100)/100 %></font> %
            <% end if %>
        <% end if %>
      </td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fcommission,0) %></td>
      <td align="right"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice,0) %></td>
      <td align="center">
      <% if selecteddefaultmargin<>orgsellmargin then %>
        <font color="red"><%= orgsellmargin %></font> %
      <% else %>
        <%= orgsellmargin %> %
      <% end if %>
      </td>
      <td align="center">
      <% if orgsellmargin<>realsellmargin then %>
        <font color="blue"><%= realsellmargin %></font> %
      <% else %>
        <%= realsellmargin %> %
      <% end if %>
      </td>
      <td align="center"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).FItemNo,0) %></td>
      <td align="center">
      <% if ooffjungsandetaillist.FItemList(i).Fcentermwdiv="M" then %>
      <b><%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %></b>
      <% else %>
      <%= ooffjungsandetaillist.FItemList(i).Fcentermwdiv %>
      <% end if %>
      </td>
      <td align="right">
      <% if ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo<1 then %>
     <font color="red"><%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %></font>
      <% else %>
      <%= FormatNumber(ooffjungsandetaillist.FItemList(i).Fsuplyprice*ooffjungsandetaillist.FItemList(i).FItemNo,0) %>
      <% end if %>
      </td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
        <td align="center">합계</td>
        <td colspan="9"></td>
        <td align="center">
            <% if (ooffjungsan.FOneItem.Ftot_itemno<>subitemno) then %>
            <b><%= FormatNumber(subitemno,0) %></b>
            <% else %>
            <%= FormatNumber(subitemno,0) %>
            <% end if %>
        </td>
        <td></td>
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