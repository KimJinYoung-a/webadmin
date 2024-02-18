<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정산진행내역
' History : 서동석 생성
'			2021.04.28 한용민 수정(노출되는 항목들 변경요청. 재무팀:최현희)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim i, suplysum, suplytotalsum, duplicated
dim sumttl1, sumttl2, ItemNo, CpnNotAppliedPrice, CouponDiscountCommission, CouponDiscount, Reducedprice, Commission, suplycash
dim TTLitemCNT, TTLSellcashSum, TTLCouponDiscountSum, TTLReducedpriceSum
dim TTLCommissionSum, TTLSuplycashSum
dim TTLPgCommissionSum, TTLCpnNotAppliedPrice
dim id,gubun, makerid, groupid, itemvatyn
makerid = requestCheckVar(request("makerid"),32)
id      = requestCheckVar(request("id"),10)
gubun   = requestCheckVar(request("gubun"),20)
itemvatyn = requestCheckVar(request("itemvatyn"),10)


''groupid = requestCheckvar(request("groupid"),10) ''getPartnerId2GroupID(makerid)

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectgubun = gubun
ojungsan.FRectGroupID = groupid

''if (groupid<>"") then '' 어드민은 그룹ID 와 상관없이 조회
    ojungsan.JungsanMasterList
''end if

if ojungsan.FresultCount <1 then
    set ojungsan = Nothing
	dbget.close()	:	response.End
end if

Dim IsShowCpnNotAppliedPrice ''2018/07/02
IsShowCpnNotAppliedPrice = (ojungsan.FItemList(0).FYYYYMM>="2018-06") and (ojungsan.FItemList(0).FJGubun="CC") 
if (application("Svr_Info")	= "Dev") then IsShowCpnNotAppliedPrice = true

dim ojungsanSubsmr
set ojungsanSubsmr = new CUpcheJungsan
ojungsanSubsmr.FRectId = id
ojungsanSubsmr.FRectdesigner = ojungsan.FItemList(0).Fdesignerid
ojungsanSubsmr.getJungsanSubSummary

Dim IsCommissionTax : IsCommissionTax=ojungsan.FItemList(0).IsCommissionTax
Dim IsCommissionETCTax : IsCommissionETCTax=ojungsan.FItemList(0).IsCommissionETCTax
sumttl1=0
sumttl2=0
ItemNo=0
CpnNotAppliedPrice=0
CouponDiscountCommission=0
CouponDiscount=0
Reducedprice=0
Commission=0
suplycash=0
%>

<script language='javascript'>
function JunsanDetailList(id,gubun,itemvatyn){
    location.href = '?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn + '&makerid=<%=makerid%>';
}
function ExcelJunsanDetailList(id,gubun,itemvatyn){
//alert('..');
//return;
    location.href = '/admin/upchejungsan/jungsandetailsumAdm_excel.asp?id=' + id + '&gubun=' + gubun + '&itemvatyn='+itemvatyn + '&makerid=<%=makerid%>';
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
        	<b>
        	온라인 <%= ojungsan.FItemList(0).Ftitle %>&nbsp;[<%= ojungsan.FItemList(0).Fdesignerid %>]
        	&nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).Fdifferencekey %> 차
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).getJGuBunName %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            <%= ojungsan.FItemList(0).getTaxTypeName %>&nbsp;&nbsp;
            </b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="100">정산구분</td>
    	<td width="100">구분</td>
    	<td width="50">과세<br>구분</td>
		<td width="50">총건수</td>
		<% if (IsCommissionTax) then %>
		    <% if (IsCommissionETCTax) then %>
    		<td width="100">프로모션비용<br>(협력사 부담)</td>
    		<td width="100">지급대상액<br>(정산확정액)</td>  
		    <% else %>
		    <% if (IsShowCpnNotAppliedPrice) then %><td width="90">판매총액</td><% end if %>
    		<td width="90">구매총액<br>(상품쿠폰적용)</td>
    		<td width="80">기본판매<br>수수료</td>
    		<td width="50">&nbsp;</td>
            <td width="80">쿠폰할인액<br>(텐바이텐부담)</td>
            <td width="80">고객실주문액<br>(협력사매출액)</td>
    		<td width="90">수수료</td>
    		<td width="90">결제대행수수료</td>
    		<td width="100">지급대상액<br>(정산확정액)</td>
    	    <% end if %>
		<% else %>
    		<td width="150">판매가총액</td>
    		<td width="150">공급가총액</td>
    		<td width="100">공급마진율</td>
		<% end if %>
		<td width="50">상세내역</td>
	</tr>

    <% for i=0 to ojungsanSubsmr.FResultCount-1 %>
    <%
        TTLitemCNT              = TTLitemCNT + ojungsanSubsmr.FItemList(i).FitemCNT
        TTLSellcashSum          = TTLSellcashSum + ojungsanSubsmr.FItemList(i).getSellcashSum
        TTLCouponDiscountSum    = TTLCouponDiscountSum + ojungsanSubsmr.FItemList(i).getCouponDiscountSum
        TTLReducedpriceSum      = TTLReducedpriceSum + ojungsanSubsmr.FItemList(i).getReducedpriceSum
        TTLCommissionSum        = TTLCommissionSum + ojungsanSubsmr.FItemList(i).getCommissionSum
        TTLPgCommissionSum      = TTLPgCommissionSum + ojungsanSubsmr.FItemList(i).getPgCommissionSum
        TTLSuplycashSum         = TTLSuplycashSum + ojungsanSubsmr.FItemList(i).getSuplycashSum
        
        TTLCpnNotAppliedPrice   = TTLCpnNotAppliedPrice + ojungsanSubsmr.FItemList(i).FCpnNotAppliedPriceSum

    %>
    <tr bgcolor='<%=CHKIIF(gubun=ojungsanSubsmr.FItemList(i).Fgubuncd and (Not IsCommissionTax or (IsCommissionTax and itemvatyn=ojungsanSubsmr.FItemList(i).FitemVatyn)),"#CCCCFF","#FFFFFF") %>' >
        <td align="center"><%= ojungsanSubsmr.FItemList(i).getJSummaryGugunName %></td>
        <td><%= ojungsanSubsmr.FItemList(i).getJGubuncd2Name %></td>
        <td align="center"><%= ojungsanSubsmr.FItemList(i).getTaxTypeName %></td>
        <td align="center"><%= ojungsanSubsmr.FItemList(i).FitemCNT %></td>
        <% if (IsCommissionTax) then %>
            <% if (IsCommissionETCTax) then %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <% else %>
            <% if (IsShowCpnNotAppliedPrice) then %><td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).FCpnNotAppliedPriceSum,0) %></td><% end if %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSellcashSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCouponDiscountSum+ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="center">
              <% if (ojungsanSubsmr.FItemList(i).getSellcashSum<>0) then %>
              <%= CLNG((ojungsanSubsmr.FItemList(i).getCouponDiscountSum+ojungsanSubsmr.FItemList(i).getCommissionSum)/ojungsanSubsmr.FItemList(i).getSellcashSum*100*100)/100 %> %
              <% end if %>
            </td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCouponDiscountSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getReducedpriceSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getCommissionSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getPGCommissionSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <% end if %>
        <% else %>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSellcashSum,0) %></td>
            <td align="right"><%= FormatNumber(ojungsanSubsmr.FItemList(i).getSuplycashSum,0) %></td>
            <td align="center">
                <% if ojungsanSubsmr.FItemList(i).getSellcashSum<>0 then %>
                <%= CLng((1-ojungsanSubsmr.FItemList(i).getSuplycashSum/ojungsanSubsmr.FItemList(i).getSellcashSum)*10000)/100 %> %
                <% end if %>
            </td>
        <% end if %>

        <td align="center">
		  <a href="javascript:JunsanDetailList('<%= id %>','<%=ojungsanSubsmr.FItemList(i).Fgubuncd%>','<%=CHKIIF(IsCommissionTax,ojungsanSubsmr.FItemList(i).FitemVatyn,"")%>')"><img src="/images/icon_search.jpg" width="16" border="0"></a>
		  <a href="#" onclick="ExcelJunsanDetailList('<%= id %>','<%=ojungsanSubsmr.FItemList(i).Fgubuncd%>','<%=CHKIIF(IsCommissionTax,ojungsanSubsmr.FItemList(i).FitemVatyn,"")%>'); return false;"><img src="/images/iexcel.gif" width="16" border="0"></a>
		</td>
    </tr>
    <% next %>

	<tr bgcolor="#FFFFFF">
		<td align=center>총계</td>
		<td></td>
		<td></td>
		<td align=center><%= FormatNumber(TTLitemCNT,0) %></td>
		<% if (IsCommissionTax) then %>
		    <% if (IsCommissionETCTax) then %>
    		<td align=right><%= FormatNumber(TTLCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
		    <% else %>
		    <% if (IsShowCpnNotAppliedPrice) then %><td align=right><%= FormatNumber(TTLCpnNotAppliedPrice,0) %></td><% end if %>
    		<td align=right><%= FormatNumber(TTLSellcashSum,0) %></td>
    		<td align="right"><%= FormatNumber(TTLCouponDiscountSum+TTLCommissionSum,0) %></td>
            <td align="center">
              <% if (TTLSellcashSum<>0) then %>
              <%= CLNG((TTLCouponDiscountSum+TTLCommissionSum)/TTLSellcashSum*100*100)/100 %> %
              <% end if %>
            </td>
    		<td align=right><%= FormatNumber(TTLCouponDiscountSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLReducedpriceSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLPgCommissionSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
    	    <% end if %>
		<% else %>
    		<td align=right><%= FormatNumber(TTLSellcashSum,0) %></td>
    		<td align=right><%= FormatNumber(TTLSuplycashSum,0) %></td>
    		<td align=center>
		<% if TTLSellcashSum<>0 then %>
		    <%= CLng((1-TTLSuplycashSum/TTLSellcashSum)*10000)/100 %> %
		<% end if %>
		</td>
		<% end if %>
		<td></td>
	</tr>
</table>

<br>

<%
dim ojungsandetail
set ojungsandetail = new CUpcheJungsan
ojungsandetail.FRectId = id
ojungsandetail.FRectgubun = gubun
ojungsandetail.FRectdesigner = session("ssBctID")
ojungsandetail.FRectOrder = "orderserial"
ojungsandetail.FRectItemVatYn = itemvatyn
'' 1357 이전내역은 정산방식이 다름(재고기준정산)
if (id>1357) and (gubun<>"")   then
    ojungsandetail.JungsanDetailList
end if
%>
<!-- 주문건별 리스트 시작-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="<%=CHKIIF(IsCommissionTax,20,13)%>" align="left">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
			<b>주문/출고/입고건별 상세리스트</b>
			&nbsp;&nbsp;
			<% if ojungsandetail.FRectgubun="maeip" then %>
			물류센터입고확인일 기준으로 등록됩니다.
			<% else %>
			배송완료일 기준
			<% end if %>

			<% if ojungsandetail.FResultCount>=10000 then %>
			(최대 <%= ojungsandetail.FResultCount %> 건 표시)
			<% end if %>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td>입출코드</td>
      <td>판매채널</td>
      <td>구매자</td>
      <td>수령인</td>
      <td>상품코드</td>
      <td>옵션코드</td>
      <td>상품명</td>
      <td>옵션명</td>
      <td>수량</td>
      <% if (IsCommissionTax) then %>
        <% if (IsCommissionETCTax) then %>
            <td>프로모션비용<br>(협력사 부담)</td>  
            <td>정산액</td>
            <% '<td>정산합계<br>(수량*정산액)</td> %>
            <td>주결제수단</td>
        <% else %>
            <% if (IsShowCpnNotAppliedPrice) then %><td>판매총액</td><% end if %>
            <td>구매총액<br>(상품쿠폰적용)</td>
            <td>기본판매<br>수수료</td>
            <td>계약수수료율</td>
            <td>쿠폰할인액<br>(텐바이텐부담)</td>
            <td>고객실주문액<br>(협력사매출액)</td>
            <td>계약수수료</td>
            <% '<td>결제대행<br>수수료</td> %>
            <td>정산액</td>
            <% '<td>정산합계<br>(수량*정산액)</td> %>
            <td>주결제수단</td>
        <% end if %>
      <% else %>
      <td>판매단가</td>
      <td>공급단가</td>
      <td>공급마진율</td>
      <td>공급가합계<br>(수량*공급가)</td>
      <% end if %>
    </tr>
<% if ojungsandetail.FResultCount>0 and ojungsandetail.FRectgubun<>"" then %>
    <% for i=0 to ojungsandetail.FResultCount-1 %>

    <%
	sumttl1 = sumttl1 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsellcash
	sumttl2 = sumttl2 + ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash
	ItemNo = ItemNo + ojungsandetail.FItemList(i).FItemNo
	CpnNotAppliedPrice = CpnNotAppliedPrice + ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo
	CouponDiscountCommission = CouponDiscountCommission + (ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)*ojungsandetail.FItemList(i).FItemNo
	CouponDiscount = CouponDiscount + ojungsandetail.FItemList(i).getCouponDiscount*ojungsandetail.FItemList(i).FItemNo
	Reducedprice = Reducedprice + ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo
	Commission = Commission + ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo
	suplycash = suplycash + ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo
	%>
    <tr bgcolor="#FFFFFF" align="center">
      <td><%= ojungsandetail.FItemList(i).Fmastercode %></td>
      <td><%= ojungsandetail.FItemList(i).Fsitename %></td>
        <td>
            <% if C_CriticInfoUserLV1 then %>
                <%= ojungsandetail.FItemList(i).FBuyname %>
            <% else %>
                <%= AstarUserName(ojungsandetail.FItemList(i).FBuyname) %>
            <% end if %>
        </td>
        <td>
            <% if C_CriticInfoUserLV1 then %>
                <%= ojungsandetail.FItemList(i).FBuyname %>
            <% else %>
                <%= AstarUserName(ojungsandetail.FItemList(i).FReqname) %>
            <% end if %>
        </td>
      <td align="center"><%= ojungsandetail.FItemList(i).Fitemid %></td>
      <td align="center"><%= ojungsandetail.FItemList(i).FitemOption %></td>
      <td align="left"><%= ojungsandetail.FItemList(i).FItemName %></td>
      <td><%= ojungsandetail.FItemList(i).FItemOptionName %></td>
      <td><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo) %>"><%= ojungsandetail.FItemList(i).FItemNo %></font></td>
      <% if (IsCommissionTax) then %>
          <% if (IsCommissionETCTax) then %>
			<% '프로모션비용(협력사 부담)  %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% '정산액  %>
			<td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo,0) %></font></td>
            <!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>-->
            <td align="center"><%=ojungsandetail.FItemList(i).getpaymethodName %></td>  
          <% else %>
			<%
			' 판매총액
			if (IsShowCpnNotAppliedPrice) then
			%>
				<td align="right">
					<font color="<%= MinusFont(ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo) %>">
					<%= FormatNumber(ojungsandetail.FItemList(i).FCpnNotAppliedPrice*ojungsandetail.FItemList(i).FItemNo,0) %></font>
				</td>
			<% end if %>
			<% ' 구매총액 %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% ' 기본판매수수료 %>
			<td align="right">
				<%= FormatNumber((ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)*ojungsandetail.FItemList(i).FItemNo,0) %>
			</td>
            <td align="center">
            <% if (ojungsandetail.FItemList(i).Fsellcash<>0) then %>
            <%= CLNG((ojungsandetail.FItemList(i).getCouponDiscount+ojungsandetail.FItemList(i).getCommission)/ojungsandetail.FItemList(i).Fsellcash*100) %> %
            <% end if %>
            </td>
			<% ' 쿠폰할인액(텐바이텐부담) %>
			<td align="right"><%= FormatNumber(ojungsandetail.FItemList(i).getCouponDiscount*ojungsandetail.FItemList(i).FItemNo,0) %></td>
			<% ' 고객실주문액(협력사매출액) %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).getReducedprice*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
			<% ' 수수료 %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).getCommission*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
            <!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).getPgCommission) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).getPgCommission,0) %></font></td>-->
			<% ' 정산액 %>
			<td align="right">
				<font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo) %>">
				<%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash*ojungsandetail.FItemList(i).FItemNo,0) %></font>
			</td>
            <!--<td align="right"><font color="<%'= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%'= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>-->
            <td align="center"><%=ojungsandetail.FItemList(i).getpaymethodName %></td>
          <% end if %>
      <% else %>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center">
      <% if ojungsandetail.FItemList(i).Fsellcash<>0 then %>
      <%= 100-CLNG((ojungsandetail.FItemList(i).Fsuplycash)/ojungsandetail.FItemList(i).Fsellcash*100) %> %
      <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsandetail.FItemList(i).FItemNo*ojungsandetail.FItemList(i).Fsuplycash,0) %></font></td>
      <% end if %>
   </tr>
    <% if (i mod 1000)=0 then response.flush %>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
        <% if (IsCommissionTax) then %>
            <% if (IsCommissionETCTax) then %>
                <td><strong>합계</strong></td>
				<td colspan="8"></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
                <!--<td align="right"><strong><%'=FormatNumber(sumttl2,0)%></strong></td>-->
            <% else %>
			    <td colspan=8><strong>합계</strong></td>
				<td><strong><%= FormatNumber(ItemNo,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CpnNotAppliedPrice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(sumttl1,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscountCommission,0) %></strong></td>
				<td></td>
				<td align="right"><strong><%= FormatNumber(CouponDiscount,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Reducedprice,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(Commission,0) %></strong></td>
				<td align="right"><strong><%= FormatNumber(suplycash,0) %></strong></td>
                <!--<td align="right"><strong><%'=FormatNumber(sumttl2,0)%></strong></td>-->
            <% end if %>

            <td>&nbsp;</td>
        <% else %>
            <td><strong>합계</strong></td>
            <td colspan="11"></td>
            <td align="right"><strong><%=FormatNumber(sumttl2,0)%></strong></td>
        <% end if %>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="<%=CHKIIF(IsCommissionTax,20,13)%>" align="center"><img src="/images/icon_search.jpg" width="16" border="0" align="absbottom">&nbsp;상세내역을 선택하세요.</td>
    </tr>
<% end if %>
</table>
<!-- 주문건별 리스트 끝-->

<%
set ojungsan = Nothing
set ojungsanSubsmr = Nothing
set ojungsandetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->