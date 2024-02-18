<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminUPCHE.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/jumun/jumuncls.asp"-->
<%
dim orderserial
orderserial = requestCheckVar(request("orderserial"),11)

dim ojumun
set ojumun = new CJumunMaster
ojumun.FRectOrderSerial = orderserial

if (ojumun.FRectOrderSerial<>"") then
    ojumun.SearchJumunList
end if

if (ojumun.FResultCount<1) then
    dbget.close()	:	response.End
end if

dim ix
dim sellsum
sellsum = 0

dim IsCanNotView
ojumun.SearchJumunDetail orderserial
for ix=0 to ojumun.FJumunDetail.FDetailCount-1
    if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then
	    if UCase(ojumun.FJumunDetail.FJumunDetailList(ix).FMakerid)=UCase(session("ssBctID")) and (ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y") and (ojumun.FJumunDetail.FJumunDetailList(ix).Fcancelyn<>"Y") then
            if (ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate<3) or IsNULL(ojumun.FJumunDetail.FJumunDetailList(ix).FCurrstate) then
                IsCanNotView = true
            end if
        end if
    end if
next

%>
<% if (IsCanNotView) then %>
<script language='javascript'>
alert('발주이전 주문이거나 / 주문 확인 안하신 상품이 있습니다. \n\n텐바이텐에서 발주 후 \n\n업체주문관리>>*업체배송주문확인 에서 주문 확인 하신 후 사용하실 수 있습니다.');
window.close();
</script>
<% end if  %>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="200" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>"  background="/images/menubar_1px.gif">
						<font color="#333333"><b>주문상세내역</b></font>
					</td>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">

					</td>

				</tr>
			</table>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
        	<b>주문번호</b> : <%= ojumun.FMasterItemList(0).FOrderSerial %>&nbsp;&nbsp;&nbsp;&nbsp;
        	<b>구매자명</b> : <%= ojumun.FMasterItemList(0).FBuyName %>
		</td>
	</tr>
    <tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
		<td width="225" bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FOrderSerial %></td>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">사이트</td>
		<td width="225" bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FSitename %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
		<td bgcolor="#FFFFFF" colspan="3"><font color="<%= ojumun.FMasterItemList(0).IpkumDivColor %>"><%= ojumun.FMasterItemList(0).IpkumDivName %></font></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">주문일</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FRegDate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">입금일</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FIpkumDate %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">할인율</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FDiscountRate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">취소여부</td>
		<td bgcolor="#FFFFFF"><font color="<%= ojumun.FMasterItemList(0).CancelYnColor %>"><%= ojumun.FMasterItemList(0).CancelYnName %></font></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자</td>
		<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyName %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자전화</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyPhone %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FBuyHp %></td>
	</tr>
<% ''수정요망 detail.currstate로 %>
<% if (true) or ojumun.FMasterItemList(0).Fipkumdiv > 4 then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">수령인</td>
		<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqName %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">수령인전화</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqPhone %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">수령인핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FMasterItemList(0).FReqHp %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">수령인주소</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%= ojumun.FMasterItemList(0).FReqZipCode %>
			<br>
			<%= ojumun.FMasterItemList(0).FReqZipAddr %>
			&nbsp;<%= ojumun.FMasterItemList(0).FReqAddress %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<%= nl2br(ojumun.FMasterItemList(0).FComment) %>
		</td>
	</tr>
	<% If Not IsNULL(ojumun.FMasterItemList(0).Freqdate) then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>"><b>플라워 지정일</b></td>
		<td colspan="3" bgcolor="#FFFFFF">
		<%= ojumun.FMasterItemList(0).Freqdate %> 일  <%= ojumun.FMasterItemList(0).Freqtime %> 시경<br>
		(※ 플라워 지정상품일 경우에만 해당, 일반상품은 해당 안됨.)
		</td>
	</tr>
	<% end if %>
	<% If Not IsNULL(ojumun.FMasterItemList(0).Fcardribbon) then %>
	<% If ojumun.FMasterItemList(0).Fcardribbon <> 3 then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">카드 리본 구분</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% If ojumun.FMasterItemList(0).Fcardribbon = 1 then %>카드<% elseIf ojumun.FMasterItemList(0).Fcardribbon = 2 then %>리본<% else %> 없음<% End if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">카드리본메세지</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if ojumun.FMasterItemList(0).Ffromname<>"" then %>
		From.<%= nl2br(ojumun.FMasterItemList(0).Ffromname) %><br>
		<% End if %>
		<%= nl2br(ojumun.FMasterItemList(0).Fmessage) %>
		</td>
	</tr>
	<% End if %>
	<% End if %>

<% else %>
	<tr align="center">
		<td colspan=4 bgcolor="#FFFFFF"><font color="blue"><b>배송정보는 [업체주문통보] 상태 이후에 확인가능합니다.</b></font></td>
	</tr>
<% end if %>

</table>


<!--
<table border="1" cellspacing="0" cellpadding="0" class="a">
	<tr>
		<td width="100">배송옵션</td>
		<td width="200"><%= ojumun.FJumunDetail.BeasongOptionStr %></td>
	</tr>
	<tr>
		<td>배송비</td>
		<td align="right"><%= FormatNumber(ojumun.FJumunDetail.BeasongPay,0) %></td>
	</tr>
</table>
-->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
        	<b>주문상품정보</b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="35">수량</td>
		<td width="50">판매가격</td>
		<td width="35">배송<br>구분</td>
		<td width="35">취소<br>여부</td>
	</tr>
<% for ix=0 to ojumun.FJumunDetail.FDetailCount-1 %>
<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid <>0 then %>
	<% if UCase(ojumun.FJumunDetail.FJumunDetailList(ix).FMakerid)=UCase(session("ssBctID")) then %>
	<% sellsum = sellsum + ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost*ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %></td>
		<td><a href="http://www.thefingers.co.kr/diyshop/shop_prd.asp?itemid=<%= ojumun.FJumunDetail.FJumunDetailList(ix).Fitemid %>" target="_blank"><img src="<%= ojumun.FJumunDetail.FJumunDetailList(ix).FImageSmall %>" border="0"></a></td>
		<td align="left">
			<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemName %>
			<br>
			<% if ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName <> "" then %>
				<font color="blue">[<%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemoptionName %>]</font>
			<% end if %>
		</td>
		<td><%= ojumun.FJumunDetail.FJumunDetailList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(ojumun.FJumunDetail.FJumunDetailList(ix).Fitemcost,0) %></td>
		<td>
			<% if ojumun.FJumunDetail.FJumunDetailList(ix).Fisupchebeasong="Y" then %>
			<font color="red">업체</font>
			<% else %>
			10x10
			<% end if %>
		</td>
		<td>
			<font color="<%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateColor %>"><%= ojumun.FJumunDetail.FJumunDetailList(ix).CancelStateStr %></font>
		</td>
	</tr>
	<% if ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail <> "" then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="7"><%= nl2BR(ojumun.FJumunDetail.FJumunDetailList(ix).Frequiredetail) %></td>
	</tr>
	<% end if %>
	<% end if %>
<% end if %>
<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td>합계</td>
		<td colspan="4" align="right"><%= FormatNumber(sellsum,0) %></td>
		<td colspan="2"></td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="50" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">

		</td>
		<td colspan="15">
        	<font color="blue">
        		<b>본 자료는 배송을 위한 정보로만 사용해야 합니다.<br>
				이외의 목적으로 사용시 민,형사상 책임은 해당 업체에게 있습니다.</b>
			</font>
		</td>
	</tr>
</table>

<!-- 표 하단바 끝-->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->