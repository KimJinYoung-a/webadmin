<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################

dim IsItemCanceled, IsPossibleModifyItem, IsItemDisabled, IsItemChecked
dim ocsOrderDetail
dim BaesongMethod, SumBeasongPayNotCanceled, SumItemCostSumNotCanceled
dim strhtmldisabled, strhtmlcancel, strhtmlmodify

'디테일
set ocsOrderDetail = new COrder

	'//접수 상태에서는 전체 주문목록
	if (IsStatusRegister) then

		'/주문 테이블 masteridx
		ocsOrderDetail.FRectmasteridx = masteridx
	    ocsOrderDetail.fGetOrderDetailByCsDetail

	'//수정, 완료상태에서는 접수된 내역만 보여줌
	else

		'/cs테이블 masteridx
		ocsOrderDetail.FRectCsAsID = csmasteridx
	    ocsOrderDetail.fGetCsDetailList
	end if

%>
<% if (IsDisplayItemList and (ocsOrderDetail.FTotalCount > 0)) then %>
<tr>
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">접수상품</td>
            <td bgcolor="#FFFFFF">
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td>선택</td>
					<td>구분</td>
					<td>상품코드</td>
					<td>브랜드ID</td>
					<td>상품명<font color="blue">[옵션명]</font></td>
					<td>
						<% if (ocsaslist.FOneItem.IsCancelProcess_off) then %>
							취소/원주문
						<% else %>
							접수/원주문
						<% end if %>
					</td>
					<td>판매가격</td>
					<td>사유구분</td>
				</tr>
				<% '스크립트를 단순화하기 위해 아래와 같이 더미를 더 만들어 둔다.(orderdetailidx 가 한개인 경우와 2개이상인 경우를 분리해서 작성하지 않아도 된다.) %>
				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="Deliverdetailidx">
				<input type="hidden" name="DeliverMakerid">
				<input type="hidden" name="Deliveritemcost">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="dummystopper" value="">

				<input type="hidden" name="dummystarter" value="">
				<input type="hidden" name="orderdetailidx">
				<input type="hidden" name="odlvtype">
				<input type="hidden" name="itemno">
				<input type="hidden" name="regitemno">
				<input type="hidden" name="makerid">
				<input type="hidden" name="dummystopper" value="">
				<%
				SumBeasongPayNotCanceled = 0
				SumItemCostSumNotCanceled = 0

				for i = 0 to ocsOrderDetail.FResultCount - 1

				IsItemCanceled = (ocsOrderDetail.FItemList(i).FCancelyn = "Y")

				strhtmlcancel = ""

				if (IsItemCanceled) then
					'취소
					strhtmlcancel = "bgcolor='#CCCCCC' class='gray'"
				else
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'체크가능상품
						strhtmlcancel = "bgcolor='#FFFFFF'"
					else
						'체크불가상품
						strhtmlcancel = "bgcolor='#EEEEEE' class='gray'"
					end if
				end if

				strhtmldisabled = ""

				if (IsStatusRegister) then
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'체크가능상품
						strhtmldisabled = ""
					else
						'체크불가상품
						strhtmldisabled = "disabled"
					end if
				else
					'접수이후에는 항상 체크되어 있고, 수정불가
					strhtmldisabled = "checked disabled"
				end if

				strhtmlmodify = ""

				if (IsStatusRegister or (IsStatusEdit and IsCSReturnProcess_off(divcd))) then
					if (IsPossibleCheckItem_off(divcd, IsOrderCanceled, IsItemCanceled) = true) then
						'체크가능상품
						strhtmlmodify = ""
					else
						'체크불가상품
						strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
					end if
				else
					'접수이후에는 수정불가
					strhtmlmodify = "style='text-align:center;background-color:#DDDDFF;' readonly"
				end if

				%>

				<%
				'상품 리스트
				distinctid = ocsOrderDetail.FItemList(i).Forgdetailidx
				%>
				<tr align="center" <%= strhtmlcancel %>>
					<td height="25">
						<input type="hidden" name="dummystarter" value="">
						<input type="checkbox" name="orderdetailidx" value="<%= ocsOrderDetail.FItemList(i).Forgdetailidx %>" onClick="AnCheckClick(this); CheckSelect(this);" <%= strhtmldisabled %>>
					</td>
					<td><font color="<%= ocsOrderDetail.FItemList(i).CancelStateColor %>"><%= ocsOrderDetail.FItemList(i).CancelYnName %></font></td>
					<td>
						<%=ocsOrderDetail.FItemList(i).fitemgubun%>-<%=CHKIIF(ocsOrderDetail.FItemList(i).fitemid>=1000000,Format00(8,ocsOrderDetail.FItemList(i).fitemid),Format00(6,ocsOrderDetail.FItemList(i).fitemid))%>-<%=ocsOrderDetail.FItemList(i).fitemoption%>
					</td>
					<td width="90">
						<acronym title="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
							<%= Left(ocsOrderDetail.FItemList(i).Fmakerid,32) %>
						</acronym>
					</td>
					<td align="left">
						<acronym title="<%= ocsOrderDetail.FItemList(i).FItemName %>"><%= DDotFormat(ocsOrderDetail.FItemList(i).FItemName,64) %></acronym>
						<% if (ocsOrderDetail.FItemList(i).FItemoptionName <> "") then %>
							<br><font color="blue">[<%= ocsOrderDetail.FItemList(i).FItemoptionName %>]</font><br>
						<% end if %>
						<div id="causepop_<%= distinctid %>" style="position:absolute;"></div>
					</td>
					<td>
						<input type="text" name="regitemno" onKeyUp="CheckMaxItemNo(this, <%= ocsOrderDetail.FItemList(i).FItemNo %>);" <%= strhtmlmodify %> value="<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo_off(IsStatusRegister) %>" size="2" style="text-align:center">
						/
						<input type="text" name="itemno" value="<%= ocsOrderDetail.FItemList(i).FItemNo %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
					</td>
					<input type="hidden" name="sellprice" value="<%= ocsOrderDetail.FItemList(i).fsellprice %>">
					<% if (IsItemCanceled) then %>
						<td align="right"><font color="gray"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font></td>
					<% elseif (ocsOrderDetail.FItemList(i).FItemNo < 1) then %>
						<td align="right"><font color="red"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font></td>
					<% else %>
					<td align="right">
						<font color="blue"><%= FormatNumber(ocsOrderDetail.FItemList(i).fsellprice,0) %></font>
					</td>
					<% end if %>
					<td align="center">
						<% if (IsStatusFinished) and ((divcd="A010") or (divcd="A011")) then %>
							<br><input type="button" class="button" value="불량등록" onClick="popBadItemReg('10<%= Format00(6,ocsOrderDetail.FItemList(i).FItemid) %><%= ocsOrderDetail.FItemList(i).FItemOption %>','<%= ocsOrderDetail.FItemList(i).GetDefaultRegNo(IsStatusRegister) %>');">
						<% end if %>
						<% if ocsOrderDetail.FItemList(i).fcancelorgdetailidx <> "" then %>
							취소상세idx : <%=ocsOrderDetail.FItemList(i).fcancelorgdetailidx%>
						<% end if %>
					</td>
					<input type="hidden" name="makerid" value="<%= ocsOrderDetail.FItemList(i).Fmakerid %>">
					<input type="hidden" name="odlvtype" value="<%= ocsOrderDetail.FItemList(i).Fodlvtype %>">
					<input type="hidden" name="dummystopper" value="">
				</tr>
				<% next %>
            	</table>
            </td>
		</tr>
		</table>
	</td>
</tr>
<% end if %>