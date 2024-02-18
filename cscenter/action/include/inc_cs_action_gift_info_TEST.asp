<% if (IsDisplayGift = True) then %>
<tr >
    <td >
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#F4F4F4">
            <td bgcolor="<%= adminColor("topbar") %>" align="center" width="80">
				사은품<br />
				<input type="button" class="button" value="체크하기" onClick="popChkGiftItem()">
				<input type="hidden" id="evt_chk_need" value="<%= CHKIIF(divcd="A008" and oGift.FResultCount>0 and IsStatusRegister, "Y", "N") %>">
			</td>
			<td>
                <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr height="20" align="center" bgcolor="#F4F4F4">
					<td width="50">배송<br />구분</td>
					<td>기프트<br />코드</td>
					<td>이벤트<br />코드</td>
					<td>이벤트명</td>
					<td width="70">이벤트<br />시작일</td>
					<td width="70">이벤트<br />종료일</td>
					<td>증정 대상</td>
					<td>증정 조건</td>
					<td>사은품</td>
					<td>수량</td>
					<td>한정수량</td>
				</tr>
				<% for i=0 to oGift.FResultCount -1 %>
				<input type="hidden" id="evt_code_<%= i %>" value="<%= oGift.FItemList(i).Fevt_code %>">
				<input type="hidden" id="evt_startdate_<%= i %>" value="<%= oGift.FItemList(i).Fevt_startdate %>">
				<input type="hidden" id="evt_enddate_<%= i %>" value="<%= oGift.FItemList(i).Fevt_enddate %>">
				<input type="hidden" id="evt_gift_scope_<%= i %>" value="<%= oGift.FItemList(i).Fgift_scope %>">
				<input type="hidden" id="evt_gift_type_<%= i %>" value="<%= oGift.FItemList(i).Fgift_type %>">
				<input type="hidden" id="evt_gift_range1_<%= i %>" value="<%= oGift.FItemList(i).Fgift_range1 %>">
				<input type="hidden" id="evt_gift_range2_<%= i %>" value="<%= oGift.FItemList(i).Fgift_range2 %>">
				<tr height="60" align="center" bgcolor="#FFFFFF">
					<td>
						<% if oGift.FItemList(i).Fisupchebeasong="Y" then %>
	                    <font color="red">업체</font>
	                    <% else %>
	                    <font color="blue">텐배</font>
	                    <% end if %>
					</td>
					<td><%= oGift.FItemList(i).Fgift_code %></td>
					<td><%= oGift.FItemList(i).Fevt_code %></td>
					<td>
						<% if (oGift.FItemList(i).Fevt_code<>0) then %>
		                <a target="_blank" href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oGift.FItemList(i).Fevt_code %>"><font color="blue"><%= oGift.FItemList(i).Fevt_name %></font></a>
		                <% end if %>
					</td>
					<td><%= oGift.FItemList(i).Fevt_startdate %></td>
					<td><%= oGift.FItemList(i).Fevt_enddate %></td>
					<td>
						<%
						select case oGift.FItemList(i).Fgift_scope
							case "1"
								response.write "전체고객"
							case "2"
								response.write "이벤트<br />등록상품<br />구매고객"
							case "3"
								response.write "특정<br />브랜드상품<br />구매고객"
							case "4"
								response.write "이벤트<br />그룹상품<br />구매고객"
							case "5"
								response.write "특정상품<br />구매고객"
							case "9"
								response.write "다이어리<br />샵상품<br />구매고객"
							case else
								response.write "ERR"
						end select
						%>
					</td>
					<td>
						<%
						select case oGift.FItemList(i).Fgift_type
							case "1"
								response.write "없음"
							case "2"
								response.write "구매금액<br />"
								if (oGift.FItemList(i).Fgift_range2 = 0) then
									response.write CStr(FormatNumber(oGift.FItemList(i).Fgift_range1,0)) + " 원 이상"
								else
									response.write CStr(FormatNumber(oGift.FItemList(i).Fgift_range1,0)) + "~" + CStr(FormatNumber(oGift.FItemList(i).Fgift_range2,0)) + " 원"
								end if
							case "3"
								response.write "구매수량<br />"
								if (oGift.FItemList(i).Fgift_range2 = 0) then
									response.write CStr(oGift.FItemList(i).Fgift_range1) + " 개 이상"
								else
									response.write CStr(oGift.FItemList(i).Fgift_range1) + "~" + CStr(oGift.FItemList(i).Fgift_range2) + " 개"
								end if
							case else
								response.write "ERR"
						end select
						%>
					</td>
					<td>
						<%
						if Not IsNull(oGift.FItemList(i).Fchg_giftSTR) then
							if oGift.FItemList(i).Fchg_giftSTR <> "" then
								response.write oGift.FItemList(i).Fchg_giftSTR
							else
								response.write oGift.FItemList(i).getGiftName()
							end if
						else
							response.write oGift.FItemList(i).getGiftName()
						end if
						%>
					</td>
					<td>
						<%= oGift.FItemList(i).Fgiftkind_cnt %> 개
						<%
						select case oGift.FItemList(i).Fgiftkind_type
							case "2"
								response.write "<br />[1+1]"
							case "3"
								response.write "<br />[1:1]"
							case else
								'//
						end select
						%>
					</td>
					<td>
						<%
						if (oGift.FItemList(i).Fgiftkind_limit <> 0) and ((oGift.FItemList(i).Fgiftkind_limit - oGift.FItemList(i).Fgiftkind_givecnt) <= 100) then
							response.write (oGift.FItemList(i).Fgiftkind_limit - oGift.FItemList(i).Fgiftkind_givecnt) & " / " & oGift.FItemList(i).Fgiftkind_limit
						end if
						%>
					</td>
				</tr>
				<% next %>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% end if %>
