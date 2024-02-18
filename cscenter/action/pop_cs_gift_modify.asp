<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%

function removeSpace(inputstr)
	removeSpace = inputstr

	removeSpace = Replace(removeSpace, " ", "")
	removeSpace = Replace(removeSpace, Chr(9), "")
	removeSpace = Replace(removeSpace, vbCr, "")
	removeSpace = Replace(removeSpace, vbLf, "")
end function


dim orderserial, mode, itemlist

orderserial		= requestCheckVar(request("orderserial"), 32)
mode 			= requestCheckVar(request("mode"), 32)
itemlist 		= removeSpace(request("itemlist"))


dim oGift, oGiftModi
set oGift = new COrderGift

select case mode
	case "chk"
		oGift.FRectOrderSerial = orderserial
		oGift.GetOneOrderGiftlist
	case else
		response.end
end select

dim i, j, k
dim IsGiftOK

dim sqlStr

%>
<script>
function jsDelGift(frm, gift_code) {
	if (confirm("사은품 삭제하시겠습니까?") != true) { return false; }

	frm.mode.value = "del";
	frm.gift_code.value = gift_code;
	frm.submit();
}

function jsModiGift(frm, gift_code, modi_gift_code, modi_giftkind_code) {
	if (confirm("사은품을 변경합니다.\n\n진행하시겠습니까?") != true) { return false; }

	frm.mode.value = "modi";
	frm.gift_code.value = gift_code;
	frm.modi_gift_code.value = modi_gift_code;
	frm.modi_giftkind_code.value = modi_giftkind_code;
	frm.submit();
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
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>사은품 변경하기</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="post" onSubmit="return false;">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
		<td>취소후<br />사은품조건<br />충족여부</td>
		<td>비고</td>
	</tr>
	<%
	for i=0 to oGift.FResultCount -1
		'// 1:전체고객
		'// 2:이벤트 등록상품구매고객
		'// 3:특정 브랜드상품 구매고객
		'// 4:이벤트 그룹상품 구매고객
		'// 5:특정상품 구매고객
		'// 9:다이어리 샵상품 구매고객
		IsGiftOK = False
		if (oGift.FItemList(i).Fgift_scope <> "1") and (oGift.FItemList(i).Fgift_scope <> "2") and (oGift.FItemList(i).Fgift_scope <> "3") and (oGift.FItemList(i).Fgift_scope <> "4") and (oGift.FItemList(i).Fgift_scope <> "5") and (oGift.FItemList(i).Fgift_scope <> "9") then
			response.write "시스템 에러"
			response.end
		end if

		sqlStr = " exec [db_order].[dbo].[sp_Ten_order_gift_chkValid_CS] '" & orderserial & "', " & oGift.FItemList(i).Fgift_scope & ", " & oGift.FItemList(i).Fgift_code & ", '" & itemlist & "' "
        if (orderserial = "20100547132") then
            response.write "<!-- " & sqlStr & " -->"
        end if
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
			if (rsget("validYN") = 1) then
				IsGiftOK = True
			end if
		end if
		rsget.close

		set oGiftModi = Nothing
		set oGiftModi = new COrderGift
		if oGift.FItemList(i).Fevt_code <> 0 then
			oGiftModi.FRectOrderSerial = orderserial
			oGiftModi.FRectGiftScope = oGift.FItemList(i).Fgift_scope
			oGiftModi.FRectGiftCode = oGift.FItemList(i).Fgift_code
			oGiftModi.FRectEvtCode = oGift.FItemList(i).Fevt_code
			oGiftModi.FRectItemListArr = itemlist
			oGiftModi.GetOneOrderValidGiftlist
		end if
	%>
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
		<td>
			<%= CHKIIF(IsGiftOK=True, "충족", "<font color='red'>충족안함</font>") %>
		</td>
		<td>
			<% if (IsGiftOK<>True) then %>
			<input type="button" class="button" value="삭제하기" onclick="jsDelGift(frmAct, <%= oGift.FItemList(i).Fgift_code %>);">
			<% end if %>
		</td>
	</tr>
	<% if (oGiftModi.FResultCount>0) then %>
	<% for j=0 to oGiftModi.FResultCount - 1 %>
	<tr height="60" align="center" bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td>
			<%
			select case oGiftModi.FItemList(j).Fgift_type
				case "1"
					response.write "없음"
				case "2"
					response.write "구매금액<br />"
					if (oGiftModi.FItemList(j).Fgift_range2 = 0) then
						response.write CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range1,0)) + " 원 이상"
					else
						response.write CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range1,0)) + "~" + CStr(FormatNumber(oGiftModi.FItemList(j).Fgift_range2,0)) + " 원"
					end if
				case "3"
					response.write "구매수량<br />"
					if (oGiftModi.FItemList(j).Fgift_range2 = 0) then
						response.write CStr(oGiftModi.FItemList(j).Fgift_range1) + " 개 이상"
					else
						response.write CStr(oGiftModi.FItemList(j).Fgift_range1) + "~" + CStr(oGiftModi.FItemList(j).Fgift_range2) + " 개"
					end if
				case else
					response.write "ERR"
			end select
			%>
		</td>
		<td>
			<%= oGiftModi.FItemList(j).Fgiftkind_name %>
		</td>
		<td></td>
		<td></td>
		<td><%= oGiftModi.FItemList(j).FvalidStr %></td>
		<td>
			<% if (oGiftModi.FItemList(j).FvalidStr = "OK") then %>
			<input type="button" class="button" value="변경하기" onclick="jsModiGift(frmAct, <%= oGift.FItemList(i).Fgift_code %>, <%= oGiftModi.FItemList(j).Fgift_code %>, <%= oGiftModi.FItemList(j).Fgiftkind_code %>);">
			<% end if %>
		</td>
	</tr>
	<% next %>
	<% end if %>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->

<form name="frmAct" method="post" onSubmit="return false;" action="pop_cs_gift_modify_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="gift_code" value="">
	<input type="hidden" name="modi_gift_code" value="">
	<input type="hidden" name="modi_giftkind_code" value="">
	<input type="hidden" name="itemlist" value="<%= itemlist %>">
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
