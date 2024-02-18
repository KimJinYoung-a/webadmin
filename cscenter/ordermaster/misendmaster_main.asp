<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno,obalju
dim didx, mode
didx = request("didx")
mode = request("mode")

deliveryno = request("deliveryno")
orderserial = request("orderserial")

if (orderserial = "") then
        orderserial = "-"
end if

dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.FRectDeliveryNo = deliveryno
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

set obalju = New CBalju
obalju.FRectOrderSerial = orderserial
obalju.GetMiSendOrderDetail

dim i
%>
<script language='javascript'>
function DelMiSend(frm){
	var ret = confirm('삭제하시겠습니까?');

	if (ret){
		frm.mode.value="del";
		frm.submit();
	}
}

function SaveMiSend(frm){
	if (frm.ipgodate.value.length>0){
		if (frm.ipgodate.value.length!=10){
			alert('입고예정일을 정확히 입력하세요.');
			frm.ipgodate.focus();
			return;
		}
	}

	var ret = confirm('저장하시겠습니까?');

	if (ret){
		frm.submit();
	}
}
function calender_open() {
       document.all.cal.style.display="";
}

function SearchThis(){
	location.href="/admin/ordermaster/misendmaster_main.asp?orderserial=" + frmsearch.orderserial.value;
}

</script>
<style type="text/css">
<!--
td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "굴림", "돋움";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name=frmsearch>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=13 >
	        	<input type="button" class="button_s" value="검색" onClick="SearchThis()">
	        	&nbsp;&nbsp;
	        	<% if omasterwithcs.FOneItem.FCancelyn="Y" then %>
				<b><font color="#CC3333">취소 주문건입니다.</font></b>
				<script language='javascript'>alert('취소된 거래 입니다.');</script>
				<% else %>
				정상 주문건입니다.
				<% end if %>
		</td>
	</tr>
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션]</font></td>
		<td width="40">수량</td>
		<td width="30">취소<br>삭제</td>
		<td width="80">출고기준일</td>
		<td width="30">D+</td>
		<td width="80">미출고사유</td>
		<td width="80">출고예정일</td>
		<td width="80">요청사항</td>
		<td width="80">처리결과</td>
		<td width="80">처리구분</td>
	</tr>
	<form name="frmmisend" method="post" action="domisendmaster_main.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">
	<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
	<tr align="center" bgcolor="FFFFFF">
		<% if obalju.FBaljuDetailList(i).IsUpcheBeasong then %>
		<td ><font color="red"><%= obalju.FBaljuDetailList(i).FItemID %></font></td>
		<% else %>
		<td ><%= obalju.FBaljuDetailList(i).FItemID %></td>
		<% end if %>
		<td><img src="<%= obalju.FBaljuDetailList(i).FImageSmall %>" width="50" height="50"></td>
		<td align="left">
			<%= obalju.FBaljuDetailList(i).FItemName %>
			<% if obalju.FBaljuDetailList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= obalju.FBaljuDetailList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= obalju.FBaljuDetailList(i).FItemNo %></td>
		<td><font color="<%= obalju.FBaljuDetailList(i).CancelYnColor %>"><%= obalju.FBaljuDetailList(i).CancelYnName %></font></td>
		<td>2009-04-05</td>
		<td>D+2</td> <!-- D+2 이상일경우, 빨간색으로 표시 -->
		<% if Not IsNull(obalju.FBaljuDetailList(i).FmiSendCode) and (CStr(obalju.FBaljuDetailList(i).FDetailIDx)=Cstr(didx)) then %>
		<td><font color="red">입력중</font></td>
		<% else %>
		<td><font color="<%= obalju.FBaljuDetailList(i).getMiSendCodeColor %>"><%= obalju.FBaljuDetailList(i).getMiSendCodeName %></font></td>
		<% end if %>
		<td><%= obalju.FBaljuDetailList(i).FmiSendIpgodate %></td>
		<td><%= obalju.FBaljuDetailList(i).FrequestString %></td>
		<% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		<input type="hidden" name="didx" value="<%= obalju.FBaljuDetailList(i).FDetailIDx %>">
		<% end if %>
		<td>
		  <% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		  <input type="text" name="finishstr" value="<%= obalju.FBaljuDetailList(i).FfinishString %>" size="10">
		  <% end if %>
		</td>
		<td>
		  <% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		      <% if (obalju.FBaljuDetailList(i).FmiSendState = "7") then %>
		      완료
		      <input type=hidden name=state value="7">
		      <% else %>
		  <select name="state">
		    <option value="0" <% if (obalju.FBaljuDetailList(i).FmiSendState = "0") then response.write "selected" end if %>>미처리</option>
		    <option value="1" <% if (obalju.FBaljuDetailList(i).FmiSendState = "1") then response.write "selected" end if %>>SMS완료</option>
		    <option value="2" <% if (obalju.FBaljuDetailList(i).FmiSendState = "2") then response.write "selected" end if %>>안내Mail완료</option>
		    <option value="3" <% if (obalju.FBaljuDetailList(i).FmiSendState = "3") then response.write "selected" end if %>>통화완료</option>
		   <!-- <option value="3" <% if (obalju.FBaljuDetailList(i).FmiSendState = "3") then response.write "selected" end if %>>배송실처리</option> -->
		    <option value="6" <% if (obalju.FBaljuDetailList(i).FmiSendState = "6") then response.write "selected" end if %>>CS처리완료</option>
		  </select>
		      <% end if %>
		  <% end if %>
		</td>
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value=" 처리입력 " onclick="document.frmmisend.submit();"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- 표 하단바 끝-->


<%
set omasterwithcs = Nothing
set obalju = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->