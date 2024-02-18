<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim masteridx,deliveryno ,i ,didx, mode ,omasterwithcs ,omisendList
	didx = requestCheckVar(request("didx"),10)
	mode = requestCheckVar(request("mode"),32)
	deliveryno = requestCheckVar(request("deliveryno"),32)
	masteridx = requestCheckVar(request("masteridx"),10)

	if (masteridx = "") then
	    masteridx = "-"
	end if

set omasterwithcs = new COrder
	omasterwithcs.FRectmasteridx = masteridx
	omasterwithcs.FRectDeliveryNo = deliveryno
	omasterwithcs.fGetOneOrderMasterWithCS

	masteridx = omasterwithcs.FOneItem.Fmasteridx

set omisendList = new COrder
	omisendList.FRectmasteridx = masteridx
	omisendList.fgetMiSendOrderDetailList
%>

<script language='javascript'>

function confirmSubmit(){
    if (confirm('저장 하시겠습니까?')){
        document.frmmisend.submit();
    }
}

function popMisendInput(detailidx){
    var popwin = window.open('/common/offshop/beasong/upche_popMisendInput.asp?detailidx=' + detailidx,'popMisendInput','width=600,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSendCallChange(detailidx){
    if (confirm('고객님께 안내전화를 드렸습니까?')){
        frmmisendOne.detailIDx.value=detailidx;
        frmmisendOne.submit();
    }
}

function SearchThis(){
	location.href="/admin/ordermaster/misendmaster_main.asp?masteridx=" + frmsearch.masteridx.value;
}

</script>
<style type="text/css">

td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "굴림", "돋움";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}

</style>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name=frmsearch>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		주문번호 : <input type="text" class="text" name="masteridx" value="<%= masteridx %>" size=13 >
    	<% if omasterwithcs.FOneItem.FCancelyn<>"N" then %>
			<b><font color="#CC3333">[취소주문]</font></b>
			<script language='javascript'>alert('취소된 거래 입니다.');</script>
		<% else %>
			[정상주문]
		<% end if %>
		고객명 : <%= omasterwithcs.FOneItem.FBuyName %>
		핸드폰번호 : <%= omasterwithcs.FOneItem.FBuyHp %>
		이메일 : <%= omasterwithcs.FOneItem.FBuyEmail %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SearchThis();">
	</td>
</tr>
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="csbutton" value="처리내용저장" onclick="confirmSubmit();">
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmmisend" method="post" action="misendmaster_process.asp">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="mode" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>상품명<font color="blue">[옵션]</font></td>
	<td >주문<br>수량</td>
	<td>부족<br>수량</td>
	<td>취소<br>삭제</td>
	<td>출고기준일</td>
	<td>소요<br>일수</td>
	<td>진행상태</td>
	<td>미출고사유</td>
	<td>출고예정일</td>
	<td>물류/업체<br>작성메모</td>
	<td>SMS</td>
	<td>MAIL</td>
	<td>CALL</td>
	<td>CS처리구분</td>
	<td>CS처리메모</td>

</tr>
<% if omisendList.FResultCount > 0 then %>
<% for i=0 to omisendList.FResultCount -1 %>

<% if omisendList.FItemList(i).FItemLackNo<>0 then %>
<tr align="center" bgcolor="<%= adminColor("pink") %>">
<% else %>
<tr align="center" bgcolor="FFFFFF">
<% end if %>
	<td>
		<% if omisendList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%=omisendList.FItemList(i).fitemgubun%>-<%=CHKIIF(omisendList.FItemList(i).fitemid>=1000000,Format00(8,omisendList.FItemList(i).fitemid),Format00(6,omisendList.FItemList(i).fitemid))%>-<%=omisendList.FItemList(i).fitemoption%></font>
		<% else %>
			<%=omisendList.FItemList(i).fitemgubun%>-<%=CHKIIF(omisendList.FItemList(i).fitemid>=1000000,Format00(8,omisendList.FItemList(i).fitemid),Format00(6,omisendList.FItemList(i).fitemid))%>-<%=omisendList.FItemList(i).fitemoption%>
		<% end if %>
	</td>
	<td align="left">
		<%= omisendList.FItemList(i).FItemName %>
		<% if omisendList.FItemList(i).FItemOptionName<>"" then %>
		<br>
		<font color="blue">[<%= omisendList.FItemList(i).FItemOptionName %>]</font>
		<% end if %>
	</td>
	<td><%= omisendList.FItemList(i).FItemNo %></td>
	<td><font color="red"><b><% if omisendList.FItemList(i).FItemLackNo=0 then response.write "-" else  response.write  omisendList.FItemList(i).FItemLackNo end if%></b></font></td>
	<td>
	    <%= fnColor(omisendList.FItemList(i).FDetailCancelYn,"cancelyn") %>
	</td>
	<td>
	    <% if IsNULL(omisendList.FItemList(i).FbaljuDate) then %>

	    <% else %>
	    <%= Left(omisendList.FItemList(i).FbaljuDate,10) %>
	    <% end if %>
	</td>
	<td>
	    <!-- D+2 이상일경우, 빨간색으로 표시 -->
	    <% if (Not IsNULL(omisendList.FItemList(i).getBeasongDPlusDate_off)) and (omisendList.FItemList(i).getBeasongDPlusDate_off<>"")  then %>
		    <% if (omisendList.FItemList(i).getBeasongDPlusDate_off>=2) then %>
		    	<strong><font color="Red"><%= omisendList.FItemList(i).getBeasongDPlusDateStr_off %></font></strong>
		    <% else %>
		    	<%= omisendList.FItemList(i).getBeasongDPlusDateStr_off %>
		    <% end if %>
	    <% else %>
	    	<%= omisendList.FItemList(i).getBeasongDPlusDateStr_off %>
	    <% end if %>
	</td>
	<td>
	    <font color="<%= omisendList.FItemList(i).GetStateColor %>"><%= omisendList.FItemList(i).GetStateName %></font>
	</td>
	<td>
		<% if Not IsNull(omisendList.FItemList(i).FMisendReason) and (CStr(omisendList.FItemList(i).fdetailidx)=Cstr(didx)) then %>
			<font color="red">입력중</font>
		<% else %>
			<font color="<%= omisendList.FItemList(i).getMiSendCodeColor_off %>"><%= omisendList.FItemList(i).getMiSendCodeName_off %></font>
		<% end if %>
	</td>
	<td>
		<% if (omisendList.FItemList(i).FMisendReason<>"") and (omisendList.FItemList(i).FMisendReason<>"00") and (omisendList.FItemList(i).FMisendReason<>"05") then %>
			<%= omisendList.FItemList(i).FmiSendIpgodate %>
		<% end if %>
	</td>
	<td><%= omisendList.FItemList(i).FrequestString %></td>
	<td>
	    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
	        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).fdetailidx %>')"><%= omisendList.FItemList(i).FisSendSMS %></a>
	    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
	        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).fdetailidx %>')">N</a>
	    <% end if %>
	</td>
	<td>
	    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
	        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).fdetailidx %>')"><%= omisendList.FItemList(i).FisSendEmail %></a>
	    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
	        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).fdetailidx %>')">N</a>
	    <% end if %>
	</td>
	<td>
	    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
	        <% if (omisendList.FItemList(i).FisSendCall<>"Y") then %>
	            <a href="javascript:popSendCallChange('<%= omisendList.FItemList(i).fdetailidx %>')"><%= omisendList.FItemList(i).FisSendCall %></a>
	        <% else %>
	            <%= omisendList.FItemList(i).FisSendCall %>
	        <% end if %>

	    <% end if %>
	</td>

	<% if (omisendList.FItemList(i).FMisendReason <> "") then %>
	<input type="hidden" name="didx" value="<%= omisendList.FItemList(i).fdetailidx %>">
	<% end if %>

	<td>
	  <% if (omisendList.FItemList(i).FMisendReason <> "") then %>
	      <% if (omisendList.FItemList(i).FmiSendState = "7") then %>
	      완료
	      <input type=hidden name=state value="7">
	      <% else %>
	  	<select class="select" name="state">
			<option value="0" <% if (omisendList.FItemList(i).FmiSendState = "0") then response.write "selected" end if %>>미처리</option>
			<!-- <option value="1" <% if (omisendList.FItemList(i).FmiSendState = "1") then response.write "selected" end if %>>SMS완료</option> -->
			<!-- <option value="2" <% if (omisendList.FItemList(i).FmiSendState = "2") then response.write "selected" end if %>>안내Mail완료</option> -->
			<!-- <option value="3" <% if (omisendList.FItemList(i).FmiSendState = "3") then response.write "selected" end if %>>통화완료</option> -->
			<!-- <option value="3" <% if (omisendList.FItemList(i).FmiSendState = "3") then response.write "selected" end if %>>배송실처리</option> -->
			<option value="4" <% if (omisendList.FItemList(i).FmiSendState = "4") then response.write "selected" end if %>>고객안내</option><!-- 신규(SMS/mail/통화시) -->
			<option value="6" <% if (omisendList.FItemList(i).FmiSendState = "6") then response.write "selected" end if %>>CS처리완료</option>
	  	</select>
	      <% end if %>
	  <% end if %>
	</td>
	<td>
	  <% if (omisendList.FItemList(i).FMisendReason <> "") then %>
	  <input type="text" class="text" name="finishstr" value="<%= omisendList.FItemList(i).FfinishString %>" size="10">
	  <% end if %>
	</td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</form>
<form name="frmmisendOne" method="post" action="/admin/offshop/cscenter/order/misendmaster_process.asp">
	<input type="hidden" name="mode" value="SendCallChange">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="detailIDx" value="">
</form>
</table>

<%
set omasterwithcs = Nothing
set omisendList = Nothing
%>
<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->