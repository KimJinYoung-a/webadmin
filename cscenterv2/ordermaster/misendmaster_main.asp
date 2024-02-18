<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno
dim didx, mode
didx = requestCheckVar(request("didx"),10)
mode = requestCheckVar(request("mode"),16)

deliveryno = requestCheckVar(request("deliveryno"),16)
orderserial = requestCheckVar(request("orderserial"),16)

if (orderserial = "") then
    orderserial = "-"
end if

dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.FRectDeliveryNo = deliveryno
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

dim omisendList
set omisendList = new COldMiSend
omisendList.FRectOrderSerial = orderserial
omisendList.GetMiSendOrderDetailList

dim i
%>
<script language='javascript'>
function confirmSubmit(){
    if (confirm('저장 하시겠습니까?')){
        document.frmmisend.submit();
    }
}

function popMisendInput(iidx){
    var popwin = window.open('popMisendInput.asp?idx=' + iidx,'popMisendInput','width=440,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSendCallChange(iidx){
    if (confirm('고객님께 안내전화를 드렸습니까?')){
        frmmisendOne.detailIDx.value=iidx;
        frmmisendOne.submit();
    }
}

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

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frmsearch>
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			주문번호 : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=13 >
        	<% if omasterwithcs.FOneItem.FCancelyn<>"N" then %>
			<b><font color="#CC3333">[취소주문]</font></b>
			<script language='javascript'>alert('취소된 거래 입니다.');</script>
			<% else %>
			[정상주문]
			<% end if %>
			&nbsp;
			&nbsp;
			고객명 : <%= omasterwithcs.FOneItem.FBuyName %>
			&nbsp;
			핸드폰번호 : <%= omasterwithcs.FOneItem.FBuyHp %>
			&nbsp;
			이메일 : <%= omasterwithcs.FOneItem.FBuyEmail %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SearchThis();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="처리내용저장" onclick="confirmSubmit();" disabled>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드</td>
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션]</font></td>
		<td width="30">주문<br>수량</td>
		<td width="30">부족<br>수량</td>

		<td width="30">취소<br>삭제</td>
		<td width="80">출고기준일</td>
		<td width="30">소요<br>일수</td>
		<td width="60">진행상태</td>
		<td width="100">미출고사유</td>
		<td width="80">출고예정일</td>
		<td width="120">물류/업체<br>작성메모</td>
		<td width="35">SMS</td>
		<td width="35">MAIL</td>
		<td width="35">CALL</td>
		<td width="100">CS처리구분</td>
		<td width="85">CS처리메모</td>

	</tr>
	<form name="frmmisend" method="post" action="misendmaster_main_process.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">
	<% for i=0 to omisendList.FResultCount -1 %>

	<% if omisendList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td><%= omisendList.FItemList(i).FMakerid %></td>
		<td>
			<% if omisendList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%= omisendList.FItemList(i).FItemID %></font>
			<% else %>
			<%= omisendList.FItemList(i).FItemID %>
			<% end if %>
		</td>
		<td><img src="<%= omisendList.FItemList(i).FSmallImage %>" width="50" height="50"></td>
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
		    <% if (Not IsNULL(omisendList.FItemList(i).getBeasongDPlusDate)) and (omisendList.FItemList(i).getBeasongDPlusDate<>"")  then %>
    		    <% if (omisendList.FItemList(i).getBeasongDPlusDate>=2) then %>
    		    <strong><font color="Red"><%= omisendList.FItemList(i).getBeasongDPlusDateStr %></font></strong>
    		    <% else %>
    		    <%= omisendList.FItemList(i).getBeasongDPlusDateStr %>
    		    <% end if %>
		    <% else %>
		    <%= omisendList.FItemList(i).getBeasongDPlusDateStr %>
		    <% end if %>
		</td>
		<td>
		    <font color="<%= omisendList.FItemList(i).getUpCheDeliverStateColor %>"><%= omisendList.FItemList(i).getUpCheDeliverStateName %></font>
		</td>
		<td>
			<% if (Trim(omisendList.FItemList(i).FPrevMisendReason) <> "") then %>
				<%= MiSendCodeToName(omisendList.FItemList(i).FPrevMisendReason) %><br>
				-&gt;
			<% end if %>
			<% if Not IsNull(omisendList.FItemList(i).FMisendReason) and (CStr(omisendList.FItemList(i).FIDx)=Cstr(didx)) then %>
				<font color="red">입력중</font>
			<% else %>
				<font color="<%= omisendList.FItemList(i).getMiSendCodeColor %>"><%= omisendList.FItemList(i).getMiSendCodeName %></font>
				<% if (omisendList.FItemList(i).FMisendReason = "05") then %>
				<br><acronym title="<%= omisendList.FItemList(i).FMiRegDate %>"><%= omisendList.FItemList(i).FMiRegUserid %></acrpnym>
				<% end if %>
			<% end if %>
		</td>
		<td>
			<% if (omisendList.FItemList(i).FMisendReason<>"") and (omisendList.FItemList(i).FMisendReason<>"00") and (omisendList.FItemList(i).FMisendReason<>"05") then %>
				<%= omisendList.FItemList(i).FmiSendIpgodate %>
			<% end if %>
		</td>
		<td>
			<%= omisendList.FItemList(i).FrequestString %>
			<%= nl2br(omisendList.FItemList(i).FupcheRequestString) %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendSMS %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
<!--
    	    <% if omisendList.FItemList(i).FMisendReason = "05" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[텐바이텐 품절안내]주문하신 상품중 <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)상품이 품절되어 발송이 불가합니다.쇼핑에 불편을 드려 죄송합니다');">N</a>
	    	<% elseif omisendList.FItemList(i).FMisendReason = "03" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[텐바이텐 출고지연안내]주문하신 상품중 <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)상품이 <%= omisendList.FItemList(i).FmiSendIpgodate %>에 발송될 예정입니다');">N</a>
	    	<% elseif omisendList.FItemList(i).FMisendReason = "01" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[텐바이텐 출고지연안내]주문하신 상품중 <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)상품이 <%= omisendList.FItemList(i).FmiSendIpgodate %>에 발송될 예정입니다');"> N </a>
	    	<% else %>
	    	    N
	    	<% end if %>
 -->
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendEmail %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>

<!--
    			<% if omisendList.FItemList(i).FMisendReason = "05" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');">N</a>
    	    	<% elseif omisendList.FItemList(i).FMisendReason = "03" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');">N</a>
    	    	<% elseif omisendList.FItemList(i).FMisendReason = "01" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');"> N </a>
    	    	<% else %>
    	    	    N
    	    	<% end if %>
-->
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <% if (omisendList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= omisendList.FItemList(i).FisSendCall %>
		        <% end if %>

    	    <% end if %>
<!--
		    <% if omisendList.FItemList(i).FisSendCall="Y" then %>
		        <%= omisendList.FItemList(i).FisSendCall %>
		    <% else %>
    			<% if (omisendList.FItemList(i).FMisendReason<>"") and (omisendList.FItemList(i).FMisendReason<>"00") then %>
    			N
    			<% end if %>
    		<% end if %>
  -->
		</td>

		<% if (omisendList.FItemList(i).FMisendReason <> "") then %>
		<input type="hidden" name="didx" value="<%= omisendList.FItemList(i).FIDx %>">
		<% end if %>

		<td>
		<% if (omisendList.FItemList(i).FMisendReason <> "") then %>

			<input type=hidden name=prevstate value="<%= omisendList.FItemList(i).FmiSendState %>">

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
	</form>
</table>

<form name="frmmisendOne" method="post" action="misendmaster_main_process.asp">
<input type="hidden" name="mode" value="SendCallChange">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="detailIDx" value="">
</form>
<!-- 표 하단바 끝-->


<%
set omasterwithcs = Nothing
set omisendList = Nothing
%>

<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
