<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 미출고 상품리스트
' Hieditor : 이상구 생성
'			 2019.01.16 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno, detailcancelyn, didx, mode, reload, i
dim sellsite
	didx = requestCheckvar(request("didx"),20)
	mode = request("mode")
	detailcancelyn = requestCheckvar(request("detailcancelyn"),2)
	reload = requestCheckvar(request("reload"),2)
	deliveryno = requestCheckvar(request("deliveryno"),30)
	orderserial = requestCheckvar(request("orderserial"),20)

if (orderserial = "") then
    orderserial = "-"
end if
if reload="" and detailcancelyn="" then detailcancelyn="Y"

dim omasterwithcs
set omasterwithcs = new COldMiSend
	omasterwithcs.FRectOrderSerial = orderserial
	omasterwithcs.FRectDeliveryNo = deliveryno
	omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial
sellsite = omasterwithcs.FOneItem.Fsitename

dim omisendList
set omisendList = new COldMiSend
	omisendList.FRectOrderSerial = orderserial
	omisendList.frectdetailcancelyn = detailcancelyn
	omisendList.GetMiSendOrderDetailList


'// 적용가능 API
dim availApiCS : availApiCS = "stockout"
if (sellsite = "coupang") or (sellsite = "interpark") then
	availApiCS = "cancel"
end if

%>
<script type="text/javascript">

function confirmSubmit(){
    if (confirm('저장 하시겠습니까?')){
        document.frmmisend.submit();
    }
}

function popMisendInput(iidx){
    var popwin = window.open('/partner/jumunmaster/popMisendInput.asp?idx=' + iidx,'popMisendInput','width=1200,height=800,scrollbars=yes,resizable=yes');
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
	frmsearch.submit();
}

function jsSendStockOut(detailidx) {
	var sellsite = '<%= sellsite %>';
    var orderserial = '<%= orderserial %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case '11st1010':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'nvstorefarm':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
        case 'gmarket1010':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'WMP':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'wmpfashion':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		default:
			alert('지원하지 않는 제휴몰입니다.[' + sellsite + ']');
			return;
	}

	if (confirm('결품등록 : 진행하시겠습니까?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendStockCnclOut(detailidx) {
	var sellsite = '<%= sellsite %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutCnclOne&detailidx=" + detailidx;
			break;
		//case 'lotteCom':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutCnclOne&detailidx=" + detailidx;
		//	break;
		default:
			alert('지원하지 않는 제휴몰입니다.[' + sellsite + ']');
			return;
	}

	if (confirm('결품등록취소 : 진행하시겠습니까?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendStockOutAll() {
	var sellsite = '<%= sellsite %>';
	var orderserial = '<%= orderserial %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutAll&orderserial=" + orderserial;
			break;
		//case 'lotteCom':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutAll&orderserial=" + orderserial;
		//	break;
		default:
			alert('지원하지 않는 제휴몰입니다.[' + sellsite + ']');
			return;
	}

	if (confirm('결품등록 : 진행하시겠습니까?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendCancelAll() {
	var sellsite = '<%= sellsite %>';
	var orderserial = '<%= orderserial %>';
	var url;
	var arrdetailidx = "";
	var arritemno = "";
	var chk, orderitemno, itemno, i, j, k;

	for (i = 0; ; i++) {
		chk = document.getElementById('chk_' + i);
		orderitemno = document.getElementById('orderitemno_' + i);
		itemno = document.getElementById('itemno_' + i);

		if (chk == undefined) { break; }
		if (chk.disabled == true) { continue; }
		if (chk.checked != true) { continue; }

		if ((itemno.value == "") || (itemno.value*0 != 0)) {
			alert('취소수량은 숫자만 가능합니다.');
			itemno.focus();
			return;
		}
		if (itemno.value*1 <= 0) {
			alert('취소수량은 0보다 커야합니다.');
			itemno.focus();
			return;
		}
		if (itemno.value*1 > orderitemno.value*1) {
			alert('취소수량은 주문수량보다 클 수 없습니다.');
			itemno.focus();
			return;
		}

		arrdetailidx = arrdetailidx + ',' + chk.value;
		arritemno = arritemno + ',' + itemno.value;
	}

	if (arrdetailidx == "") {
		alert('선택된 상품이 없습니다.');
		return;
	}

	switch (sellsite) {
		case 'coupang':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
			break;
		case 'interpark':
			if (confirm('인터파크는 선택수량에 무관하게 선택상품 전부가 취소됩니다.\n\n진행하시겠습니까?')) {
				url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
			} else {
				return;
			}

			break;
		//case 'cjmall':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
		//	break;
		default:
			alert('지원하지 않는 제휴몰입니다.[' + sellsite + ']');
			return;
	}

	if (confirm('주문취소 : 진행하시겠습니까?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
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
<form name="frmsearch" style="margin:0px;" method="get">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>" height="25">검색<br>조건</td>
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

			<% if C_CriticInfoUserLV1 then %>
				&nbsp;
				핸드폰번호 : <%= omasterwithcs.FOneItem.FBuyHp %>
				&nbsp;
				이메일 : <%= omasterwithcs.FOneItem.FBuyEmail %>
		    <% else %>
				&nbsp;
				핸드폰번호 : XXX-XXX-XXXX
				&nbsp;
				이메일 :
			<% end if %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="SearchThis();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="25">
			<input type="checkbox" value="Y" name="detailcancelyn" <% if detailcancelyn="Y" then response.write " checked" %> > 취소주문제외
		</td>
	</tr>
</table>
</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px; padding-bottom:10px;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="처리내용저장" onclick="confirmSubmit();">
			<% if (sellsite<>"10x10") then %>
			&nbsp;
			제휴몰 : <%= sellsite %>
            <%
            select case sellsite
                case "ssg"
                    response.write "(상품별 or 일괄품절 전송가능, 상품별 품절등록취소 전송가능)"
                case "11st1010"
                    response.write "(상품별 품절 전송가능)"
                case "nvstorefarm"
                    response.write "(상품별 취소신청승인/상품주문취소 전송가능)"
                case "coupang"
                    response.write "(선택상품 제휴몰 취소전송 가능, 수량 일부취소 가능)"
                case "interpark"
                    response.write "(선택상품 제휴몰 취소전송 가능, 수량 일부취소 가능)"
                case "gmarket1010"
                    response.write "(선택상품 제휴몰 취소전송 가능, 수량전부취소)"
                case "WMP", "wmpfashion"
                    response.write "(선택상품 제휴몰 취소전송 가능, 수량전부취소)"
                case else
                    response.write "(API 작업이전)"
            end select
            %>
			<% end if %>
		</td>
		<td align="right">
			<% if (sellsite<>"10x10") then %>
			<% if (availApiCS = "stockout") then %>
			<input type="button" class="csbutton" value="품절정보 제휴몰 일괄전송" onClick="jsSendStockOutAll()" <%= CHKIIF(C_ADMIN_AUTH, "", "disabled") %>>
			<% end if %>
			<% if (availApiCS = "cancel") then %>
			<input type="button" class="csbutton" value="선택상품 제휴몰 취소전송" onClick="jsSendCancelAll()">
			<% end if %>
			<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
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
		<td width="35">제휴<br />API</td>
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
		<td>
			<% if (sellsite<>"10x10") and (availApiCS = "cancel") then %>
			<input type="checkbox" name="chk" id="chk_<%= i %>" value="<%= omisendList.FItemList(i).Fidx %>" <%= CHKIIF(omisendList.FItemList(i).FMisendReason="05", "", "disabled")%>>
			<% end if %>
		</td>
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
		<td>
			<% if omisendList.FItemList(i).FItemLackNo=0 then %>
			-
			<% else %>
			<input type="hidden" name="orderitemno" id="orderitemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemNo %>">
			<input type="text" class="text" name="itemno" id="itemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemLackNo %>" size="1">
			<% end if %>
		</td>
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
		    <%
'				If (Not IsNULL(omisendList.FItemList(i).getBeasongDPlusDate)) and (omisendList.FItemList(i).getBeasongDPlusDate<>"")  then
'					if (omisendList.FItemList(i).getBeasongDPlusDate>=2) then
'						response.write "<strong><font color='Red'>"& omisendList.FItemList(i).getBeasongDPlusDateStr &"</font></strong>"
'					else
'					response.write omisendList.FItemList(i).getBeasongDPlusDateStr
'			   		end if
'				else
'					response.write omisendList.FItemList(i).getBeasongDPlusDateStr
'				end if

				If (Not IsNULL(omisendList.FItemList(i).FDday)) and (omisendList.FItemList(i).FDday<>"")  then
					if (omisendList.FItemList(i).FDday>=2) then
						response.write "<strong><font color='Red'>"& omisendList.FItemList(i).getNewBeasongDPlusDateStr &"</font></strong>"
					else
    		    		response.write omisendList.FItemList(i).getNewBeasongDPlusDateStr
			   		end if
				else
					response.write omisendList.FItemList(i).getNewBeasongDPlusDateStr
				end if
			%>
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
				<% if True or (omisendList.FItemList(i).FMisendReason = "05") then %>
				<br><acronym title="<%= omisendList.FItemList(i).FMiRegDate %>"><%= omisendList.FItemList(i).FMiRegUserid %></acrpnym>
				<% end if %>
			<% end if %>
			<% if Not IsNull(omisendList.FItemList(i).Freqreguserid) then %>
				<br /><%= omisendList.FItemList(i).Freqreguserid %>
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
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="N") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendEmail %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="N") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <% if (omisendList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= omisendList.FItemList(i).FisSendCall %>
		        <% end if %>
    	    <% end if %>
		</td>
		<td>
			<% if (sellsite<>"10x10") then %>
				<%= omisendList.FItemList(i).FisSendAPI %><br />
				<% if (omisendList.FItemList(i).FMisendReason="05") and (omisendList.FItemList(i).FisSendAPI = "N") then %>
				<input type="button" class="csbutton" value="전송" onClick="jsSendStockOut(<%= omisendList.FItemList(i).Fidx %>)">
				<% elseif (omisendList.FItemList(i).FisSendAPI = "Y") then %>
				<input type="button" class="csbutton" value="취소" onClick="jsSendStockCnclOut(<%= omisendList.FItemList(i).Fidx %>)">
                <% elseif (orderserial = "20030XXX942396") then %>
                <input type="button" class="csbutton" value="전송" onClick="jsSendStockOut(<%= omisendList.FItemList(i).Fidx %>)">
				<% end if %>
			<% end if %>
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

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
