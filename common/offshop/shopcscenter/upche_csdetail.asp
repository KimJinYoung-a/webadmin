<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 cs내역
' Hieditor : 2011.03.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim csmasteridx ,ioneas,i ,ioneasDetail ,deliverydivname
	csmasteridx = request("csmasteridx")

set ioneas = new corder
	ioneas.FRectMakerID = session("ssBctID")
	ioneas.FRectCsAsID = csmasteridx
	ioneas.fGetOneCSASMaster

if (ioneas.FResultCount<1) then
    response.write "<script>"
    response.write "	alert('유효한 접수번호가 아닙니다.');"
    response.write "	history.back();"
    response.write "</script>"
    response.write dbget.close()	:	response.End
end if

set ioneasDetail= new corder
	ioneasDetail.FRectCsAsID = csmasteridx
	ioneasDetail.fGetCsDetailList

if ioneas.FOneItem.Fdivcd = "A030" then
	deliverydivname = "A/S완료후수령지"
elseif ioneas.FOneItem.Fdivcd = "A031" then
	deliverydivname = "A/S업체"
end if
%>

<script language='javascript'>

function ViewOrderDetail(frm,orgmasteridx){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "ViewOrderDetail", props);

    frm.target = 'ViewOrderDetail';
    frm.masteridx.value = orgmasteridx;
    frm.action="/common/offshop/shopcscenter/upche_viewordermaster.asp";
	frm.submit();
}

function SaveFin(frm){
	//alert('잠시 준비중입니다.');
	//return;

	if (frm.finishmemo.value.length<1){
		alert('처리 내용을 입력해 주세요.');
		frm.finishmemo.focus();
		return;
	}

	if (frm.songjangdiv.value.length<1){
		alert('운송장 택배사를 입력해 주세요.');
		frm.songjangdiv.focus();
		return;
	}

	if (frm.songjangno.value.length<1){
		alert('운송장 번호를 입력해 주세요.');
		frm.songjangno.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

//업체a/s , 업체a/s(매장회수) 주소지 변경
function popEditCsDelivery(CsAsID){
    var window_width = 600;
    var window_height = 450;

    var popEditCsDelivery=window.open("/admin/offshop/shopcscenter/action/pop_CsDeliveryEdit.asp?CsAsID=" + CsAsID ,"popEditCsDelivery","width=600 height=500 scrollbars=yes resizable=yes");
    popEditCsDelivery.focus();
}

</script>

<% if ioneas.Ftotalcount > 0 then %>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="FFFFFF">
	<tr>
		<td>
			<% getcurrstate_table ioneas.FOneItem.Fcurrstate,ioneas.FOneItem.Fdivcd %>
		</td>
	</tr>
	</table>
<% end if %>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="upche_csprocess.asp">
<input type="hidden" name="orgmasteridx" value="<%= ioneas.FOneItem.forgmasteridx %>">
<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="masteridx" value="<%= ioneas.FOneItem.fmasteridx %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>배송CS 처리답변</b>
		&nbsp;&nbsp;
    	작성일 : <b><%= ioneas.FOneItem.Fregdate %></b>
    	&nbsp;&nbsp;
    	<% if not IsNULL(ioneas.FOneItem.Ffinishdate) then %>
    	완료일 : <b><%= ioneas.FOneItem.Ffinishdate %></b>
    	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td>
		<%= ioneas.FOneItem.Forderno %>
		<input type="button" class="button" value="주문상세보기" onclick="ViewOrderDetail(frmshow,'<%= ioneas.FOneItem.forgmasteridx %>');">
	</td>
	<td width="45%" rowspan="5" valign="top">
		<%
		if ioneas.FOneItem.Fdivcd="A030" then
		%>
			* 업체A/S
			<br><br>업체에서 a/s 상품을 수리 완료후, 매장이나 고객분께 발송하는 내역 입니다.
			<br>매장이나 고객분께 발송하신 택배 송장 내역을 입력해 주세요.
		<%
		elseif ioneas.FOneItem.Fdivcd="A031" then
		%>
			* 업체A/S(매장회수)
			<br><br>매장에서 고객분께 접수한 a/s 상품을, 업체로 발송하는 내역 입니다.
		<% end if %>

	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">고객명</td>
	<td><%= ioneas.FOneItem.FCustomerName %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td><%= ioneas.FOneItem.FTitle %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">접수내용</td>
	<td><%= replace(ioneas.FOneItem.Fcontents_jupsu,VbCrlf,"<br>") %></td>
</tr>
<% if (ioneasDetail.FResultCount>0) then %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>">접수상품</td>
    <td>
        <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a">
        <tr bgcolor="<%= adminColor("topbar") %>" align="center">
            <td width="100">상품코드</td>
            <td width="100">브랜드ID</td>
            <td>상품명<font color="blue">[옵션명]</font></td>
            <td width="50">판매가</td>
            <td width="40">수량</td>
        </tr>
        <% for i=0 to ioneasDetail.FResultCount-1 %>
        <tr bgcolor="#FFFFFF" align="center">
            <td>
            	<%=ioneasDetail.FItemList(i).fitemgubun%>-<%=FormatCode(ioneasDetail.FItemList(i).fitemid)%>-<%=ioneasDetail.FItemList(i).fitemoption%>
            </td>
            <td>
            	<%=ioneasDetail.FItemList(i).fmakerid%>
            </td>
            <td align="left">
            	<%= ioneasDetail.FItemList(i).Fitemname %>
            	<% if ioneasDetail.FItemList(i).Fitemoptionname<>"" then %>
            	<br>
            	<font color="blue">[<%= ioneasDetail.FItemList(i).Fitemoptionname %>]</font>
            	<% end if %>
            </td>
            <td align="right"><%= FormatNumber(ioneasDetail.FItemList(i).Fsellprice,0) %></td>
            <td align="center"><%= ioneasDetail.FItemList(i).Fitemno %></td>
        </tr>
        <% next %>
        </table>
    </td>
</tr>
<% end if %>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>배송CS 처리결과작성</b>
		&nbsp;&nbsp;
		*처리 내용 입력시 <font color=red>송장번호</font>등 상세내역을 기재해 주세요
	</td>
</tr>
<% if ioneas.FOneItem.Fdivcd = "A030" or ioneas.FOneItem.Fdivcd = "A031" then %>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><%= deliverydivname %></td>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
		    <td width="50" bgcolor="<%= adminColor("pink") %>">받는분</td>
		    <td width="80" bgcolor="#FFFFFF"><%= ioneas.FOneItem.Freqname %></td>
		    <td width="50" bgcolor="<%= adminColor("pink") %>">연락처</td>
		    <td bgcolor="#FFFFFF"><%= ioneas.FOneItem.Freqphone %> / <%= ioneas.FOneItem.Freqhp %></td>
		</tr>
		<tr>
		    <td bgcolor="<%= adminColor("pink") %>">주소</td>
		    <td colspan="3" bgcolor="#FFFFFF">
				[<%= ioneas.FOneItem.Freqzipcode %>] <%= ioneas.FOneItem.Freqzipaddr %> &nbsp;<%= ioneas.FOneItem.FReqAddress %>

				<% if (ioneas.FOneItem.Fcurrstate="B001") then %>
					<%
					'/업체a/s 이면서 업체일경우 ..  업체a/s(매장회수) 이면서 매장일 경우에만 .. 주소 수정 가능
					if (ioneas.FOneItem.Fdivcd="A030" and C_IS_Maker_Upche) or (ioneas.FOneItem.Fdivcd="A031" and C_IS_SHOP) then
					%>
					    <input class="button" type="button" value="주소변경" onclick="popEditCsDelivery('<%= ioneas.FOneItem.Fasid %>');" >
					<%
					'/관리자나 , 오프라인관리자 일경우 수정 가능
					elseif C_ADMIN_AUTH or C_OFF_AUTH then
					%>
						 <input class="button" type="button" value="주소변경(관리자모드)" onclick="popEditCsDelivery('<%= ioneas.FOneItem.Fasid %>');" >
					<% end if %>
				<% else %>
					<input class="button" type="button" value="주소변경불가" onclick="alert('접수상태에서만 변경가능 합니다.');" >
				<% end if %>
		    </td>
		</tr>
		</table>
	</td>
	<td width="45%" rowspan="3" valign="top">
		<%
		if ioneas.FOneItem.Fdivcd="A030" then
		%>
			*처리내용으로 입력된 정보는 매장과 업체가 공유하는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">매장이나 고객분께 출고후, 택배정보를 꼭 입력 부탁드립니다</font>
			<br>
			<br><font color="blue">*처리내용 입력요청사항</font>
			<br>출고일 :
			<br>기타내용 :
			<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
		<%
		elseif ioneas.FOneItem.Fdivcd="A031" then
		%>
			*처리내용으로 입력된 정보는 매장과 업체가 공유하는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">업체로 출고후, 택배정보를 꼭 입력 부탁드립니다</font>
			<br>
			<br><font color="blue">*처리내용 입력요청사항</font>
			<br>출고일 :
			<br>기타내용 :
			<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
		<% end if %>

	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">처리내용</td>
	<td>
		<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">관련운송장</td>
	<td>
		<% drawSelectBoxDeliverCompany "songjangdiv",ioneas.FOneItem.FSongjangdiv %>
		<input type="text" class="text" name="songjangno" value="<%= ioneas.FOneItem.Fsongjangno %>" size="14" maxlength="14">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<%
		'/업체처리완료 / 최종처리완료 / 매장처리완료 상태가 아닐경우
		if ioneas.FOneItem.Fcurrstate <> "B006" and ioneas.FOneItem.Fcurrstate <> "B007" and ioneas.FOneItem.Fcurrstate <> "B008" then

			'/업체a/s 이면서 업체일경우 ..  업체a/s(매장회수) 이면서 매장일 경우에만 .. 주소 수정 가능
			if (ioneas.FOneItem.Fdivcd="A030" and C_IS_Maker_Upche) or (ioneas.FOneItem.Fdivcd="A031" and C_IS_SHOP) then
		%>
				<input type="button" class="button" value="업체처리완료" onclick="javascript:SaveFin(frm);">
		<%
			'/관리자나 , 오프라인관리자 일경우 수정 가능
			elseif C_ADMIN_AUTH or C_OFF_AUTH then
		%>
			 	<input type="button" class="button" value="업체처리완료(관리자모드)" onclick="javascript:SaveFin(frm);">
			<% else %>
				<input class="button" type="button" value="완료처리불가" onclick="alert('매장 에서만 변경가능 합니다.');" >
			<% end if %>
		<% end if %>

		<input type="button" class="button" value="목록보기" onClick="location.href='/common/offshop/shopcscenter/upche_cslist.asp';">
	</td>
</tr>
</form>
<form name="frmshow" method="post">
	<input type="hidden" name="masteridx" value="">
</form>
</table>
<!-- 표 하단바 끝-->

<%
set ioneas = Nothing
set ioneasDetail = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->