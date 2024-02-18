<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 cs내역
' Hieditor : 2011.03.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->

<%
dim csmasteridx ,ioneas,i ,ioneasDetail
	csmasteridx = request("csmasteridx")

set ioneas = new corder
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
%>

<script language='javascript'>

function ViewOrderDetail(frm,orgmasteridx){
	var props = "width=600, height=600, location=no, status=yes, resizable=no,";
	window.open("about:blank", "ViewOrderDetail", props);

    frm.target = 'ViewOrderDetail';
    frm.masteridx.value = orgmasteridx;
    frm.action="/common/offshop/beasong/upche_viewordermaster.asp";
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

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="shop_csprocess.asp">
<input type="hidden" name="orgmasteridx" value="<%= ioneas.FOneItem.forgmasteridx %>">
<input type="hidden" name="finishuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="masteridx" value="<%= ioneas.FOneItem.fmasteridx %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<b>배송CS 처리답변</b>
		&nbsp;&nbsp;
    	작성일 : <b><%= CStr(ioneas.FOneItem.Fregdate) %></b>
    	&nbsp;&nbsp;
    	<% if not IsNULL(ioneas.FOneItem.Ffinishdate) then %>
    	완료일 : <b><%= CStr(ioneas.FOneItem.Ffinishdate) %></b>
    	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td>
		<%= ioneas.FOneItem.Forderno %>
		<input type="button" class="button" value="주문상세보기" onclick="ViewOrderDetail(frmshow,'<%= ioneas.FOneItem.forgmasteridx %>');">
	</td>
	<td width="45%" rowspan="7" valign="top">
		<% if ioneas.FOneItem.Fdivcd="A000" then %> <!-- 맞교환 설명 -->
			<b>* 맞교환 도움말</b>
		<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- 누락재발송 설명 -->
			<b>* 누락재발송 도움말</b>
		<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- 반품 설명 -->
			<b>* 반품관련 도움말</b>
			<br>반품접수가 될경우, 고객님께 발송하신 택배사 전화번호를 안내해드리며,
			<br>상품을 받으신 택배사를 통해 <font color="blue">착불반송</font>을 해주시도록 안내를 해드리고 있습니다.
			<br><font color="blue">변심 반품의 경우, 착불반송포함 왕복배송비를 차감한 금액을 고객님께 환불해드리며,
			<br>차감된 금액응 업체정산내역에 자동으로 등록됩니다.</font>
			<br><font color="red">(편도 2,000원 / 왕복 4,000원 차감)</font>
			<br>
			<br>반송상품이 도착하면, 접수내용과 확인하신 후,
			<br>아래쪽 처리내용에 내용을 남겨주시면, 고객센터에 내용이 전달되며,
			<br>고객센터에서 반품취소처리 및 고객환불을 진행합니다.
			<br>
			<br>*처리프로세스
			<br>1.접수
			<br>2.업체완료처리 --> 고객센터에 처리결과 전달
			<br>3.고객센터완료처리 --> 고객에게 처리결과 안내 및 메일발송
		<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- 출고시 유의사항 설명 -->
			<b>* 출고시 유의사항 도움말</b>
			<br>주문건 확인 후, 고객님이 주문관련 변경을 요청하셨을 경우,
			<br>출고시 유의사항으로 등록됩니다.
			<br>ex)배송지변경/상품변경/상품옵션변경
			<br>
			<br><font color="red">텐바이텐 고객센터에서 별도로 가능여부 확인을 위해 연락드립니다.</font>
		<% else %>

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
            <td width="50">상품코드</td>
            <td>상품명<font color="blue">[옵션명]</font></td>
            <td width="50">판매가</td>
            <td width="40">수량</td>
        </tr>
        <% for i=0 to ioneasDetail.FResultCount-1 %>
        <tr bgcolor="#FFFFFF" align="center">
            <td>
            	<%=ioneasDetail.FItemList(i).fitemgubun%>-<%=CHKIIF(ioneasDetail.FItemList(i).fitemid>=1000000,Format00(8,ioneasDetail.FItemList(i).fitemid),Format00(6,ioneasDetail.FItemList(i).fitemid))%>-<%=ioneasDetail.FItemList(i).fitemoption%>
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
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">처리내용</td>
	<td>
		<textarea class="textarea" name="finishmemo" cols="60" rows="8" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
	</td>
	<td width="45%" rowspan="2" valign="top">
		<% if ioneas.FOneItem.Fdivcd="A000" then %> <!-- 맞교환 설명 -->
			*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</font>
			<br>
			<br><font color="blue">*처리내용 입력요청사항</font>
			<br>출고일 :
			<br>기타내용 :
			<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
		<% elseif ioneas.FOneItem.Fdivcd="A001" then %> <!-- 누락재발송 설명 -->
			*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">맞교환상품 출고후, 택배정보를 꼭 입력 부탁드립니다.</font>
			<br>
			<br><font color="blue">*처리내용 입력요청사항</font>
			<br>출고일 :
			<br>기타내용 :
			<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
		<% elseif ioneas.FOneItem.Fdivcd="A004" then %> <!-- 반품 설명 -->
			*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">반품상품 입고 완료 후, 처리내용 입력과 함께 완료처리 부탁드립니다.</font>
			<br>
			<br><font color="blue">*처리내용 입력요청사항</font>
			<br>반품방법 : 고객선불 / 착불
			<br>반품사유 : 불량반품 / 고객반품
			<br>환불계좌 : 은행명 + 계좌번호 + 예금주명(고객님이 첨부한 경우)
			<br>기타내용 :
			<br><font color="blue">*위 내용을 카피하셔서, 처리내용에 남겨주시면 감사하겠습니다.</font>
		<% elseif ioneas.FOneItem.Fdivcd="A006" then %> <!-- 출고시 유의사항 설명 -->
			*처리내용으로 입력된 정보는 고객센터에 전달되는 정보입니다.
			<br>(고객님께 오픈되는 정보가 아닙니다.)
			<br>
			<br><font color="red">고객센터에서 요청한 출고유의사항에 대한 처리유무를 알려주시기 바랍니다.</font>
			<br>발송 후, 이 내용을 확인하셨을 경우에도, 미반영 출고로 완료처리 부탁드립니다.
		<% else %>

		<% end if %>
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
	<% if ioneas.FOneItem.Fcurrstate="B007" then %>

	<% else %>
		<input type="button" class="button" value="완료처리" onclick="javascript:SaveFin(frm);">
    <% end if %>
		<input type="button" class="button" value="목록보기" onClick="location.href='/common/offshop/beasong/shop_cslist.asp';">
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