<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim masteridx , ix , sellsum
	masteridx = requestCheckVar(request("masteridx"),10)

sellsum = 0

dim ojumun
set ojumun = new cupchebeasong_list
	ojumun.FRectmasteridx = masteridx

	if C_IS_Maker_Upche then
		ojumun.FRectDesignerID = session("ssBctID")
	end if
	'ojumun.FRectIpkumdiv = " and Currstate <= 3"

if masteridx<>"" then
    ojumun.fSearchJumunList()
end if

if (ojumun.FTotalCount < 1) then
	response.write "<script language='javascript'>"
	response.write "	alert('발주이전 주문이거나 / 주문 확인 안하신 상품이 있습니다. \n\n텐바이텐에서 발주 후 \n\n오프샾관리>>*업체배송주문확인 에서 주문 확인 하신 후 사용하실 수 있습니다');"
	response.write "	window.close();"
	response.write "</script>"
    dbget.close()	:	response.End
end if

%>

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

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
    	<b>주문번호</b> : <%= ojumun.FItemList(0).Forderno %>&nbsp;&nbsp;&nbsp;&nbsp;
    	<b>수령인명</b> : <%= ojumun.FItemList(0).FreqName %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td width="225" bgcolor="#FFFFFF"><%= ojumun.FItemList(0).Forderno %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">주문상태</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).shopIpkumDivName %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">주문입력일</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FRegDate %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">취소여부</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).fcancelyn %></td>
</tr>

<% if ojumun.FItemList(0).Fipkumdiv > 4 then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인</td>
	<td colspan=3 bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FReqName %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인전화</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FReqPhone %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인핸드폰</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FReqHp %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">수령인주소</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<%= ojumun.FItemList(0).FReqZipCode %>
		<br>
		<%= ojumun.FItemList(0).FReqZipAddr %>
		&nbsp;<%= ojumun.FItemList(0).FReqAddress %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">기타사항</td>
	<td colspan="3" bgcolor="#FFFFFF">
	<%= nl2br(ojumun.FItemList(0).FComment) %>
	</td>
</tr>
<% else %>
<tr align="center">
	<td colspan=4 bgcolor="#FFFFFF"><font color="blue"><b>배송정보는 [업체주문통보] 상태 이후에 확인가능합니다.</b></font></td>
</tr>
<% end if %>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
    	<b>주문상품정보</b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>상품명<font color="blue">[옵션명]</font></td>
	<td>수량</td>
	<td>판매가격</td>
	<td>배송<br>구분</td>
	<td>취소여부</td>
	<td>상태</td>
</tr>
<% if ojumun.FResultCount > 0 then %>
<% for ix=0 to ojumun.FResultCount - 1 %>

<% sellsum = sellsum + ojumun.FItemList(ix).Fsellprice*ojumun.FItemList(ix).FItemNo %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ojumun.fitemlist(ix).fitemgubun %>-<%= FormatCode(ojumun.fitemlist(ix).FitemID) %>-<%= ojumun.fitemlist(ix).fitemoption %></td>
	<td align="left">
		<%= ojumun.FItemList(ix).FItemName %>
		<br>
		<% if ojumun.FItemList(ix).FItemoptionName <> "" then %>
			<font color="blue">[<%= ojumun.FItemList(ix).FItemoptionName %>]</font>
		<% end if %>
	</td>
	<td><%= ojumun.FItemList(ix).FItemNo %></td>
	<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fsellprice,0) %></td>
	<td>
		<%= ojumun.FItemList(ix).getbeasonggubun %>
	</td>
	<td>
		<%= ojumun.FItemList(ix).fdetailcancelyn %>
	</td>
	<td>
		<%= ojumun.FItemList(ix).shopNormalUpcheDeliverState %>
		<% if ojumun.FItemList(ix).FMisendReason <> "" then %>
			<br><font color='red'><%= ojumun.FItemList(ix).getMisendText %></font>
		<% end if %>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td>합계</td>
	<td colspan="4" align="right"><%= FormatNumber(sellsum,0) %></td>
	<td colspan="2"></td>
</tr>
<% end if %>
</table>

<br>

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

<%
set ojumun = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->