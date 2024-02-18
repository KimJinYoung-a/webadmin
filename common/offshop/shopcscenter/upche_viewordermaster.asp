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
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/shopcscenterupchebeasong_cls.asp" -->

<%
dim masteridx , ix , sellsum , menupos
	masteridx = requestCheckVar(request("masteridx"),10)
	menupos = requestCheckVar(request("menupos"),10)

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
	response.write "	alert('해당 내역이 없습니다');"
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
    	<b>구매자명</b> : <%= ojumun.FItemList(0).FBuyName %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
	<td width="225" bgcolor="#FFFFFF"><%= ojumun.FItemList(0).Forderno %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">매장</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).fshopname %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">주문입력일</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FRegDate %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">취소여부</td>
	<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).fcancelyn %></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">온라인ID</td>
	<td bgcolor="#FFFFFF" colspan=3><%= ojumun.FItemList(0).fonlineuserid %></td>
</tr>
<%
'/마일리지 카드 정보
if ojumun.FItemList(0).fpointuserno <> "" then
%>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">포인트카드</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).fpointuserno %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FBuyName %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자전화</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FBuyPhone %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">구매자핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ojumun.FItemList(0).FBuyHp %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이메일</td>
		<td bgcolor="#FFFFFF" colspan=3><%= ojumun.FItemList(0).Fbuyemail %></td>
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
	<td>취소여부</td>
</tr>
<% if ojumun.FResultCount > 0 then %>
<% for ix=0 to ojumun.FResultCount - 1 %>

<% sellsum = sellsum + ojumun.FItemList(ix).fsellprice*ojumun.FItemList(ix).FItemNo %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ojumun.fitemlist(ix).fitemgubun %>-<%= CHKIIF(ojumun.fitemlist(ix).FitemID>=1000000,Format00(8,ojumun.fitemlist(ix).FitemID),Format00(6,ojumun.fitemlist(ix).FitemID)) %>-<%= ojumun.fitemlist(ix).fitemoption %></td>
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
		<%= ojumun.FItemList(ix).fdetailcancelyn %>
	</td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
	<td>합계</td>
	<td colspan="4" align="right"><%= FormatNumber(sellsum,0) %></td>
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
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->