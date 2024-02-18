<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원 구매 히스토리
' Hieditor : 2011.02.16 한용민 생성
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
<!-- #Include virtual = "/lib/classes/totalpoint/totalpointCls.asp" -->

<%
Dim ohistory, i, page , orderno ,totrealprice
dim vCardNo, vUserName, vUserID, posuid, pssnkey, dummikey, shopid
	orderno = requestCheckVar(Request("orderno"),20)
	vCardNo			= requestCheckVar(Request("cardno"),20)
	vUserName		= requestCheckVar(Request("username"),20)
	vUserID			= requestCheckVar(Request("userid"),32)
	posuid			= Request("posuid")
	pssnkey			= Request("pssnkey")
	dummikey		= Request("dummikey")
	shopid = request("shopid")
	menupos = request("menupos")

set ohistory = new TotalPoint
	ohistory.frectorderno = orderno
	ohistory.fsell_history_detail()
%>

<script language="javascript">

function refer(){
	frm.action='/admin/totalpoint/customer_sell_history.asp';
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="posuid" value="<%=posuid%>">
<input type="hidden" name="pssnkey" value="<%=pssnkey%>">
<input type="hidden" name="dummikey" value="<%=dummikey%>">
<input type="hidden" name="cardno" value="<%=vCardNo%>">
<input type="hidden" name="username" value="<%=vUserName%>">
<input type="hidden" name="userid" value="<%=vUserID%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		주문번호: <input type="text" class="text" name="orderno" value="<%=orderno%>">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% '<input type="button" class="button" value="목록으로" onClick="refer();"> %>
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ohistory.FTotalCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ohistory.FTotalCount %></b> ※총 1000건 까지 검색 됩니다
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>주문번호</td>
	<td>매장이름</td>
	<td>매장ID</td>
	<td>상품번호</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td>판매금액</td>
	<td>실결제액</td>
	<td>판매수량</td>
	<td>합계</td>
	<td>비고</td>
</tr>
<%

for i=0 to ohistory.FTotalCount-1

if ohistory.FItemList(i).fcancelyn = "N" then
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>
		<%= ohistory.FItemList(i).forderno %>
		
		<% if ohistory.FItemList(i).fcancelyn = "Y" then %>
			<br>(취소)
		<% end if %>
	</td>
	<td><%= ohistory.FItemList(i).fshopname %></td>
	<td><%= ohistory.FItemList(i).fshopid %></td>
	<td><%= ohistory.FItemList(i).fitemgubun %>-<%= CHKIIF(ohistory.FItemList(i).fitemid>=1000000,Format00(8,ohistory.FItemList(i).fitemid),Format00(6,ohistory.FItemList(i).fitemid)) %>-<%= ohistory.FItemList(i).fitemoption %></td>
	<td><%= ohistory.FItemList(i).fitemname %><br></td>
	<td><%= ohistory.FItemList(i).fitemoptionname %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).fsellprice,0) %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).frealsellprice,0) %></td>
	<td><%= ohistory.FItemList(i).fitemno %></td>
	<td><%= FormatNumber(ohistory.FItemList(i).frealsellprice*ohistory.FItemList(i).fitemno,0) %></td>
	<td></td>
</tr>
<%
totrealprice = totrealprice + (ohistory.FItemList(i).frealsellprice*ohistory.FItemList(i).fitemno)
next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=6>합계</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=10></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set ohistory = nothing
%>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->