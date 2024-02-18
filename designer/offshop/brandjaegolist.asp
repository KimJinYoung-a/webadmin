<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim ooffsell,makerid
dim shopid

makerid = session("ssBctID")
shopid = request("shopid")

set ooffsell = new COffShopDailyStock
ooffsell.FRectMakerid = makerid
ooffsell.FRectShopId = shopid
ooffsell.GetRealJaegoList

dim i
%>
<script language='javascript'>
function inputjaego(){
    alert('재고 시스템 변경작업으로 한시적으로 재고 입력을 중지합니다.');
    return;
    
	if (frm.shopid.value.length<1){
		alert('샾을 선택하세요.');
		frm.shopid.focus();
		return;
	}

	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&shopid=' + frm.shopid.value;
}

function jaegoedit(idx){
    alert('재고 시스템 변경작업으로 한시적으로 재고 입력을 중지합니다.');
    return;
	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&idx=' + idx;
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			SHOP : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="실사재고 입력" onClick="inputjaego();">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">idx</td>
		<td width="100">오프샾ID</td>
		<td width="150">브랜드ID</td>
		<td width="150">실사재고파악일시</td>
		<td width="150">등록일</td>
		<td>수정</td>
	</tr>
	<% for i=0 to ooffsell.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ooffsell.FItemList(i).Fidx %></td>
		<td><%= ooffsell.FItemList(i).Fshopid %></td>
		<td><%= ooffsell.FItemList(i).Fmakerid %></td>
		<td><%= ooffsell.FItemList(i).Fjeagodate %></td>
		<td><%= ooffsell.FItemList(i).Fregdate %></td>
		<td><input type="button" class="button" value="수정" onClick="jaegoedit('<%= ooffsell.FItemList(i).Fidx %>');"></td>
	</tr>
	<% next %>
</table>
<%
set ooffsell = Nothing
%>
<script language='javascript'>
alert('재고 시스템 변경작업으로 한시적으로 재고 입력을 중지합니다.');
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->