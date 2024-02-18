<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, research, i, page, itemid, makerid, nowsDate,iSD
research	= requestCheckvar(request("research"),2)
dim yyyy1, mm1, beasongpayChk
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)

if (yyyy1="") then yyyy1=LEFT(dateadd("d",-3,NOW()),4)
if (mm1="") then mm1=MID(dateadd("d",-3,NOW()),6,2)

SET oOrder = new COrder

	oOrder.FRectYYYYMM	= yyyy1&"-"&mm1
	oOrder.getKakaoDlvJCheckList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function popDlvPriceEdit(iorderserial){
	var popwin;
	popwin = window.open("","_popDlvPriceEdit");
	popwin.location.href="popDlvPriceEdit.asp?orderserial="+iorderserial;
	popwin.focus();
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">

<table width="100%" align="center" cellspacing="1" cellpadding="3" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="40">
	<td align="left">
		정산기준월일 :
		<% DrawYMBox yyyy1,mm1 %>
	</td>
	<td align="right" width="100">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		검색결과 : <b><%= FormatNumber(oOrder.FResultCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>주문번호</td>
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>소비자가(판매시)</td>
	<td>판매가(판매시)</td>
	<td>현재소비자가</td>
	<td>현재판매가</td>
	<td>매입구분</td>
	<td>출고일</td>
    <td>정산일</td>
    <td>비고</td>
</tr>

<% If oOrder.FResultCount > 0 Then %>
<% For i=0 to oOrder.FResultCount - 1 %>
<%
If (session("ssBctID")="kjy8517") Then
	If (oOrder.FItemList(i).Fitemcost - oOrder.FItemList(i).Forgitemcost = "3000") OR (oOrder.FItemList(i).Fitemcost - oOrder.FItemList(i).Forgitemcost = "2500") Then
		beasongpayChk = "Y"
	Else
		beasongpayChk = ""
	End If
End If
%>
<tr align="center" <% If beasongpayChk <> "" Then response.write "bgcolor= 'PINK'" Else response.write "bgcolor= '#FFFFFF'" End If %>>
    <td><a href="#" onClick="popDlvPriceEdit('<%=oOrder.FItemList(i).FOrderserial%>');return false;"><%=oOrder.FItemList(i).FOrderserial%></a></td>
	<td><%= oOrder.FItemList(i).FItemid %></td>
	<td><%= oOrder.FItemList(i).FMakerid %></td>
	<td><%= FormatNumber(oOrder.FItemList(i).Forgitemcost,0) %></td>
	<td><%= FormatNumber(oOrder.FItemList(i).Fitemcost,0) %></td>
	<td><%= FormatNumber(oOrder.FItemList(i).Fitemorgprice,0) %></td>
	<td><%= FormatNumber(oOrder.FItemList(i).Fitemsellcash,0) %></td>
	<td><%= oOrder.FItemList(i).Fomwdiv %>/<%= oOrder.FItemList(i).Fmwdiv %></td>
	<td><%= oOrder.FItemList(i).Fbeasongdate %></td>
    <td><%= oOrder.FItemList(i).FJungsanFixDate %></td>
    <td>

    </td>
</tr>
<%  if (i mod 1000)=0 then response.flush  %>
<% Next %>
<% Else %>
<tr height="50">
    <td colspan="11" align="center" bgcolor="#FFFFFF">
		데이터가 없습니다
    </td>
</tr>
<% End If %>
</table>
<% SET oOrder = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->