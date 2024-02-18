<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<%
Dim midx, arrList, olow, i, myrank, bcolor
midx	= request("midx")
myrank	= request("myrank")
SET olow = new epShop
	olow.FRectMidx			= midx
    olow.getNaverLowpriceDetailList
	if (olow.FRectMidx="") then
    arrList = olow.getNaverLowpriceDetailList
	end if
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">상품코드</td>
	<td width="60">Rank</td>
	<td width="100">판매가</td>
</tr>
<% if olow.FRectMidx="" then %>
	<tr align="center" bgcolor="<%=bcolor%>">
		<td colspan="3">No idx</td>
	</tr>
<% else %>
	<% if isArray(arrList) then %>
	<% For i = 0 To UBound(arrList,2) %>
	<%
		If CInt(myrank) = CInt(arrList(1,i)) Then
			bcolor = "GOLD"
		Else
			bcolor = "WHITE"
		End If
	%>
	<tr align="center" bgcolor="<%=bcolor%>">
		<td><%= arrList(0,i) %></td>
		<td><%= arrList(1,i) %></td>
		<td><%= FormatNumber(arrList(2,i),0) %></td>
	</tr>
	<% Next %>
	<% end if %>
<% end if %>
</table>
<% SET olow = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->