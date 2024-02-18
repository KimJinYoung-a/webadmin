<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/100proshopCls.asp" -->
<%
dim eCode
eCode = request("eC")

dim o100pro
set o100pro = new C100ProShop
o100pro.FRectIdx = eCode
o100pro.getCouponOpenList

dim i

dim sum1, sum2, sum3, sum4
%>

<table width="800" border="0" cellpadding="3" cellspacing="1" bgcolor=#3d3d3d class=a>
<tr bgcolor="#DDDDFF">
	<td width=50>상품ID</td>
	<td width=50>이미지</td>
	<td width=60>발급기간중 판매수</td>
	<td width=60>쿠폰발급가능총수 (배송완료)</td>
	<td width=60>쿠폰발급수</td>
	<td width=30>발급율</td>
	<td width=60>쿠폰사용수</td>
	<td width=30>사용율</td>
	<td></td>
</tr>
<% for i=0 to o100pro.FResultCount -1 %>
<%
sum1 = sum1 + o100pro.FITemList(i).FTotalSellCount
sum2 = sum2 + o100pro.FITemList(i).FMatchCount
sum3 = sum3 + o100pro.FITemList(i).FRegCount
sum4 = sum4 + o100pro.FITemList(i).FUseCount
%>
<tr bgcolor="#FFFFFF">
	<td><%= o100pro.FITemList(i).FItemID %></td>
	<td><img src="<%= o100pro.FITemList(i).FImgSmall %>" width=50 height=50 ></td>
	<td align=right><%= FormatNumber(o100pro.FITemList(i).FTotalSellCount,0) %></td>
	<td align=right><%= FormatNumber(o100pro.FITemList(i).FMatchCount,0) %></td>
	<td align=right><%= FormatNumber(o100pro.FITemList(i).FRegCount,0) %></td>
	<td align=center>
	<% if o100pro.FITemList(i).FMatchCount<>0 then %>
	<%= CLng(o100pro.FITemList(i).FRegCount/o100pro.FITemList(i).FMatchCount*100) %> %
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(o100pro.FITemList(i).FUseCount,0) %></td>
	<td align=center>
	<% if o100pro.FITemList(i).FRegCount<>0 then %>
	<%= CLng(o100pro.FITemList(i).FUseCount/o100pro.FITemList(i).FRegCount*100) %> %
	<% end if %>
	</td>
	<td></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>총계</td>
	<td></td>
	<td align=right><%= FormatNumber(sum1,0) %></td>
	<td align=right><%= FormatNumber(sum2,0) %></td>
	<td align=right><%= FormatNumber(sum3,0) %></td>
	<td align=center>
	<% if sum2<>0 then %>
	<%= CLng(sum3/sum2*100) %> %
	<% end if %>
	</td>
	<td align=right><%= FormatNumber(sum4,0) %></td>
	<td align=center>
	<% if sum3<>0 then %>
	<%= CLng(sum4/sum3*100) %> %
	<% end if %>
	</td>
	<td></td>
</tr>
</table>

<%
set o100pro = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
