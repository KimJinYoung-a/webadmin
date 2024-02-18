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
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/100proshopCls.asp" -->
<%
dim i
dim page

page = request("page")
if (page = "") then
        page = "1"
end if

'==============================================================================
dim o100pro
set o100pro = new C100ProShop

o100pro.FCurrPage = CInt(page)
o100pro.FPageSize = 20
o100pro.getMasterList

dim premasteridx
%>
<a href="/admin/eventmanage/event/event_regist.asp?eK=3"><b>[신규등록]</b></a>
<table width="800" border="0" cellpadding="3" cellspacing="1" bgcolor=#3d3d3d class=a>
	<tr bgcolor="#DDDDFF">
		<td width="30" align="center">ID</td>
		<td align="center" width="50" height="50">마스타 이미지</td>
		<td align="center" width="50" height="50">아이템 이미지</td>
		<td align="center" width="50">아이템ID</td>		
		<td align="center">쿠폰발급기준일</td>
		<td align="center">쿠폰유효기간</td>
		<td align="center">쿠폰금액</td>
		<td align="center">최소구매금액</td>
		<td align="center">상품수정</td>
	</tr>
<% for i=0 to o100pro.FResultcount -1 %>
	<% if premasteridx<>o100pro.FItemList(i).FIdx then %>
	<tr bgcolor="#FFFFFF">
		<td height="50" align="center">
			<%= o100pro.FItemList(i).FIdx %>
		</td>
		<td align="center" width="50"><a href="/admin/eventmanage/event/event_modify.asp?eC=<%= o100pro.FItemList(i).FIdx %>&menupos=<%=menupos%>"><img src="<%= o100pro.FItemList(i).Flistimg %>" width="50" border=0 alt=""></a></td>
		<td colspan="7" align=right>
			<input type="button" value="쿠폰발급현황" class="button" onClick="self.location='100proshop_couponregview.asp?eC=<%= o100pro.FItemList(i).FIdx %>&menupos=<%=menupos%>'">
			<input type="button" value="상품추가" class="button" onClick="self.location='100proshop_itemwrite.asp?eC=<%= o100pro.FItemList(i).Fidx %>&mode=write&menupos=<%=menupos%>'">
		</td>
	</tr>
	<% end if %>
	<%
		premasteridx = o100pro.FItemList(i).FIdx

		'// 관련 상품 목록 출력
		if Not(o100pro.FItemList(i).FItemId="" or isNull(o100pro.FItemList(i).FItemId)) then
	%>
	<tr bgcolor="#FAFAFA">
		<td colspan="2" align="center" bgcolor="#F0F0F0">&nbsp;</td>
		<td align="center" width="50"><img src="<%= o100pro.FItemList(i).FItemImageSmall %>" width=50></td>
		<td align="center">
			<a href="100proshop_itemwrite.asp?eC=<%= o100pro.FItemList(i).Fidx %>&idx=<%= o100pro.FItemList(i).Fdetailidx %>&mode=modify&menupos=<%=menupos%>"><%= o100pro.FItemList(i).FItemId %></a>
		</td>		
		<td align="center">
			<% = o100pro.FItemList(i).FStartDate %>
			<br>~<br><%= o100pro.FItemList(i).FEndDate %>
		</td>
		<td align="center">
			<% = o100pro.FItemList(i).FCouponStartDate %>
			<br>~<br><%= o100pro.FItemList(i).FCouponExpireDate %>
		</td>
		<td align="right"><% = FormatNumber(o100pro.FItemList(i).FCouponValue,0) %></td>
		<td align="right"><% = FormatNumber(o100pro.FItemList(i).Fminbuyprice,0) %></td>
		<td align="center">
			<input type="button" value="상품수정" class="button" onClick="self.location='100proshop_itemwrite.asp?eC=<%= o100pro.FItemList(i).Fidx %>&idx=<%= o100pro.FItemList(i).Fdetailidx %>&mode=modify&menupos=<%=menupos%>'">
		</td>
	</tr>
	<%	else %>
	<tr>
		<td colspan="2" align="center" bgcolor="#F0F0F0">&nbsp;</td>
		<td colspan="7" align="center" bgcolor="#FAFAFA">관련 상품이 없습니다. [상품추가]를 눌러 추가해주십시요.</td>
	</tr>
<%
		end if
	next
%>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center">
			<% if o100pro.HasPreScroll then %>
				<a href="?page=<%= o100pro.StarScrollPage-1 %>&menupos=<%=menupos%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + o100pro.StarScrollPage to o100pro.FScrollCount + o100pro.StarScrollPage - 1 %>
				<% if i>o100pro.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%=menupos%>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if o100pro.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%=menupos%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set o100pro = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
