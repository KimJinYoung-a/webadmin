<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lectureday_auctioncls.asp"-->
<%
dim mode,page
mode = request("mode")
page = request("page")

if page="" then
	page =1 
end if

dim Oauction
set Oauction = New CBoardAuction
Oauction.FCurrPage = page
Oauction.FPageSize = 10
Oauction.GetAllAuction 

dim i
%>
<table border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="right" height="30"><a href="lectureday_auction_write.asp?mode=add"><font color="red">NEW</font></a>&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td width="50" align="center">
			ID
		</td>
		<td width="120" align="center">
			경매이름
		</td>
		<td width="100" align="center">
			시작일
		</td>
		<td width="120" align="center">
			종료일
		</td>
		<td width="80" align="center">
			판매갯수
		</td>
		<td width="80" align="center">
			낙찰유무
		</td>
		<td width="80" align="center">
			낙찰가격
		</td>
		<td width="80" align="center">
			사용유무
		</td>
		<td width="100" align="center">
			등록일
		</td>
	</tr>
	<% for i=0 to Oauction.FResultcount -1 %>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">
			<%= Oauction.FAuctionList(i).Fidx %>
		</td>
		<td>
			&nbsp;<a href="lectureday_auction_write.asp?mode=edit&idx=<%= Oauction.FAuctionList(i).Fidx %>"><%= Oauction.FAuctionList(i).Fauctionname %></a>
		</td>
		<td align="center">
			<%= Oauction.FAuctionList(i).Fstartdate %>
		</td>
		<td align="center">
			<%= Oauction.FAuctionList(i).Ffinishdate %>
		</td>
		<td align="center">
			<%= Oauction.FAuctionList(i).Fitemea %>
		</td>
		<td align="center">
			<% if Oauction.FAuctionList(i).Fnakchaluser <> "" then %>
			<font color="red">Y</font>
			<% else %>
			N
			<% end if %>
		</td>
		<td align="center">
			<%= Oauction.FAuctionList(i).Fnakchalprice %>
		</td>
		<td align="center">
			<% if Oauction.FAuctionList(i).Fisusing="Y" then %>
			Y
			<% else %>
			<font color="red">N</font>
			<% end if %>
		</td>
		<td align="center">
			<%= Left(Oauction.FAuctionList(i).Fregdate,10) %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center" height="30">
			<% if Oauction.HasPreScroll then %>
				<a href="?page=<%= Oauction.StarScrollPage-1 %>">[pre]</a>
			<% else %> 
				[pre]
			<% end if %>
			
			<% for i=0 + Oauction.StarScrollPage to Oauction.FScrollCount + Oauction.StarScrollPage - 1 %>
				<% if i>Oauction.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>">[<%= i %>]</a>
				<% end if %>
			<% next %>
			
			<% if Oauction.HasNextScroll then %>
				<a href="?page=<%= i %>">[next]</a>
			<% else %> 
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set Oauction = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->