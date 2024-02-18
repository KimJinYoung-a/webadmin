<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim userid, sitename
dim page

userid = request("userid")
sitename = request("sitename")
page = request("page")
if (page="") then page=1


dim ojumun
set ojumun = new CJumunMaster
ojumun.FRectUserID = userid
ojumun.FRectSiteName = sitename
ojumun.FRectIpkumDiv4 = "on"
ojumun.FCurrPage = page
ojumun.SearchJumunList

dim ix
%>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
	<td width="30" align="center">선택</td>
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">Site</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="65" align="center">수령인</td>
	<td width="60" align="center">할인율</td>
	<td width="72" align="center">결제금액</td>
	<td width="72" align="center">구매총액</td>
	<td width="74" align="center">결제방법</td>
	<td width="74" align="center">거래상태</td>
	<td width="40" align="center">삭제여부</td>
	<td width="120" align="center">주문일</td>
</tr>
<% for ix=0 to ojumun.FresultCount-1 %>
<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" >
<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).UserIDName %>">
<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
<tr class="a"  bgcolor="#FFFFFF">

<% else %>
<tr class="gray"  bgcolor="#FFFFFF">
<% end if %>
	<td align="center"><input type="checkbox" name="ckbox"></td>
	<td align="center"><a href="#" onclick="ViewOrderDetail(frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
	<td align="center"><font color="<%= ojumun.FMasterItemList(ix).SiteNameColor %>"><%= ojumun.FMasterItemList(ix).FSitename %></font></td>
	<% if ojumun.FMasterItemList(ix).UserIDName<>"&nbsp;" then %>
	<td align="center"><a href="#" onclick="ViewUserInfo(frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FMasterItemList(ix).UserIDName %></a></td>
	<% else %>
	<td align="center"><%= ojumun.FMasterItemList(ix).UserIDName %></td>
	<% end if %>
	<td align="center"><%= ojumun.FMasterItemList(ix).FBuyName %></td>
	<td align="center"><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center"><%= ojumun.FMasterItemList(ix).FDisCountrate %></td>
	<td align="right"><font color="<%= ojumun.FMasterItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FMasterItemList(ix).FSubTotalPrice,0) %></font></td>
	<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
	<td align="center"><%= ojumun.FMasterItemList(ix).JumunMethodName %></td>
	<td align="center"><font color="<%= ojumun.FMasterItemList(ix).IpkumDivColor %>"><%= ojumun.FMasterItemList(ix).IpkumDivName %></font></td>
	<td align="center"><font color="<%= ojumun.FMasterItemList(ix).CancelYnColor %>"><%= ojumun.FMasterItemList(ix).CancelYnName %></font></td>
	<td align="center"><%= Left(ojumun.FMasterItemList(ix).GetRegDate,16) %></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="13" height="30" align="center">
	<% if ojumun.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
	<% else %> 
		[pre]
	<% end if %>
	
	<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
		<% if ix>ojumun.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(ix) then %>
		<font color="red">[<%= ix %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
		<% end if %>
	<% next %>
	
	<% if ojumun.HasNextScroll then %>
		<a href="javascript:NextPage('<%= ix %>')">[next]</a>
	<% else %> 
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->