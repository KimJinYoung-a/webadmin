<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_rankingCls.asp"-->

<%
	dim page

	page = requestCheckVar(request("page"),9)	
	if page = "" then page = 1 End If

	dim cMomoRankingList, PageSize , ttpgsz , i

	'### 내가 사용 코인 내역
	set cMomoRankingList = new ClsMomoRanking
	cMomoRankingList.FPageSize = 30
	cMomoRankingList.FCurrPage = page
	cMomoRankingList.FRankingList
%>

<script language="javascript">
function RankingDetail(idx)
{
	window.open('/admin/momo/ranking/ranking_detail.asp?idx='+idx+'','ranking_deatil'+idx+'','width=700,height=800,scrollbars=yes');
}
</script>

<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr height="30">
	<td><a href="javascript:RankingDetail('');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoRankingList.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr>
				<td>검색결과 : <b><%= cMomoRankingList.FTotalCount %></b></td>
				<td align="right">&nbsp;</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="100">idx</td>
	    <td align="center">주제</td>
	    <td align="center" width="200">기간</td>
	    <td align="center" width="100">사용여부</td>
	    <td align="center" width="100">등록일</td>
	    <td align="center" width="80"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoRankingList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">	
	    <td align="center"><%= cMomoRankingList.FItemList(i).fidx %></td>
	    <td><%= cMomoRankingList.FItemList(i).ftitle %></td>
	    <td align="center"><%= cMomoRankingList.FItemList(i).fstartdate %> ~ <%= cMomoRankingList.FItemList(i).fenddate %></td>
	    <td align="center"><% If cMomoRankingList.FItemList(i).fisusing = "Y" Then %>사용중<% Else %>삭제됨<% End If %></td>
	    <td align="center"><%= cMomoRankingList.FItemList(i).fregdate %></td>
	    <td align="center">
	    	<input type="button" value="상세보기" onClick="RankingDetail(<%=cMomoRankingList.FItemList(i).fidx%>);">
	    </td>
	</tr>
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cMomoRankingList.HasPreScroll then %>
				<span class="list_link"><a href="?menupos=<%=menupos%>&page=<%= cMomoRankingList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoRankingList.StartScrollPage to cMomoRankingList.StartScrollPage + cMomoRankingList.FScrollCount - 1 %>
				<% if (i > cMomoRankingList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoRankingList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?menupos=<%=menupos%>&page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoRankingList.HasNextScroll then %>
				<span class="list_link"><a href="?menupos=<%=menupos%>&page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set cMomoRankingList = nothing	
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
