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
<!-- #include virtual="/lib/classes/sitemasterclass/vipcornerCls.asp"-->
<%
'###############################################
' PageName : vipcorner.asp
' Discription : 우수회원 전용코너 관리
' History : 2015.04.15 원승현 생성
'###############################################

dim page, div, i, lp, vUsing

page = request("page")
vUsing = request("using")
if page = "" then page=1

If vUsing = "" Then vUsing="Y"

dim oVip
set oVip = New CVip
oVip.FCurrPage = page
oVip.FPageSize=20
oVip.FRectUsing = vUsing
oVip.GetVipCornerList






%>
<script type="text/javascript">
<!--
// 페이지 이동
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="vipcorner.asp";
	document.refreshFrm.submit();
}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="vipcorner.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		사용여부
		<select name="using">
			<option value="all" >전체</option>
			<option value="Y" <% If vUsing="Y" Then %> selected <% End If %>>Y</option>
			<option value="N" <% If vUsing="N" Then %> selected <% End If %>>N</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="left">※ 사용여부가 Y 이여도 해당 이벤트의 이벤트 기간이 끝나면 자동으로 페이지에서 내려갑니다.</td>
	<td align="right"><input type="button" value="이벤트 추가" onclick="window.open('vip_Write.asp?mode=add&menupos=<%= menupos %>', '','width=800, height=300, scrollbars=yes');" class="button"></td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oVip.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oVip.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>이벤트코드</td>
	<td>이벤트명</td>
	<td>이벤트기간</td>
	<td>pc용이미지</td>
	<td>모바일/앱용 이미지</td>
	<td>순서번호</td>
	<td>사용여부</td>
	<td>등록일</td>
</tr>
<%	if oVip.FResultCount < 1 then %>
<tr>
	<td colspan="9" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 이벤트가 없습니다.</td>
</tr>
<%
	else
		for i=0 to oVip.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= oVip.FItemList(i).Fidx %></td>
	<td align="center"><a href="" onclick="window.open('vip_Write.asp?mode=edit&idx=<%=oVip.FItemList(i).Fidx%>&menupos=<%= menupos %>', '','width=800, height=300, scrollbars=yes');return false;" ><%= oVip.FItemList(i).FevtCode %></a></td>
	<td align="center"><a href="" onclick="window.open('vip_Write.asp?mode=edit&idx=<%=oVip.FItemList(i).Fidx%>&menupos=<%= menupos %>', '','width=800, height=300, scrollbars=yes');return false;" ><%= oVip.FItemList(i).FevtName %></a></td>
	<td align="center"><%= oVip.FItemList(i).FevtStartDate %>~<%= oVip.FItemList(i).FevtEndDate %></td>
	<td align="center"><a href="" onclick="window.open('vip_Write.asp?mode=edit&idx=<%=oVip.FItemList(i).Fidx%>&menupos=<%= menupos %>', '','width=800, height=300, scrollbars=yes');return false;" ><img src="<%= webImgUrl&"/vipcorner/"&oVip.FItemList(i).Fpcimg %>" border="0"></a></td>
	<td align="center"><img src="<%= webImgUrl&"/vipcorner/"&oVip.FItemList(i).Fmaing %>" border="0"></td>
	<td align="center"><%= oVip.FItemList(i).Forderby %></td>
	<td align="center"><%= oVip.FItemList(i).Fisusing %></td>
	<td align="center"><%= oVip.FItemList(i).Fregdate %></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<!-- 페이지 시작 -->
	<%
		if oVip.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oVip.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oVip.StartScrollPage to oVip.FScrollCount + oVip.StartScrollPage - 1

			if lp>oVip.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oVip.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<%
set oVip = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->