<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_noticecls.asp" -->
<!-- #include virtual="/lib/classes/board/10x10_boardcls.asp" -->
<%

dim i, ix
dim page

page = request("page")
if (page = "") then
        page = "1"
end if

'==============================================================================
'공지사항

dim gubun

''프랜차이즈
if session("ssBctDiv") = 501 then
	gubun = "51"
elseif session("ssBctDiv") = 502 then
	gubun = "52"
elseif session("ssBctDiv") = 503 then
	gubun = "53"
end if

dim boardnotice
set boardnotice = New CNotice

boardnotice.FPageSize = 20
boardnotice.FCurrPage = CInt(page)
boardnotice.FRectGubun = gubun

''boardnotice.offshopnoticelist

%>

<script language="JavaScript">
<!--

function NextPage(ipage){
	document.noticeform.page.value= ipage;
	document.noticeform.submit();
}

//-->
</script>

<% IF (FALSE) then %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
<tr>
	<td class="a">

		<table width="780"  class="a">
		<tr>
			<td width=100></td>
			<td align="center">****** 공지사항 ******</td>
			<td width=100 align="right"></td>
		</tr>
		</table>
		<table width="780" cellspacing="1"  class="a" bgcolor=#3d3d3d>
		<form method=post name="noticeform">
		<input type="hidden" name="page" value="1">
		<tr bgcolor="#FFFFFF">
			<td colspan="4" height="25" align="right">검색결과 : 총 <font color="red"><% = boardnotice.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<tr bgcolor="#DDDDFF">
			<td width="50" align="center">번호</td>
			<td align="center">제목</td>
			<td width="100" align="center">글쓴이</td>
			<td width="100" align="center">날짜</td>
		</tr>
		<% for i = 0 to (boardnotice.FResultCount - 1) %>
		  <tr bgcolor="#FFFFFF">
			<td width="50" align="center"><%= boardnotice.FItemList(i).Fidx %></td>
			<td>&nbsp;<a href="notics_read.asp?idx=<%= boardnotice.FItemList(i).Fidx %>"><%= boardnotice.FItemList(i).Ftitle %></a>
			<% if datediff("d",boardnotice.FItemList(i).Fregdate,now())<6 then %>
			&nbsp;&nbsp;&nbsp;<img src="/images/new.gif">
			<% end if %>
			</td>
			<td width="100" align="center"><%= boardnotice.FItemList(i).Fusername %></td>
			<td width="100" align="center"><%= FormatDateTime(boardnotice.FItemList(i).Fregdate,2) %></td>
		  </tr>
		<% next %>
		</form>
		<tr bgcolor="#FFFFFF">
		        <td colspan=4 align=center>
		        <% if boardnotice.HasPreScroll then %>
			<a href="javascript:NextPage('<%= boardnotice.StarScrollPage-1 %>')">[pre]</a>
        		<% else %>
        			[pre]
        		<% end if %>
        		<% for ix=0 + boardnotice.StarScrollPage to boardnotice.FScrollCount + boardnotice.StarScrollPage - 1 %>
        			<% if ix>boardnotice.FTotalpage then Exit for %>
        			<% if CStr(page)=CStr(ix) then %>
        			<font color="red">[<%= ix %>]</font>
        			<% else %>
        			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
        			<% end if %>
        		<% next %>

        		<% if boardnotice.HasNextScroll then %>
        			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
        		<% else %>
        			[next]
        		<% end if %>
		        </td>
		</tr>
		</table>
	</td>
</tr>
</table><br><br>
<% end if %>

<% set boardnotice = Nothing %>
<%

'공지사항
dim ohope
set ohope = New CHopeBoard

ohope.FPageSize = 10
ohope.FCurrPage = 1
'ohope.list

%>
<!--
<table width="780"  class="a">
<tr bgcolor="#FFFFFF">
	<td width=100 ><a href="/admin/board/10x10_board_list.asp?menupos=417"><font color="red">전체보기</font></a></td>
	<td align="center" >****** 사원 게시판 ******</td>
	<td width=100 align="right"><a href="/admin/board/10x10_board_write.asp"><font color="red">글쓰기</font></a></td>
</tr>
</table>
<table width="780" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form method=post name="noticeform">
<input type="hidden" name="page" value="1">
<tr bgcolor="#FFFFFF">
	<td colspan="5" height="25" align="right">검색결과 : 총 <font color="red"><% = ohope.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="50" align="center">번호</td>
	<td width="100" align="center">글유형</td>
	<td align="center">제목</td>
	<td width="100" align="center">글쓴이</td>
	<td width="100" align="center">날짜</td>
</tr>
<% for i = 0 to (ohope.FResultCount - 1) %>
  <tr bgcolor="#FFFFFF">
    <td width="50" align="center"><%= ohope.FItemList(i).Fidx %></td>
    <td width="100" align="center"><%= ohope.FItemList(i).FGubunName %></td>
	<td>
	<a href="/admin/board/10x10_board_read.asp?idx=<%= ohope.FItemList(i).Fidx %>"><%= ohope.FItemList(i).Ftitle %></a>
	<% if datediff("d",ohope.FItemList(i).Fregdate,now())<8 then %>
	&nbsp;&nbsp;&nbsp;[<font color=red><b>new</b></font>]
	<% end if %>
	</td>
    <td width="100" align="center"><%= ohope.FItemList(i).Fusername %></td>
    <td width="100" align="center"><%= FormatDateTime(ohope.FItemList(i).Fregdate,2) %></td>
  </tr>
<% next %>

</form>
</table>
<br><br>
-->
<%
set ohope = Nothing
%>

<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->