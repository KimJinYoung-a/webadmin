<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/board/lib/classes/boardfaqcls.asp" -->
<%

dim i, j
dim page, divcd

page = request("pg")
if (page = "") then
        page = "1"
end if

divcd = request("divcd")
if (divcd = "") then
        divcd = "01"
end if

'==============================================================================
'자주묻는질문
dim boardfaq
set boardfaq = New CBoardFAQ

boardfaq.PageSize = 20
boardfaq.CurrPage = 1
boardfaq.ScrollCount = 100

boardfaq.SearchDivCode = divcd
boardfaq.SearchSort = ""

boardfaq.list

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("menubar") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>FAQ 관리</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			FAQ 내용을 등록 및 수정 가능합니다.
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<a href="cscenter_faq_board_list.asp?divcd=01"><% if divcd=01 then %><b>회원정보관련 FAQ</b><% else %>회원정보관련 FAQ<% end if %></a>
        	&nbsp;|&nbsp;
        	<a href="cscenter_faq_board_list.asp?divcd=02"><% if divcd=02 then %><b>상품문의 FAQ</b><% else %>상품문의 FAQ<% end if %></a>
        	&nbsp;|&nbsp;
        	<a href="cscenter_faq_board_list.asp?divcd=03"><% if divcd=03 then %><b>주문/결재 FAQ</b><% else %>주문/결재 FAQ<% end if %></a>
        	&nbsp;|&nbsp;
        	<a href="cscenter_faq_board_list.asp?divcd=04"><% if divcd=04 then %><b>취소/반품 FAQ</b><% else %>취소/반품 FAQ<% end if %></a>
        	&nbsp;|&nbsp;
        	<a href="cscenter_faq_board_list.asp?divcd=05"><% if divcd=05 then %><b>기타 FAQ</b><% else %>기타 FAQ<% end if %></a>
        </td>
        <td align="right">
        	<a href="cscenter_faq_board_write.asp"><img src="/images/icon_new_registration.gif" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">분류</td>
		<td width="100">소분류</td>
		<td width="400">FAQ제목</td>
		<td width="100">작성일</td>
		<td width="100">조회수</td>
		<td width="100">사용여부</td>
	</tr>
	<% for i = 0 to (boardfaq.ResultCount - 1) %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= boardfaq.code2name(boardfaq.results(i).divcd) %></td>
		<td><%= boardfaq.code2name(boardfaq.results(i).subcd) %></td>
		<td><a href="cscenter_faq_board_modify.asp?id=<%= boardfaq.results(i).id %>"><%= boardfaq.results(i).title %></a></td>
		<td><%= left(boardfaq.results(i).regdate,10) %></td>
		<td><%= boardfaq.results(i).hitcount %></td>
		<td><font color="<%= yncolor(boardfaq.results(i).isusing) %>"><%= boardfaq.results(i).isusing %></font></td>
	</tr>
	<% next %>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% for i = 0 to (boardfaq.TotalPage - 1) %>
		  	<a href="cscenter_faq_board_list.asp?pg=<%= (i+1) %>"><%= (i+1) %></a>
			<% next %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<!-- #include virtual="/lib/db/dbclose.asp" -->

