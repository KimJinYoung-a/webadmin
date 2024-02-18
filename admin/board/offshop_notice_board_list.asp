<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshop_noticecls.asp" -->
<%

dim i, j
dim page

page = request("pg")
if (page = "") then
        page = "1"
end if

'==============================================================================
'공지사항
dim boardnotice
set boardnotice = New CBoardNotice

boardnotice.PageSize = 30
boardnotice.CurrPage = CInt(page)
boardnotice.ScrollCount = 100

boardnotice.list

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
고객센타 - 공지사항<br><br>

<table width="600" border="1">
  <tr>
    <td width="300">제목</td>
    <td width="100">유효시작일</td>
    <td width="100">유효종료일</td>
    <td width="100">등록일</td>
  </tr>
<% for i = 0 to (boardnotice.ResultCount - 1) %>
  <tr>
    <td width="300"><a href="offshop_notice_board_modify.asp?id=<%= boardnotice.results(i).id %>"><%= boardnotice.results(i).title %></a></td>
    <td width="100"><%= boardnotice.results(i).yuhyostart %></td>
    <td width="100"><%= boardnotice.results(i).yuhyoend %></td>
    <td width="100"><%= boardnotice.results(i).regdate %></td>
  </tr>
<% next %>
</table>
<table width="600" border="1">
  <tr>
    <td>
<% for i = 0 to (boardnotice.TotalPage - 1) %>
      <a href="offshop_notice_board_list.asp?pg=<%= (i+1) %>"><%= (i+1) %></a>
<% next %>
    </td>
  </tr>
</table>
<br><br>

<a href="offshop_notice_board_write.asp">글쓰기</a>
<!-- #include virtual="/lib/db/dbclose.asp" -->