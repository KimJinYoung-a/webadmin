<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshopqnacls.asp" -->
<%

dim i, j,masterid
dim itemqanotinclude, research

if session("ssBctDiv") = "101" then
	masterid = "'01','02','05','06','10'"
elseif session("ssBctDiv") = "201" then
	masterid = "'20'"
elseif session("ssBctDiv") = "301" then
	masterid = "'30','31'"
end if

'==============================================================================
'나의 1:1질문답변
dim boardqna,qadiv
set boardqna = New CMyQNA

qadiv = request("qadiv")
itemqanotinclude = request("itemqanotinclude")
research = request("research")
if (itemqanotinclude="") and (research="") then itemqanotinclude="on"

boardqna.PageSize = 200
boardqna.CurrPage = 1
boardqna.RectQadiv = masterid
boardqna.ScrollCount = 20

boardqna.SearchNew = "Y"
boardqna.FRectItemNotInclude = itemqanotinclude

boardqna.list

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<table width="720" border="0">
<form method="get" name="qnaform">
<input type="hidden" name="research" value="on">
<tr>
  <td>Offline Shop 상담 미처리 리스트</td>
  <td>&nbsp;</td>
  <td align="right"><a href="college_offshop_qna_board_all_list.asp">전체리스트</a></td>
</tr>
</form>
</table>

<table width="720" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="200" align="center">고객명(아이디/주문번호)</td>
    <td width="100" align="center">구분</td>
    <td width="300" align="center">제목</td>
    <td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
  <tr height="20">
    <td width="200">&nbsp;<%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>/<%= boardqna.results(i).orderserial %>)</td>
    <td width="150" align="center"><%= boardqna.results(i).GetGubunName %></td>
    <td width="300">&nbsp;<a href="college_offshop_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><%= db2html(boardqna.results(i).title) %></a></td>
    <td width="100" align="center"><%= FormatDate(boardqna.results(i).regdate, "0000-00-00") %></td>
  </tr>
<% next %>
</table>
<br><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->