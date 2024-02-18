<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->

<%



dim i, j
dim onlyitemqa, research
dim newsearch
'==============================================================================
'나의 1:1질문답변
dim boardqna,qadiv
dim page

page = request("page")
if page="" then page=1

set boardqna = New CMyQNA

qadiv = request("qadiv")
onlyitemqa = request("onlyitemqa")
research = request("research")
newsearch = request("newsearch")
if (onlyitemqa="") and (research="") then onlyitemqa="on"
if (newsearch="") and (research="") then newsearch="Y"
boardqna.PageSize = 100
boardqna.CurrPage = page
boardqna.RectQadiv = qadiv
boardqna.ScrollCount = 20

boardqna.SearchNew = newsearch
boardqna.FRectOnlyItemInclude = onlyitemqa

boardqna.getLecQnalist

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script language='javascript'>
function NextPage(p){
	document.location = "lecqna_list.asp?newsearch=<%= newsearch %>&page=" + p;
}
</script>
<h2>사용안하는 메뉴 입니다.</h2>
<table width="720" border="0">
<form method="get" name="qnaform">
<input type="hidden" name="research" value="on">
<tr>
  <td>강좌 문의 미처리 리스트</td>
  <td>

  </td>
  <td align="right"><a href="lecqna_list.asp?newsearch=Y">미처리리스트</a>&nbsp;<a href="lecqna_list.asp?newsearch=N">전체리스트</a></td>
</tr>
</form>
</table>

<table width="720" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="200" align="center">고객명(아이디/주문번호)</td>
    <td width="300" align="center">제목</td>
    <td width="100" align="center">구분</td>
    <td width="100" align="center">상품ID</td>
    <td width="100" align="center">강사명</td>
    <td width="100" align="center">답변자</td>
    <td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
  <tr height="20">
    <td width="200">&nbsp;<%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>/<%= boardqna.results(i).orderserial %>)</td>
    <td width="300">&nbsp;<a href="lecture_qna_board_reply.asp?id=<%= boardqna.results(i).id %>&reffrom=itemqa"><%= db2html(boardqna.results(i).title) %></a></td>
    <td width="100" align="center"><%= boardqna.code2name(boardqna.results(i).qadiv) %></td>
    <td width="100" align="center">
    <% if boardqna.results(i).IsUpchebeasong=true then %>
    	<%= boardqna.results(i).FItemID %>
    <% else %>
    	<font color="#FF3333"><%= boardqna.results(i).FItemID %></font>
    <% end if %>
    </td>
    <td width="100" align="center"><%= boardqna.results(i).FMakerID %></td>
    <td width="100" align="center"><%= boardqna.results(i).replyuser %></td>
    <td width="100" align="center"><%= FormatDate(boardqna.results(i).regdate, "0000-00-00") %></td>
  </tr>
<% next %>
</table>
<table width="720" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align=center height=30>
		<% if boardqna.HasPreScroll then %>
			<a href="javascript:NextPage('<%= boardqna.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + boardqna.StartScrollPage to boardqna.ScrollCount + boardqna.StartScrollPage - 1 %>
			<% if i>boardqna.Totalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if boardqna.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<br><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->