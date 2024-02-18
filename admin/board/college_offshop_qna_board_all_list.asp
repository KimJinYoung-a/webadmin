<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshopqnacls.asp" -->
<%

dim i, j, page, rectuserid, qadiv,masterid

rectuserid = request("rectuserid")
page = request("page")
qadiv = request("qadiv")
if page="" then page=1

if session("ssBctDiv") = "101" then
	masterid = "'01','02','05','06','10'"
elseif session("ssBctDiv") = "201" then
	masterid = "'20'"
elseif session("ssBctDiv") = "301" then
	masterid = "'30','31'"
end if

'==============================================================================
'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.PageSize = 20
boardqna.CurrPage = page
boardqna.ScrollCount = 10
boardqna.RectQadiv = masterid
boardqna.SearchUserID = rectuserid


boardqna.list
response.write qadiv
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
function  TnSearch(frm){
	if (frm.rectuserid.length<1){
		alert('검색어를 입력하세요.');
		return;
	}
	frm.method="get";
	frm.submit();
}
function NextPage(ipage){
	document.frmSrc.page.value= ipage;
	document.frmSrc.submit();
}
</script>
<table width="720" border="0">
<tr>
  <td>Offline Shop 질문 </td>
  <td align="right"><a href="college_offshop_qna_board_list.asp">미처리리스트</a></td>
</tr>
<form name="frmSrc" method="get" action="">
<input type="hidden" name="page" value="<% = page %>">
<tr>
  <td colspan="2">
  	아이디 : <input type="text" name="rectuserid" value="<%= rectuserid %>">&nbsp;<input type="submit" value="검색">
  </td>
</tr>
</form>
</table>

<table width="720" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="200" align="center">고객명(아이디/주문번호)</td>
    <td width="100" align="center">구분</td>
    <td width="300" align="center">제목</td>
    <td width="70" align="center">처리유무</td>
    <td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
  <tr height="20">
    <td width="200">&nbsp;<%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>)</td>
    <td width="100" align="center"><%= boardqna.results(i).GetGubunName %></td>
    <td width="300">&nbsp;<a href="college_offshop_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><%= boardqna.results(i).title %></a></td>
    <% if (boardqna.results(i).replyuser="") then %>
    <td width="70">&nbsp;</td>
    <% else %>
    <td width="70" align="center">완료</td>
    <% end if %>
    <td width="100" align="center"><%= FormatDate(boardqna.results(i).regdate, "0000-00-00") %></td>
  </tr>
<% next %>
</table>
<tr>
	<td colspan="5">
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
<br><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->