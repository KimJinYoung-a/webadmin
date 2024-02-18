<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/offshop_newscls.asp" -->
<%

dim i, j, page

page = request("page")
if page="" then page=1

'==============================================================================
'나의 1:1질문답변
dim offnews
set offnews = New COffshopNewsEvent

offnews.FPageSize = 20
offnews.FCurrPage = page
offnews.FScrollCount = 10
offnews.GetOffshopNewsList

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

<table width="600" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="50" align="center">번호</td>
    <td width="100" align="center">샵</td>
    <td width="100" align="center">구분</td>
    <td align="center">제목</td>
    <td width="50" align="center">작성자</td>
    <td width="100" align="center">작성일</td>
  </tr>
<% for i = 0 to (offnews.FResultCount - 1) %>
  <tr height="20">
    <td align="center">&nbsp;<%= offnews.FItemList(i).Fidx %></td>
    <td align="center"><%= offnews.FItemList(i).Fshopid %></td>
	<td align="center"><%= offnews.FItemList(i).GubunName %></td>
    <td>&nbsp;<a href="offshop_news_event_edit.asp?id=<%= offnews.FItemList(i).Fidx %>"><%= offnews.FItemList(i).Ftitle %></a></td>
    <td align="center">Y</td>
    <td align="center"><%= offnews.FItemList(i).Fuserid %></td>
    <td align="center"><%= FormatDate(offnews.FItemList(i).Fregdate, "0000.00.00") %></td>
  </tr>
<% next %>
</table>
<table width="600" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center">
		<% if offnews.HasPreScroll then %>
			<a href="javascript:NextPage('<%= offnews.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + offnews.StartScrollPage to offnews.FScrollCount + offnews.StartScrollPage - 1 %>
			<% if i>offnews.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if offnews.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<tr>
	<td align="right"><a href="offshop_event_board_write.asp"><font color="red">News & 이벤트 등록</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<br><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->