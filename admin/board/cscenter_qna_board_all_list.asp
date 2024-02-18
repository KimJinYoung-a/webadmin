<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->
<%

dim i, j, page, rectuserid, qadiv

rectuserid = request("rectuserid")
page = request("page")
qadiv = request("qadiv")
if page="" then page=1


'==============================================================================
'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.PageSize = 20
boardqna.CurrPage = page
boardqna.ScrollCount = 10
boardqna.RectQadiv = qadiv
boardqna.SearchUserID = rectuserid


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
<table width="760" border="0">
<tr>
  <td>고객센타 - 1:1상담 / 나의 질문과 답 / 전체리스트 </td>
  <td align="right"><a href="cscenter_qna_board_list.asp">미처리리스트</a></td>
</tr>
<form name="frmSrc" method="get" action="">
<input type="hidden" name="page" value="<% = page %>">
<tr>
  <td colspan="2">
  	아이디 : <input type="text" name="rectuserid" value="<%= rectuserid %>">
	&nbsp;&nbsp;질문유형선택 :
		  <select name="qadiv">
			<option value="">선택</option>
			<option value="00" <% if qadiv = "00" then response.write "selected" %>>배송문의</option>
			<option value="01" <% if qadiv = "01" then response.write "selected" %>>주문문의</option>
			<option value="02" <% if qadiv = "02" then response.write "selected" %>>상품문의</option>
			<option value="03" <% if qadiv = "03" then response.write "selected" %>>재고문의</option>
			<option value="04" <% if qadiv = "04" then response.write "selected" %>>취소,환불문의</option>
			<option value="06" <% if qadiv = "06" then response.write "selected" %>>교환문의</option>
			<option value="08" <% if qadiv = "08" then response.write "selected" %>>사은품문의</option>
			<option value="10" <% if qadiv = "10" then response.write "selected" %>>시스템문의</option>
			<option value="12" <% if qadiv = "12" then response.write "selected" %>>개인정보관련</option>
			<option value="13" <% if qadiv = "13" then response.write "selected" %>>당첨문의</option>
			<option value="14" <% if qadiv = "14" then response.write "selected" %>>반품문의</option>
			<option value="15" <% if qadiv = "15" then response.write "selected" %>>입금문의</option>
			<option value="16" <% if qadiv = "16" then response.write "selected" %>>오프라인문의</option>
			<option value="20" <% if qadiv = "20" then response.write "selected" %>>기타문의</option>
		  </select>&nbsp;<input type="submit" value="검색">
  </td>
</tr>
</form>
</table>

<table width="760" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="200" align="center">고객명(아이디/주문번호)</td>
    <td width="300" align="center">제목</td>
    <td width="100" align="center">구분</td>
    <td width="70" align="center">처리유무</td>
    <td width="50" align="center">Site</td>
    <td width="170" align="center">작성일</td>
  </tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
  <tr height="20">
    <td width="200">&nbsp;<%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>/<%= boardqna.results(i).orderserial %>)</td>
    <td width="300">&nbsp;<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><%= boardqna.results(i).title %></a></td>
    <td width="100" align="center"><%= boardqna.code2name(boardqna.results(i).qadiv) %></td>
    <td width="70" align="center">
    <% if (boardqna.results(i).replyuser<>"") then %>
    		<% if boardqna.results(i).dispyn="N" then %>
    		<font color="red">삭제</font>
    		<% else %>
    		완료
    		<% end if %>
    <% else %>
    &nbsp;
    <% end if %>
    </td>
    <td>
    <% if IsNull(boardqna.results(i).Fextsitename) then %>
    	&nbsp;
    <% else %>
    	<%= boardqna.results(i).Fextsitename %>
    <% end if %>
    </td>
    <td width="160" align="left"><%= boardqna.results(i).regdate %></td>
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