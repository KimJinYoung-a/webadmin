<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->
<%

dim i, j
dim itemqanotinclude, research

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
boardqna.RectQadiv = qadiv
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
<table width="760" border="0">
<form method="get" name="qnaform">
<input type="hidden" name="research" value="on">
<tr>
  <td>1:1 상담 미처리 리스트</td>
  <td>
  	<input type="checkbox" name="itemqanotinclude" <% if itemqanotinclude="on" then response.write "checked" %> >상품문의제외
  	&nbsp;&nbsp;&nbsp;&nbsp;
  	질문유형선택 :
		  <select name="qadiv">
			<option value="">선택</option>
			<option value="00" <% if qadiv="00" then response.write "selected" %> >배송문의</option>
			<option value="01" <% if qadiv="01" then response.write "selected" %> >주문문의</option>
			<option value="02" <% if qadiv="02" then response.write "selected" %> >상품문의</option>
			<option value="03" <% if qadiv="03" then response.write "selected" %> >재고문의</option>
			<option value="04" <% if qadiv="04" then response.write "selected" %> >취소,환불문의</option>
			<option value="06" <% if qadiv="06" then response.write "selected" %> >교환문의</option>
			<option value="08" <% if qadiv="08" then response.write "selected" %> >사은품문의</option>
			<option value="10" <% if qadiv="10" then response.write "selected" %> >시스템문의</option>
			<option value="12" <% if qadiv="12" then response.write "selected" %> >개인정보관련</option>
			<option value="13" <% if qadiv="13" then response.write "selected" %>>당첨문의</option>
			<option value="14" <% if qadiv="14" then response.write "selected" %>>반품문의</option>
			<option value="15" <% if qadiv="15" then response.write "selected" %>>입금문의</option>
			<option value="16" <% if qadiv="16" then response.write "selected" %>>오프라인문의</option>
			<option value="20" <% if qadiv="20" then response.write "selected" %> >기타문의</option>
		  </select>&nbsp;<input type="submit" value="검색">

  </td>
  <td align="right"><a href="cscenter_qna_board_all_list.asp">전체리스트</a></td>
</tr>
</form>
</table>

<table width="760" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="200" align="center">고객명(아이디/주문번호)</td>
    <td width="300" align="center">제목</td>
    <td width="100" align="center">구분</td>
    <td width="50" align="center">Site</td>
    <td width="50" align="center">레벨</td>
    <td width="160" align="center">작성일</td>
  </tr>
<% for i = 0 to (boardqna.ResultCount - 1) %>
  <tr height="20">
    <td width="200">&nbsp;<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><font color="#000000"><%= boardqna.results(i).username %>(<%= boardqna.results(i).userid %>/<%= boardqna.results(i).orderserial %>)</font></a></td>
    <td width="300">&nbsp;<a href="cscenter_qna_board_reply.asp?id=<%= boardqna.results(i).id %>"><%= db2html(boardqna.results(i).title) %></a></td>
    <td width="100" align="center"><%= boardqna.code2name(boardqna.results(i).qadiv) %></td>

    <td width="50" align="center">
    <% if IsNull(boardqna.results(i).Fextsitename) then %>
    	&nbsp;
    <% else %>
    	<%= boardqna.results(i).Fextsitename %>
    <% end if %>
	</td>
	<td width="50" align="center"><%= boardqna.results(i).GetUserLevelStr %></td>
    <td width="160" align="left"><%= boardqna.results(i).regdate %></td>
  </tr>
<% next %>
</table>
<br><br>

<!-- #include virtual="/lib/db/dbclose.asp" -->