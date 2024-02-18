<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/myqnacls.asp" -->
<%

dim i, j
dim reffrom
reffrom = request("reffrom")

'==============================================================================
'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("id"))

if boardqna.results(0).userid <> "" then
dim orderinfo
set orderinfo = New CMyQNAOrderInfo
orderinfo.UserOrderInfo (boardqna.results(0).userid)
orderinfo.UserMinusOrderInfo (boardqna.results(0).userid)
end if

if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then
dim myqnalist
set myqnalist = New CMyQNA
if boardqna.results(0).userid <> "" then
myqnalist.SearchUserID = boardqna.results(0).userid
end if
if boardqna.results(0).orderserial <> "" then
myqnalist.SearchOrderSerial = boardqna.results(0).orderserial
end if
myqnalist.PageSize = 20
myqnalist.CurrPage = 1
myqnalist.list
end if
%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
고객센타 - 자주묻는질문<br><br>
<script>
function SubmitForm()
{
        if (document.f.replytitle.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.replycontents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }

        if (confirm("입력이 정확합니까?") == true) { document.f.submit(); }
}

function updateqadiv(){
	if (confirm("수정하시겠습니까?")){
		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("삭제하시겠습니까?")){
		document.delform.submit();
	}
}
</script>
<table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
<form method=post name="updateform" action="doeditqadiv.asp">
<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
<tr>
	<td align="right">질문유형수정 :
		  <select name="qadiv">
			<option>선택</option>
			<option value="00" <% if boardqna.results(0).qadiv = "00" then response.write "selected" %>>배송문의</option>
			<option value="01" <% if boardqna.results(0).qadiv = "01" then response.write "selected" %>>주문문의</option>
			<option value="02" <% if boardqna.results(0).qadiv = "02" then response.write "selected" %>>상품문의</option>
			<option value="03" <% if boardqna.results(0).qadiv = "03" then response.write "selected" %>>재고문의</option>
			<option value="04" <% if boardqna.results(0).qadiv = "04" then response.write "selected" %>>취소,환불문의</option>
			<option value="06" <% if boardqna.results(0).qadiv = "06" then response.write "selected" %>>교환문의</option>
			<option value="08" <% if boardqna.results(0).qadiv = "08" then response.write "selected" %>>사은품문의</option>
			<option value="10" <% if boardqna.results(0).qadiv = "10" then response.write "selected" %>>시스템문의</option>
			<option value="12" <% if boardqna.results(0).qadiv = "12" then response.write "selected" %>>개인정보관련</option>
			<option value="20" <% if boardqna.results(0).qadiv = "20" then response.write "selected" %>>기타문의</option>
		  </select>
		  &nbsp; <a href="http://www.10x10.co.kr/shopping/category.asp?itemid=<%= boardqna.results(i).FItemID %>" target="_blank" >상품번호</a> : <input type="text" name="itemid" size="10" value="<% = boardqna.results(0).itemid %>">
		  <!--
		  <input type="hidden" name="itemid" value="0">
		  -->
		  &nbsp;<input type="button" value="수정" onclick="updateqadiv();">&nbsp;
		  <% if reffrom = "itemqa" then %>
		  <input type="button" value="삭제" onclick="delqadiv();">&nbsp;
		  <% end if %>
		  </td>
</tr>
</form>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
<form method="post" name="f" action="cscenter_qna_board_act.asp" onsubmit="return false">
<% if boardqna.results(0).replyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<% else %>
<input type="hidden" name="mode" value="firstreply">
<% end if %>

<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="email" value="<%= boardqna.results(0).usermail %>">
<input type="hidden" name="emailok" value="<%= boardqna.results(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.results(0).Fextsitename %>">

	<tr>
      <td background="/admin/images/topbar_bg.gif" height="26" valign="middle" colspan="2">
        <div align="left">
          <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr>
              <td>
                <div align="left">&nbsp;<span class="a"><b>☞ <%= boardqna.results(0).title %></b></span></div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr>
      <td class="a" height="5">&nbsp;<b>작성자</b> : <%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>)</td>
	  <td align="right" class="a"><b>작성일</b> : <%= FormatDate(boardqna.results(i).regdate, "0000-00-00") %>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
<% if boardqna.results(0).userid <> "" then %>
    <tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>총주문건수</b> : <%= orderinfo.OrderCount %>&nbsp;<b>총주문금액</b> : <% = FormatNumber(orderinfo.TotalPrice,0) %>원</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>주문취소건수</b> : <%= orderinfo.MOrderCount %>&nbsp;<b>주문취소금액</b> : <% = FormatNumber(orderinfo.MTotalPrice,0) %>원</td>
    </tr>
    <tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
<% end if %>
	<tr>
      <td class="a" height="5" colspan="2">&nbsp;<b>Site</b> : <%= boardqna.results(0).FextSiteName %>&nbsp;</td>
    </tr>
	<tr>
      <td colspan="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td colspan="2" class="a"><b>내용</b> : <br><%= nl2br(db2html(boardqna.results(0).contents)) %><br><br></td>
    </tr>
    <tr>
      <td colspan="2"><hr></td>
    </tr>
    <tr>
      <td colspan="2"><input type="text" name="replytitle" size="30" value="<%= boardqna.results(0).replytitle %>"></td>
    </tr>
    <tr>
      <td colspan="2"><textarea name="replycontents" cols="80" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="right"><input type="button" value=" 답글달기 " onclick="SubmitForm()">
      <a href="lecqna_list.asp">목록으로 이동</a>
      </td>
    </tr>
</form>
</table>
<% if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then %>
<table cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" border="1">
<tr>
	<td colspan="3" bgcolor="#DDDDFF" align="center">예전 질문한 목록</td>
</tr>
<tr>
	<td width="200" align="center">제목</td>
	<td width="100" align="center">구분</td>
	<td width="100" align="center">작성일</td>
</tr>
<tr>
<% if myqnalist.ResultCount < 0 then %>
<% else %>
<% for i = 0 to (myqnalist.ResultCount - 1) %>
	<td><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
	<td align="center"><%= myqnalist.code2name(myqnalist.results(i).qadiv) %></td>
	<td align="center"><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
</tr>
<% next %>
<% end if %>
</table>
<% end if %>
<!--
<% if (boardqna.results(0).emailok = "Y") then %>
<b>이메일</b> : 수신함<br><br>
<% else %>
<b>이메일</b> : 수신안함<br><br>
<% end if %>
-->
<!-- #include virtual="/lib/db/dbclose.asp" -->