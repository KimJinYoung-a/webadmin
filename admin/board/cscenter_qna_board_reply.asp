<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
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
'orderinfo.UserOrderInfo (boardqna.results(0).userid)
'orderinfo.UserMinusOrderInfo (boardqna.results(0).userid)
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
INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #E6E6E6; color: #000000; }
.link_kor:active {FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: "굴림"; TEXT-DECORATION: none; font-weight: bold}
.link_kor:visited {FONT-SIZE: 9pt; COLOR: #555555; FONT-FAMILY: "굴림"; TEXT-DECORATION: none; font-weight: normal}
.link_kor:hover {FONT-SIZE: 9pt; COLOR: #FF6600; FONT-FAMILY: "굴림"; TEXT-DECORATION: none; font-weight: normal}
.link_kor:link {	FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: "굴림"; TEXT-DECORATION: none; font-weight: normal}
-->
</STYLE>
고객센타 - 자주묻는질문<br><br>
<script>
function SubmitForm()
{
        if (document.frm.replytitle.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.frm.replycontents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }

        if (confirm("입력이 정확합니까?") == true) { document.frm.submit(); }
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
<table width="700" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form method=post name="updateform" action="doeditqadiv.asp">
<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
<tr>
	<td align="right" bgcolor="#FFFFFF">
		  질문유형수정 :
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
		  <option value="13" <% if boardqna.results(0).qadiv = "13" then response.write "selected" %>>당첨문의</option>
		  <option value="14" <% if boardqna.results(0).qadiv = "14" then response.write "selected" %>>반품문의</option>
		  <option value="15" <% if boardqna.results(0).qadiv = "15" then response.write "selected" %>>입금문의</option>
		  <option value="16" <% if boardqna.results(0).qadiv = "16" then response.write "selected" %>>오프라인문의</option>
		  <option value="20" <% if boardqna.results(0).qadiv = "20" then response.write "selected" %>>기타문의</option>
		  </select>
		  <!--
		  &nbsp; <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= boardqna.results(i).FItemID %>" class="link_kor" target="_blank" >상품번호</a> : <input type="text" name="itemid" size="10" value="<% = boardqna.results(0).itemid %>">
		  <input type="hidden" name="itemid" value="0">
		  -->
		  &nbsp;<input type="button" value="수정" onclick="updateqadiv();">&nbsp;
		  <% if reffrom = "itemqa" then %>
		  <input type="button" value="삭제" onclick="delqadiv();">&nbsp;
		  <% end if %>
	</td>
</tr>
</form>
<form method="post" name="frm" action="cscenter_qna_board_act.asp" onsubmit="return false">
<% if boardqna.results(0).replyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<% else %>
<input type="hidden" name="mode" value="firstreply">
<% end if %>
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="email" value="<%= boardqna.results(0).usermail %>">
<input type="hidden" name="emailok" value="<%= boardqna.results(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.results(0).Fextsitename %>">
<input type="hidden" name="imsitxt">
<tr>
	<td bgcolor="#FFFFFF">&nbsp;▷▷▷ <b><%= boardqna.results(0).title %></b> ◁◁◁</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF">
		  <table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
				  <td class="a" height="5">&nbsp;<b>작성자</b> : <%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>) <b>이메일</b> : <%= db2html(boardqna.results(0).usermail) %></td>
				 <td align="right" class="a"><b>작성일</b> : <%= boardqna.results(i).regdate %>&nbsp;</td>
				</tr>
		  </table>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF">
		  <table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
				  <td class="a" height="5">&nbsp;<b>고객등급 : <%= getUserLevelStrByDate(boardqna.results(0).fUserLevel, left(boardqna.results(i).regdate,10)) %></td>
				 <td align="right" class="a"><b>Site</b> : <%= boardqna.results(0).FextSiteName %>&nbsp;</td>
				</tr>
		  </table>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF">
		&nbsp;<b>내용</b> : <br>
			 <table border="0" cellpadding="0" cellspacing="0" width="90%" align="center">
			 <tr>
				 <td><font color="#464646"><%= nl2br(db2html(boardqna.results(0).contents)) %></font></td>
			 </tr>
			 </table>
		<br><br>
	</td>
</tr>
 <% if boardqna.results(0).replyuser<>"" then %>
<tr>
	<td bgcolor="#FFFFFF"><input type="text" name="replytitle" size="30" value="<%= boardqna.results(0).replytitle %>"></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF"><textarea name="replycontents" cols="80" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea></td>
</tr>
<% Else %>
<tr>
	<td bgcolor="#FFFFFF">
		  <input type="text" name="replytitle" size="30">&nbsp;
		  <% SelectBoxQnaPreface "01" %>&nbsp;
		  <% SelectBoxQnaCompliment "" %>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF"><textarea name="replycontents" cols="80" rows="10"></textarea></td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#FFFFFF" align="right">
		  <input type="button" value=" 답글달기 " onclick="SubmitForm()">
		  <% if reffrom="itemqa" then %>
		  <a href="itemqna_list.asp"  class="link_kor">목록으로 이동</a>
		  <% else %>
		  <a href="cscenter_qna_board_list.asp"  class="link_kor">목록으로 이동</a>
		  <% end if %>
	</td>
</tr>
</form>
<tr>
	<td bgcolor="#FFFFFF">
		  <% if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then %>
		  <table cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC" width="100%">
		  <tr>
			  <td colspan="4" bgcolor="#DDDDFF" align="center">예전 질문한 목록</td>
		  </tr>
		  <tr bgcolor="#FFFFFF">
			  <td align="center">제목</td>
			  <td width="80" align="center">구분</td>
			  <td width="80" align="center">사이트</td>
			  <td width="100" align="center">작성일</td>
		  </tr>
		  <% if myqnalist.ResultCount < 0 then %>
		  <% else %>
		  <% for i = 0 to (myqnalist.ResultCount - 1) %>
		  <tr bgcolor="#FFFFFF">
			  <td><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>" class="link_kor">
		    	<% if myqnalist.results(i).DispYn="N" then %>
		    		<strike><%= myqnalist.results(i).title %></strike>
		    	<% else %>
		    		<%= myqnalist.results(i).title %>
		    	<% end if %>
		    	</a></td>
			  <td align="center"><%= myqnalist.code2name(myqnalist.results(i).qadiv) %></td>
			  <td align="center">&nbsp;<%= myqnalist.results(i).FExtSiteName %></td>
			  <td align="center"><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
		  </tr>
		  <% next %>
		  <% end if %>
		  </table>
		  <% end if %>
	</td>
</tr>
</table>
<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/admin/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.results(0).userid %>&masterid=01";
 }

 function TnChangeText(str){
var basictext;
basictext = "안녕하세요. <%= boardqna.results(0).userid %>님\n"
basictext = basictext + "텐바이텐 고객행복센터 <%= session("ssBctCname") %>입니다.\n"
basictext = basictext + "(내용)\n"
basictext = basictext + "만족스런답변이 되셨는지요\n\n"

	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		document.frm.replycontents.value = str;
	}
 }
<% if boardqna.results(0).replyuser <> "" then %>
<% else %>
TnChangeText('');
<% end if %>
//-->
</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

 function TnChangeCompliment(SelectGubun){
	ComplimentFrame.location.href="/admin/board/compliment_select.asp?masterid=01&gubun=" + SelectGubun;
 }

 function TnChangeText2(str){

	if(str == ''){
	}
	else{
		document.frm.replycontents.value = document.frm.imsitxt.value + "\n" + str;
	}
 }
//-->
</script>
<%
set myqnalist = Nothing
set boardqna = Nothing
set orderinfo = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->