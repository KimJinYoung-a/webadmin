<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프샾이용문의
' Hieditor : 2009.04.07 서동석 생성
'			 2011.05.03 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/classes/board/offshopqnacls.asp" -->
<%
dim i, j ,reffrom ,shopid, page, SearchKey, SearchString, param, isNew ,boardqna
dim orderinfo ,myqnalist
	reffrom = request("reffrom")
	menupos = Request("menupos")
	page = Request("page")
	shopid = Request("shopid")
	isNew = Request("isNew")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&shopid=" & shopid & "&isNew=" & isNew & "&menupos=" & menupos

'나의 1:1질문답변
set boardqna = New CMyQNA
	boardqna.frectidx = request("idx")
	boardqna.read()

set orderinfo = New CMyQNAOrderInfo

if boardqna.FItemList(0).userid <> "" then
	orderinfo.UserOrderInfo (boardqna.FItemList(0).userid)
	orderinfo.UserMinusOrderInfo (boardqna.FItemList(0).userid)
end if

set myqnalist = New CMyQNA

if boardqna.FItemList(0).userid <> "" or boardqna.FItemList(0).orderserial <> "" then
	if boardqna.FItemList(0).userid <> "" then
		myqnalist.fSearchUserID = boardqna.FItemList(0).userid
	end if
	if boardqna.FItemList(0).orderserial <> "" then
		myqnalist.fSearchOrderSerial = boardqna.FItemList(0).orderserial
	end if
	
	myqnalist.FPageSize = 20
	myqnalist.FCurrPage = 1
	myqnalist.list
end if
%>

<script language="javascript">

function SubmitForm()
{
      //  if (document.f.replytitle.value == "") {
      //          alert("제목을 입력하세요.");
       //         return;
        //}
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
		document.f.mode.value="del";
		document.f.submit();
	}
}

</script>

<table width="520" border="0" align="left" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA">
<form method="post" name="f" action="offshop_qna_board_act.asp" onsubmit="return false">
<% if boardqna.FItemList(0).replyuser<>"" then %>
	<input type="hidden" name="mode" value="reply">
<% else %>
	<input type="hidden" name="mode" value="firstreply">
<% end if %>
<input type="hidden" name="idx" value="<%= boardqna.FItemList(0).idx %>">
<input type="hidden" name="email" value="<%= boardqna.FItemList(0).usermail %>">
<input type="hidden" name="emailok" value="<%= boardqna.FItemList(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.FItemList(0).Fextsitename %>">
<input type="hidden" name="usercell" value="<%= boardqna.FItemList(0).usercell %>">
<input type="hidden" name="cellok" value="<%= boardqna.FItemList(0).cellok %>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="shopid" value="<%=shopid%>">
<input type="hidden" name="isNew" value="<%=isNew%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="regdate" value="<%= FormatDate(boardqna.FItemList(i).regdate, "0000-00-00") %>">
<input type="hidden" name="brandname" value="<%=boardqna.FItemList(0).brandname%>">
<input type="hidden" name="itemname" value="<%=boardqna.FItemList(0).itemname%>">
<input type="hidden" name="itemid" value="<%=boardqna.FItemList(0).itemid%>">
<input type="hidden" name="title" value="<%=boardqna.FItemList(0).title%>">
<input type="hidden" name="contents" value="<%=boardqna.FItemList(0).contents%>">
<input type="hidden" name="offshopid" value="<%=boardqna.FItemList(0).Fshopid%>">
<tr bgcolor="#FFFFFF">
	<td colspan=2>
	    <div align="left">
		<table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				<div align="left">&nbsp;<span class="a"><b>☞ <%= boardqna.FItemList(0).title %></b></span></div>
			</td>
		</tr>
		</table>
	    </div>
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
	<td>
		<b>작성자</b> : <%= boardqna.FItemList(0).username %>(<%= boardqna.FItemList(0).userid %>/<%= boardqna.FItemList(0).orderserial %>)
	</td>
	<td>
		<b>작성일</b> : <%= FormatDate(boardqna.FItemList(i).regdate, "0000-00-00") %>
	</td>
</tr>
<% if boardqna.FItemList(0).userid <> "" then %>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<b>총주문건수</b> : <%= orderinfo.OrderCount %>&nbsp;<b>총주문금액</b> : <% = FormatNumber(orderinfo.TotalPrice,0) %>원
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<b>주문취소건수</b> : <%= orderinfo.MOrderCount %>&nbsp;<b>주문취소금액</b> : <% = FormatNumber(orderinfo.MTotalPrice,0) %>원
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<b>Site</b> : <%= boardqna.FItemList(0).FextSiteName %>
	</td>
</tr>
<%IF boardqna.FItemList(0).itemid <> 0 THEN %>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<table class="a" cellpadding="1" cellspacing="1">
		<tr bgcolor="#FFFFFF">
			<td><img src="<%=boardqna.FItemList(0).listimage%>"></td>
			<td valign="top"><상품코드>: <%= boardqna.FItemList(0).itemid %><br>
				<상품명> : [<%=boardqna.FItemList(0).brandname%>] <%=boardqna.FItemList(0).itemname%><br>
			</td>
		</tr>
		</table>	
	</td>
</tr>
<%END IF%>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<b>내용</b> : <br><%= nl2br(db2html(boardqna.FItemList(0).contents)) %><br><br>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<% if (boardqna.FItemList(0).emailok = "Y") then %>
			<b>이메일</b> : 수신함
		<% else %>
			<b>이메일</b> : 수신안함
		<% end if %>
		,&nbsp;
		<% if (boardqna.FItemList(0).cellok = "Y") then %>
			<b>SMS수신</b> : 수신함(H.P. <%=boardqna.FItemList(0).usercell%>)
		<% else %>
			<b>SMS수신</b> : 수신안함
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="left" colspan=2>
		<textarea name="replycontents" cols="80" rows="10"><%= db2html(boardqna.FItemList(0).replycontents) %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value=" 답글달기 " onclick="SubmitForm()" class="button">
		<% if reffrom="itemqa" then %>
			<a href="itemqna_list.asp">목록으로 이동</a>
		<% else %>
			<a href="offshop_qna_board_list.asp?page=<%=page & param%>">목록으로 이동</a>
		<% end if %>		
	</td>
</tr>
</form>
</table>
<br>
<table width="300" border="0" align="left" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA">
<% if boardqna.FItemList(0).userid <> "" or boardqna.FItemList(0).orderserial <> "" then %>
<tr>
	<td colspan="3" bgcolor="#DDDDFF" align="center">예전 질문한 목록</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="200" align="center">제목</td>
	<td width="100" align="center">작성일</td>
</tr>

<% if myqnalist.fResultCount > 0 then %>
<% for i = 0 to (myqnalist.fResultCount - 1) %>
<tr bgcolor="#FFFFFF">
	<td><a href="offshop_qna_board_reply.asp?idx=<%= myqnalist.FItemList(i).idx %>&reffrom=<%= reffrom %>&menupos=<%=menupos%>"><%= myqnalist.FItemList(i).title %></a></td>
	<td align="center"><%= FormatDate(myqnalist.FItemList(i).regdate, "0000-00-00") %></td>
</tr>
<% next %>
<% end if %>
<% end if %>
</table>

<%
set boardqna = nothing
set orderinfo = nothing
set myqnalist = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->