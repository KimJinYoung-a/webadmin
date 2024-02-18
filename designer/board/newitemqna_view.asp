<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 주문 클래스
' Hieditor : 2009.04.17 이상구 생성
'			 2016.07.19 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim id
	id= requestCheckvar(request("id"),10)

dim itemqna
set itemqna = new CItemQna
	itemqna.FRectID = id
	itemqna.FRectMakerid = session("ssBctID")
	itemqna.getOneItemQna

if itemqna.FResultCount<=0 then
	Call Alert_return("삭제되었거나 잘못된 문의번호입니다.")
	Response.End
end if

if IsNull(itemqna.FOneItem.FContents) then
	itemqna.FOneItem.FContents = ""
end if

if (itemqna.FOneItem.FContents = "") then
	itemqna.FOneItem.FContents = "(내용없음)"
end if

%>
<script type="text/javascript">

function ActReply(frm){
	var userid, username;
	userid = "<%= Replace(itemqna.FOneItem.Fuserid, Chr(34), "") %>";
	username = "<%= Replace(itemqna.FOneItem.Fusername, Chr(34), "") %>";

	if(frm.replycontents.value.length < 1){
		alert("답변 내용을 적어주세야 합니다.");
		frm.replycontents.focus();
		return;
	}

	if (userid.length>1){
		if (frm.replycontents.value.indexOf(userid) >= 0) {
			alert("입력불가!!\n\n고객 아이디를 답변내용에 입력할 수 없습니다.");
			return;
		}
	}
	if (username.length>1){
		if (frm.replycontents.value.indexOf(username) >= 0) {
			alert("입력불가!!\n\n고객 이름을 답변내용에 입력할 수 없습니다.");
			return;
		}
	}

	if(confirm("상품에 대해 답변 하시겠습니까?")){
		frm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="650" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr height="10" valign="bottom">
    <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td>
    	<img src="/images/icon_arrow_down.gif" align="absbottom">
    	<b>작성자</b> : <%= itemqna.FOneItem.Fusername %>(<%= itemqna.FOneItem.Fuserid %>)
    	&nbsp;&nbsp;&nbsp;&nbsp;
    	<b>고객등급</b> : <%= getUserLevelStrByDate(itemqna.FOneItem.fUserLevel, left(itemqna.FOneItem.Fregdate,10)) %>
    </td>
    <td align="right">
    	<b>작성일</b> : <%= itemqna.FOneItem.Fregdate %>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<table width="650" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=post action="newitemqna_process.asp">
<input type="hidden" name="id" value="<%= itemqna.FOneItem.Fid %>">
<% if itemqna.FOneItem.IsReplyOk then %>
<input type="hidden" name="mode" value="reply">
<% else %>
<input type="hidden" name="mode" value="firstreply">
<% end if %>
<tr bgcolor="#FFFFFF">
	<td width="100" align="center"><b>고객문의<br>상품정보</b>
	<td>
		<!-- 상품정보 -->
		<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
			<tr valign="top">
				<td rowspan="3" width="105"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><img src="<%= itemqna.FOneItem.Flistimage %>" border="0"></a>
				<td>
					상품코드 : <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><%= itemqna.FOneItem.FItemID %></a><br>
					상품명 : <a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FOneItem.FItemID %>" target="_blank"><%= itemqna.FOneItem.FItemName %></a><br>
					브랜드 : <%= itemqna.FOneItem.FMakerid %>(<%= itemqna.FOneItem.FbrandName %>)<br>
					가격 : <%= FormatNumber(itemqna.FOneItem.FSellcash,0) %><br>
				</td>
			</tr>
		</table>
		<!-- 상품정보 -->
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>문의내용</b></td>
	<td>
		<%= nl2br(Replace(itemqna.FOneItem.FContents, "<", "&lt;")) %>
	</td>
</tr>
<% if (FALSE) then %>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>이메일수신요청</b></td>
	<td>
		<input type="radio" name="emailok" value="Y" checked>발송 <input type="radio" name="emailok" value="N">미발송
	</td>
</tr>
<% end if %>
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	<td colspan="2">&nbsp;<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<b><font color="red">답변작성</font></b></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=center><b>답변내용</b></td>
	<td>
		<textarea name="replycontents" cols="80" rows="8" class="input_01"><%= (itemqna.FOneItem.FReplyContents) %></textarea>
		<br><br>
		* 답변 작성시 <font color=red>고객이름, 고객아이디 입력불가</font>(개인정보 유출의 우려가 있습니다.)
		<br>&nbsp;
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="650" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
    	<a href="javascript:ActReply(frm);"><img src="/images/icon_reply.gif" border="0" align="absbottom"></a>
    	<a href="/designer/board/newitemqna_list.asp"><img src="/images/icon_list.gif" border="0" align="absbottom"></a>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
</form>


<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
