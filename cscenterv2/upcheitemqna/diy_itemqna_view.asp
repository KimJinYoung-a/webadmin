<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/upcheitemqna/diy_item_qnacls.asp"-->
<%
dim id
id= RequestCheckvar(request("id"),10)

dim itemqna
set itemqna = new CItemQna
itemqna.FRectID = id
itemqna.FRectMakerid = session("ssBctID")
itemqna.getOneItemQna

%>
<script language='javascript'>
function ActReply(frm){


	if(frm.replycontents.value.length < 1){
		alert("답변 내용을 적어주세야 합니다.");
		frm.replycontents.focus();
		return;
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
	        	<b>고객등급</b> : <%= itemqna.FOneItem.GetUserLevelStr %>
	        </td>
	        <td align="right">
	        	<b>작성일</b> : <%= itemqna.FOneItem.Fregdate %>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="650" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post action="diy_itemqna_process.asp">
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
					<td rowspan="3" width="105"><img src="<%= itemqna.FOneItem.Flistimage %>">
					<td>
						상품코드 : <%= itemqna.FOneItem.FItemID %><br>
						상품명 : <%= itemqna.FOneItem.FItemName %><br>
						브랜드 : <%= itemqna.FOneItem.FMakerid %><br>
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
			<%= nl2br(itemqna.FOneItem.FContents) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>이메일수신요청</b></td>
		<td>
			<input type="radio" name="emailok" value="Y" checked>발송 <input type="radio" name="emailok" value="N">미발송
		</td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="2">&nbsp;<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<b><font color="red">답변작성</font></b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align=center><b>답변내용</b></td>
		<td>
			<textarea name="replycontents" cols="80" rows="8" class="input_01"><%= (itemqna.FOneItem.FReplyContents) %></textarea><br>
			<font color=red>고객 답변 작성시 고객이름 대신 고객아이디를 사용하세요.(개인정보 유출의 우려가 있습니다.)</font>
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->