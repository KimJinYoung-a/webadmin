<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%

dim i, j
dim idx
idx = requestCheckVar(request("idx"),10)

if idx <> "" then
dim boardqna
set boardqna = New CUpcheQnADetail
boardqna.FRectIdx = idx
boardqna.read
end if
%>

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

function updategubun(){
	if (confirm("수정하시겠습니까?")){
		document.updateform.submit();
	}
}

function GotoDel(){
	if (confirm("삭제하시겠습니까?")){
		document.delform.submit();
	}
}
</script>

<form method="post" name="frm" action="upche_qna_board_act.asp" onsubmit="return false">
<table width="600" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if boardqna.Freplyuser<>"" then %>
	<input type="hidden" name="mode" value="reply">
	<% else %>
	<input type="hidden" name="mode" value="firstreply">
	<% end if %>
	<input type="hidden" name="idx" value="<%= boardqna.Fidx %>">
	
	<tr bgcolor="#FFFFFF" >
		<td colspan="2">문의</td>
    </tr>
	<tr>
		<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">질문유형</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="gubun">
				<option value="">선택</option>
				<option value="01" <% if boardqna.Fgubun="01" then response.write "selected" %> >배송문의</option>
				<option value="02" <% if boardqna.Fgubun="02" then response.write "selected" %> >반품문의</option>
				<option value="03" <% if boardqna.Fgubun="03" then response.write "selected" %> >교환문의</option>
				<option value="04" <% if boardqna.Fgubun="04" then response.write "selected" %> >정산문의</option>
				<option value="05" <% if boardqna.Fgubun="05" then response.write "selected" %> >입고문의</option>
				<option value="06" <% if boardqna.Fgubun="06" then response.write "selected" %> >재고문의</option>
				<option value="07" <% if boardqna.Fgubun="07" then response.write "selected" %> >상품등록문의</option>
				<option value="08" <% if boardqna.Fgubun="08" then response.write "selected" %> >이벤트문의</option>
				<option value="20" <% if boardqna.Fgubun="20" then response.write "selected" %> >기타문의</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">작성자</td>
      	<td bgcolor="#FFFFFF">
	      	<%= boardqna.Fusername %>(<%= boardqna.Fuserid %>)
      	</td>
    </tr>
    <tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">작성일</td>
      	<td bgcolor="#FFFFFF">
			<%= FormatDate(boardqna.Fregdate, "0000.00.00") %>
      	</td>
    </tr>
    <tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당자</td>
      	<td bgcolor="#FFFFFF">
			<%= fnGetMemberName(boardqna.Fworkerid) %>
      	</td>
    </tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
      	<td bgcolor="#FFFFFF">
	      	<%= ReplaceBracket(boardqna.Ftitle) %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF" >
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">문의내용</td>
      	<td bgcolor="#FFFFFF">
      		<%= nl2br(ReplaceBracket(db2html(boardqna.Fcontents))) %>
      	</td>
    </tr>
    
    <tr bgcolor="#FFFFFF" >
		<td colspan="2">답변</td>
    </tr>
    <tr bgcolor="#FFFFFF" >
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">답변제목</td>
      	<td bgcolor="#FFFFFF">
      		<%= ReplaceBracket(boardqna.Freplytitle) %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF" >
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">답변내용</td>
      	<td bgcolor="#FFFFFF">
      		<%= nl2br(ReplaceBracket(db2html(boardqna.Freplycontents))) %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td align="center" colspan="2">
	    	<input type="button" class="button" value="목록이동" onClick="javascript:location.href='upche_qna_board_list.asp';">
			<% If boardqna.Freplyn = "" Then %>
	    	<input type="button" class="button" value="수정하기" onClick="javascript:location.href='upche_qna_board_write.asp?idx=<% =idx %>&mode=edit';">
	    	<% End If %>
	    	<input type="button" class="button" value="삭제하기" onClick="javascript:javascript:GotoDel();">
		</td>
    </tr>
</table>
</form>

<form method=post action="upche_qna_board_act.asp" name="delform">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="idx" value="<% =idx %>">
</form>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->