<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop_noticecls.asp" -->
<%

dim i, j

'==============================================================================
'공지사항
dim boardnotice
set boardnotice = New CNoticeDetail

boardnotice.read(request("idx"))

%>
<script>
function SubmitForm()
{
        if (document.f.gubun.value == "") {
                alert("보여주기 구분을 선택하세요.");
                return;
        }
        if (document.f.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }

        document.f.submit();
}
function SubmitDelete()
{
        if (confirm("삭제하시겠습니까?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>
<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080" class="a">
<form method="post" name="f" action="offshop_notice_act.asp" onsubmit="return false" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="idx" value="<%= request("idx") %>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="username" value="<%=session("ssBctCname")%>">
<input type="hidden" name="mode" value="modify">
<tr>
	<td align="center">보여주기 구분</td>
	<td>
		<select name="gubun">
			<option value="" <% if boardnotice.Fgubun = "" then response.write "selected" %>>선택</option>
			<option value="00" <% if boardnotice.Fgubun = "00" then response.write "selected" %>>전체</option>
			<option value="01" <% if boardnotice.Fgubun = "01" then response.write "selected" %>>1F Shop</option>
			<option value="02" <% if boardnotice.Fgubun = "02" then response.write "selected" %>>3F Zoom</option>
			<option value="03" <% if boardnotice.Fgubun = "03" then response.write "selected" %>>3F College</option>
			<option value="04" <% if boardnotice.Fgubun = "04" then response.write "selected" %>>온라인사업팀</option>
			<option value="50" <% if boardnotice.Fgubun = "50" then response.write "selected" %>>매장-전체</option>
			<option value="51" <% if boardnotice.Fgubun = "51" then response.write "selected" %>>매장-직영</option>
			<option value="52" <% if boardnotice.Fgubun = "52" then response.write "selected" %>>매장-수수료</option>
			<option value="53" <% if boardnotice.Fgubun = "53" then response.write "selected" %>>매장-가맹점</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center">제목</td>
	<td><input type="text" name="title" size="50" value="<%= boardnotice.Ftitle %>"></td>
</tr>
<tr>
	<td align="center">내용</td>
	<td><textarea name="contents" cols="80" rows="18"><%= db2html(boardnotice.Fcontents) %></textarea></td>
</tr>
<tr>
	<td align="center">파일</td>
	<td><input type="file" name="file" size="30"></td>
</tr>
<% if boardnotice.Ffile <> "" then %>
<tr>
	<td align="center">파일삭제</td>
	<td><input type="checkbox" name="dl_file"><%= boardnotice.Ffile %></td>
</tr>
<% end if %>
<tr><td colspan="2" align="right"><input type="button" value=" 수정 " onclick="SubmitForm()">
<input type="button" value=" 삭제 " onclick="SubmitDelete()">&nbsp;&nbsp;</td></tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->