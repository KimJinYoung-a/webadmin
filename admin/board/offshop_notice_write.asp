<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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
</script>
<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080" class="a">
<form method="post" name="f" action="offshop_notice_act.asp" onsubmit="return false" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="username" value="<%=session("ssBctCname")%>">
<tr>
	<td align="center">보여주기 구분</td>
	<td>
		<select name="gubun">
			<option value="">선택</option>
			<option value="00">전체</option>
			<option value="01">1F Shop</option>
			<option value="02">2F Zoom</option>
			<option value="03">3F College</option>
			<option value="04">온라인사업팀</option>
			<option value="50">매장-전체</option>
			<option value="51">매장-직영</option>
			<option value="52">매장-수수료</option>
			<option value="53">매장-가맹점</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center">제목</td>
	<td><input type="text" name="title" size="30" value=""></td>
</tr>
<tr>
	<td align="center">내용</td>
	<td><textarea name="contents" cols="50" rows="15"></textarea></td>
</tr>
<tr>
	<td align="center">파일</td>
	<td><input type="file" name="file" size="30"></td>
</tr>
<tr><td colspan="2" align="right"><input type="button" value=" 등록 " onclick="SubmitForm()">&nbsp;&nbsp;</td></tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->