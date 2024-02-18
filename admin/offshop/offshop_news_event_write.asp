<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_newscls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script>
function SubmitForm()
{
//alert('수정중입니다.');
//return;
        if (document.f.gubun.value == "") {
                alert("글 유형을 선택하세요.");
                return;
        }
        
        if (document.f.shopid.value == "") {
                alert("샵명을 선택하세요.");
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
        
        if (document.f.enddate.value == "") {
                alert("종료일을 입력하세요.");
                return;
        }
        
        if (confirm('저장 하시겠습니까?')){
            document.f.submit();
        }
}
</script>
<table  border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" width="650" class="a">
<form method="post" name="f" action="<%= uploadImgUrl %>/linkweb/offshop/OffshopNewsEvent_process.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<% ''사무실 직원만 수정가능 %>
<% if (session("ssBctDiv")<10) then %>
<input type="hidden" name="AssignFront" value="on">
<% end if %>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">샵명</td>
		<td bgcolor="white" style="padding:0">
			<% drawSelectBoxOffShopAll "shopid","" %>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">글유형</td>
		<td bgcolor="white" style="padding:0">
			<select name="gubun">
				<option value="">선택</option>
				<%=fnOptCommonCode("noticegubun","")%>
			</select>
		</td>
	</tr>	
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">제목</td>
		<td bgcolor="white" style="padding:0">
				<input name="title" style="width:450" maxlength="40" style="border:1 solid" value="">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">내용</td>
		<td bgcolor="white" style="padding:0">
				<textarea name="contents" cols="50" rows="15" style="border:1 solid"></textarea>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">첨부사진</td>
		<td bgcolor="white" style="padding:0">
				<input type="file" name="file1" size="50" class="input_b">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" bgcolor="#DDDDFF">종료일</td>
		<td bgcolor="white" style="padding:0">
				<input type="text" name="enddate" size="10" maxlength="10" style="border:1 solid" value="">
				<a href="javascript:calendarOpen(f.enddate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a>
				(<%= Left(now(),10) %>)
		</td>
	</tr>
	<tr>
		<td style="padding:0" colspan="2" align="right" bgcolor="white">
			<input type="button" value="Save" onclick="SubmitForm()" style="background-color:#dddddd; height:25; border:1 solid buttonface">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->