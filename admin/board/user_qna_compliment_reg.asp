<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/qna_complimentcls.asp"-->
<%
dim idx, gubun,mode
idx = request("idx")
gubun = request("gubun")
mode = request("mode")

dim omd
set omd = New CMDSRecommend
omd.FCurrPage = 1
omd.FPageSize=1
omd.FRectidx = idx
omd.GetMDSRecommendList

%>
<script language="JavaScript">
<!--

function AddIttems(){
	if (frmarr.gubun.value == ""){
		alert("구분을 선택해주세요!");
		frmarr.gubun.focus();
	}
	else if (frmarr.contents.value == ""){
		alert("내용을 입력해주세요!");
		frmarr.contents.focus();
	}
	else if (confirm('추가하시겠습니까?')){
		frmarr.submit();
	}
}
//-->
</script>
<% if mode = "add" then %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="docompliment.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="masterid" value="02">
<tr>
	<td>
	  분류 : <% SelectBoxQnaComplimentGubun gubun %>
	</td>
</tr>
<tr>
	<td><textarea name="contents" rows="15" cols="50"></textarea></td>
</tr>
<tr>
	<td><input type="button" value="추가" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% else %>
<table border="0" cellpadding="0" cellspacing="0">
<form method="post" name="frmarr" action="docompliment.asp">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<tr>
	<td>
	  분류 : <% SelectBoxQnaComplimentGubun gubun %>
	</td>
</tr>
<tr>
	<td><textarea name="contents" rows="15" cols="50"><% = omd.FItemList(0).Fcontents %></textarea></td>
</tr>
<tr>
	<td><input type="button" value="수정" onclick="AddIttems();" class="button"></td>
</tr>
</form>
</table>
<% end if %>
<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->