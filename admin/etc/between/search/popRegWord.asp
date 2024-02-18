<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/between/betsearchcls.asp" -->
<%
Dim idx, i
idx	= request("idx")
If NOT isnumeric(idx) AND idx <> "" Then
	response.write	"<script language='javascript'>" &_
					"	alert('글번호가 잘못 되었습니다');" &_
					"	window.close();" &_
					"</script>"	
End If
Dim vWord
SET vWord = new cSearch
	vWord.FRectIdx = idx
	vWord.getOneLikeWord
%>
<script language="javascript">
function form_check(f){
	f.submit();
}
</script>
<table width="100%" align="center" cellpadding="8" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="sfrm" method="POST" action="/admin/etc/between/search/likeword_process.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<col width="20%" />
<col  />
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2" align="center">
		<strong>추천 검색어</strong>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF" align="center">
	<td>idx</td>
	<td align="left"><%= vWord.FOneItem.Fidx %></td>
</tr>
<tr height="30" bgcolor="#FFFFFF" align="center">
	<td>순서</td>
	<td align="left">
		<select name="rank" class="select">
		<% For i = 1 to 10 %>
			<option value="<%=i%>" <%= Chkiif(vWord.FOneItem.FRank = i, "selected", "") %> ><%= i %></option>
		<% Next %>
		</select>
	</td>
</tr>
<tr height="30" bgcolor="#FFFFFF" align="center">
	<td>검색어</td>
	<td align="left"><input type="text" name="likeword" value="<%= vWord.FOneItem.FLikeword %>"></td>
</tr>
<tr height="30" bgcolor="#FFFFFF" align="center">
	<td>사용유무</td>
	<td align="left">
		<input type="radio" name="isusing" value="Y" <%= Chkiif(vWord.FOneItem.FIsusing = "" OR vWord.FOneItem.FIsusing = "Y", "checked", "") %> >Y
		<input type="radio" name="isusing" value="N" <%= Chkiif(vWord.FOneItem.FIsusing = "N", "checked", "") %> >N
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2" align="center">
		<input type="button" value="저장" onclick="form_check(this.form);" class="button_s">
	</td>
</tr>
</form>
</table>
<% SET vWord = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->