<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<%
dim lectureid,masteridx
masteridx=request("masteridx")
lectureid=request("lectureid")
%>
<table width="650" border="0" cellpadding="1" cellspacing="1" bgcolor="#3d3d3d" class="a">
<form name="frm" method="post" action="<%=uploadUrl%>/lectureday/dolectureday_work_room_img.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="masteridx" value="">
<input type="hidden" name="lectureid" value="<%= lectureid %>">
	<tr>
		<td bgcolor="#DDDDFF">masteridx</td>
		<td bgcolor="#FFFFFF"><%= masteridx %></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">강사ID</td>
		<td bgcolor="#FFFFFF"><%= lectureid %></td>
	</tr>
	<tr>		
		<td bgcolor="#DDDDFF">이미지1</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img1" size="60" class="input_b"></td>
	<tr>	
		<td bgcolor="#DDDDFF">이미지2</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img2" size="60" class="input_b"></td>
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">이미지3</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img3" size="60" class="input_b"></td>
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">사용유무</td>
		<td bgcolor="#FFFFFF"><input type="radio" name="isusing" value="Y" checked>Y<input type="radio" name="isusing" value="N">N</td>
	</tr>
	<tr>	
		<td colspan=2" bgcolor="#FFFFFF" align="center"><input type="submit" value="내용저장"></td>
	</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->