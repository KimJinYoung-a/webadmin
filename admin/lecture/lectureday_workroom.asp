<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/lectureday_userinfocls.asp"-->
<%
dim lectureid,masteridx
masteridx=request("masteridx")
lectureid=request("lectureid")

dim lec
set lec = new CLectureDetail
lec.GetWorkRoomImage lectureid
%>
<table width="650" border="0" cellpadding="1" cellspacing="1">
	<tr>		
		<td bgcolor="#DDDDFF">�̹���1</td>
		<td bgcolor="#FFFFFF"><img src="<%= lec.Fimg1 %>"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">�̹���2</td>
		<td bgcolor="#FFFFFF"><img src="<%= lec.Fimg2 %>"></td>	
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">�̹���3</td>
		<td bgcolor="#FFFFFF"><img src="<%= lec.Fimg3 %>"></td>	
	</tr>
</table>	

<table width="650" border="0" cellpadding="1" cellspacing="1" bgcolor="#3d3d3d" class="a">
<form name="frm" method="post" action="<%=uploadUrl%>/lectureday/dolectureday_work_room_img.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="lectureid" value="<%= lectureid %>">
	<tr>
		<td bgcolor="#DDDDFF">masteridx</td>
		<td bgcolor="#FFFFFF"><%= masteridx %></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">����ID</td>
		<td bgcolor="#FFFFFF"><input type="text" name="lectureid2" value="<%= lectureid %>"></td>
	</tr>
	<tr>		
		<td bgcolor="#DDDDFF">�̹���1</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img1" size="60" class="input_b"><br><%= lec.Fimg1 %></td>
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">�̹���2</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img2" size="60" class="input_b"><br><%= lec.Fimg2 %></td>
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">�̹���3</td>
		<td bgcolor="#FFFFFF"><input type="file" name="img3" size="60" class="input_b"><br><%= lec.Fimg3 %></td>
	</tr>
	<tr>	
		<td bgcolor="#DDDDFF">�������</td>
		<td bgcolor="#FFFFFF"><input type="radio" name="isusing" value="Y" <% if lec.FIsUsing="Y" then response.write "checked" %>>Y<input type="radio" name="isusing" value="N" <% if lec.FIsUsing="N" then response.write "checked" %>>N</td>
	</tr>
	<tr>	
		<td colspan=2" bgcolor="#FFFFFF" align="center"><input type="submit" value="��������"></td>
	</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->