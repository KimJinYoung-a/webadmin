<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lectureday_userinfocls.asp"-->
<%
dim idx,mode
dim olec

idx = request("idx")
mode = request("mode")

if idx="" then idx=0
set olec = new CLectureDetail
olec.GetLectureDetail idx

%>
<script language="JavaScript">
<!--
function CheckForm(){
	if (document.lecform.lectureid.value.length < 1){
		alert("강사ID를 등록해주세요");
		document.lecform.lectureid.focus();
	}else if (document.lecform.lecturer.value.length < 1){
		alert("강사명을 등록해주세요");
		document.lecform.lecturer.focus();
	}
	else if (document.lecform.lecturename.value.length < 1){
		alert("강좌명을 등록해주세요");
		document.lecform.lecturename.focus();
	}
	else{
		document.lecform.submit();
	}
}

function popLectureItemList(frm){
	var popwin = window.open('lecregitems.asp','lecitem','width=600,height=500,status=no,resizable=yes,scrollbars=yes');
	popwin.focus();
}
//-->
</script>
<form method=post name="lecform" action="http://partner.10x10.co.kr/admin/lecture/dolecturedayuser.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mode" value="<% = mode %>">
<table width="800" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >Idx</td>
	<td bgcolor="#FFFFFF"> <% = olec.Fidx %></td>
</tr>
<% if mode = "add" then %>
<tr bgcolor="#DDDDFF">
	<td >강사ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectureid" value="" size="40" maxlength="32" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" size="40" maxlength="16" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturename" value="" size="40" maxlength="16" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >title이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="titleimg" value="" size="60" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사아이콘이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="lecimg" value="" size="60" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<input type=radio name=isusing value=Y checked> 사용중
	<input type=radio name=isusing value=N  >사용안함
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>
<% else %>
<tr bgcolor="#DDDDFF">
	<td >강사ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lectureid" size="40" maxlength="32" value="<% = olec.Flectureid %>" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturer" size="40" maxlength="16" value="<% = olec.Flecturer %>" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강좌명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="lecturename" size="40" maxlength="16" value="<% = olec.Flecturename %>" class="input_b"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >title이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="titleimg" value="" size="60" class="input_b"> (<% = olec.Ftitleimg %>)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >강사아이콘이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="lecimg" value="" size="60" class="input_b"> (<% = olec.Flecimg %>)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<input type=radio name=isusing value=Y <% if olec.Fisusing = "Y" then response.write "checked" %>> 사용중
	<input type=radio name=isusing value=N <% if olec.Fisusing = "N" then response.write "checked" %>>사용안함
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="right" height="30"><input type="button" value="내용저장" onclick="CheckForm();return false;">&nbsp;&nbsp;&nbsp;</td>
</tr>
<% end if %>
</table>
</form>
<%
set olec = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->