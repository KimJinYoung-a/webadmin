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
dim masteridx,lectureid

idx = request("idx")
mode = request("mode")
masteridx = request("masteridx")
lectureid = request("lectureid")

if idx="" then idx=0
set olec = new CLectureDetail
olec.GetLectureIMGDetail idx

%>
<script language="JavaScript">
<!--
function CheckForm(){
	if (document.lecform.mode.value.length < 1){
		alert("다시 접속하세요");
		history.back();
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
<form method=post name="lecform" action="http://partner.10x10.co.kr/admin/lecture/dolecturedayuserimg.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="masteridx" value="<% = masteridx %>">
<input type="hidden" name="lectureid" value="<% = lectureid %>">
<table width="800" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >Idx</td>
	<td bgcolor="#FFFFFF"> <% = olec.Fidx %></td>
</tr>
<% if mode = "add" then %>
<tr bgcolor="#DDDDFF">
	<td >이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="img1" value="" size="60" class="input_b"></td>
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
	<td >이미지</td>
	<td bgcolor="#FFFFFF"><input type="file" name="img1" value="" size="60" class="input_b"> (<% = olec.Fimg1 %>)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>사용여부</td>
	<td bgcolor="#FFFFFF">
	&nbsp;&nbsp;&nbsp;
	<input type="radio" name="isusing" value="Y" <% if olec.Fisusing = "Y" then response.write "checked" %>> 사용중
	<input type="radio" name="isusing" value="N" <% if olec.Fisusing = "N" then response.write "checked" %>>사용안함
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