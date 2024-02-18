<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 작가/강사 프로필
' History : 2016.08.17 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->
<%
Dim lecturer_id, lecturer_name, history, history_act, socname, newImage_profile, newProfileFolder
Dim socname_kor, regdate, isusing ,homepage, onesentence

lecturer_id = requestcheckvar(request("lecturer_id"),32)

'// 있는경우에만 쿼리
Dim oip
set oip = new cgood_onelist
	oip.FRectlecturer_id = 	lecturer_id

	If lecturer_id <> "" Then
		oip.fgood_edit()
		If oip.FTotalcount > 0 Then
			If application("Svr_Info")="Dev" Then
				newProfileFolder = "http://testimage.thefingers.co.kr/corner/newImage_profile/"
			Else
				newProfileFolder = "http://image.thefingers.co.kr/corner/newImage_profile/"
			End If
			lecturer_name		= oip.foneitem.flecturer_name
			history				= oip.foneitem.fhistory
			history_act			= oip.foneitem.fhistory_act
			socname				= oip.foneitem.fsocname
			socname_kor			= oip.foneitem.fsocname_kor
			onesentence			= doubleQuote(oip.foneitem.FOnesentence)
			newImage_profile	= oip.foneitem.fnewImage_profile
			regdate				= oip.foneitem.fregdate
			isusing				= oip.foneitem.fisusing
			homepage			= oip.foneitem.fhomepage
		End If
	End If

Dim oitemvideo
Set oitemvideo = New cgood_onelist
oitemvideo.FRectArtistid = lecturer_id
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetArtistProfileVideo
%>
<script language="javascript">
document.domain = "10x10.co.kr";

function FnLecturerApp(str){
	var varArray;
	varArray = str.split(',');

	document.frmcontents.lecturer_id.value = varArray[0];
	document.frmcontents.lecturer_name.value = varArray[1];
	document.frmcontents.socname.value = varArray[2];
	document.frmcontents.socname_kor.value = varArray[3];
}

function good_reg(mode){
	if(document.frmcontents.lecturer_id.value==''){
		alert('강사를 입력하셔야 합니다.');
		return false;
	}
	if (mode == 'add'){
		frmcontents.mode.value='add';
	}else if(mode == 'edit'){
		frmcontents.mode.value='edit';
	}
	if (confirm("저장 하시겠습니까?")){
		frmcontents.submit();
	}
}

//작가/강사 전체보기 팝업
function pop_lecture(){
    var popwin = window.open("/academy/corner/pop_lecturerList.asp","pop_lecture","width=500,height=700,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/corner/profileimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="mode" >
<tr bgcolor="#FFFFFF">
	<td align="center"><b>작가/강사</b><br></td>
	<td>
		<input type="button" class="button" value="찾기" onclick="pop_lecture();">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>작가/강사ID</b><br></td>
	<td>
		<input type="text" name="lecturer_id" id="lecturer_id" value="<%=lecturer_id%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>작가/강사명</b><br></td>
	<td>
		<input type="text" name="lecturer_name" id="lecturer_name" size="50" value="<%=lecturer_name%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>브랜드(영문)</b><br></td>
	<td>
		<input type="text" name="socname" id="socname" size="50" value="<%=socname%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>브랜드(한글)</b><br></td>
	<td>
		<input type="text" name="socname_kor" id="socname_kor" size="50" value="<%=socname_kor%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>내 한마디</b><br></td>
	<td>
		<input type="text" name="onesentence" size="80" value="<%=onesentence%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>소개글</b><br></td>
	<td>
		<textarea name="history" cols="60" rows="6" class="textarea"><%=history%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>작가/강사활동</b><br></td>
	<td>
		<textarea name="history_act" cols="60" rows="6" class="textarea"><%=history_act%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">
		<strong>프로필이미지</strong>
		<br><font color="red">600x600</font>
	</td>
	<td>
	<% If newImage_profile <> "" Then %>
		<img src="<%= newProfileFolder & newImage_profile %>" width=<%=600/2%> height=<%=600/2%>><br>
		(프로필 이미지는 600X600 픽셀이상으로 지정되어 있습니다.)<br>
		<img src="<%= newProfileFolder & "thumbimg1/t1_" & newImage_profile %>" width=<%=400/2%> height=<%=400/2%>>400&nbsp;
		<img src="<%= newProfileFolder & "thumbimg2/t2_" & newImage_profile %>" width=<%=200/2%> height=<%=200/2%>>200&nbsp;
		<img src="<%= newProfileFolder & "thumbimg3/t3_" & newImage_profile %>" width=<%=100/2%> height=<%=100/2%>>100&nbsp;<br/>
	<% End If %>
		<input type="file" name="newImage_profile" size="32" maxlength="32" class="file">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>사용여부</b><br></td>
	<td><select name="isusing">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><b>프로필동영상</b><br></td>
	<td>
		<textarea name="itemvideo" rows="5" class="textarea" cols="90"><%= db2html(oitemvideo.FOneItem.FvideoFullUrl) %></textarea>
		<p>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)</p>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
	<% '//수정
		If lecturer_id <> "" and lecturer_name <> "" Then
	%>
			<input type="button" value="수정" onclick="good_reg('edit');" class="button">
	<%
		'//신규
		Else
	%>
			<input type="button" value="신규저장" onclick="good_reg('add');" class="button">
	<%	End If %>
</tr>
</form>
</table>
<%
set oip = Nothing
Set oitemvideo = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->