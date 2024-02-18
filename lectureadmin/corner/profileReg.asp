<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 작가/강사 프로필
' History : 2016.08.17 김진영 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->
<%
Dim lecturer_id, lecturer_name, history, history_act, socname, newImage_profile, newProfileFolder
Dim socname_kor, regdate, isusing, homepage, onesentence, mode

lecturer_id = session("ssBctId")

'// 있는경우에만 쿼리
Dim oip
set oip = new cgood_onelist
	oip.frectlecturer_id = 	session("ssBctId")

	If lecturer_id <> "" Then
		oip.fgood_edit()
		If oip.ftotalcount > 0 then
			If application("Svr_Info")="Dev" Then
				newProfileFolder = "http://testimage.thefingers.co.kr/corner/newImage_profile/"
			Else
				newProfileFolder = "http://image.thefingers.co.kr/corner/newImage_profile/"
			End If
			mode				= "edit"
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
		Else
			mode				= "add"
			oip.FGood_myInfo()
			If oip.ftotalcount > 0 Then
				lecturer_name	= oip.foneitem.Fcompany_name
				socname			= oip.foneitem.fsocname
				socname_kor		= oip.foneitem.fsocname_kor
			End If
		End If
	End If
Set oip = Nothing

Dim oitemvideo
Set oitemvideo = New cgood_onelist
oitemvideo.FRectArtistid = lecturer_id
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetArtistProfileVideo
%>
<script language="javascript">
//document.domain = "10x10.co.kr";
//저장
function good_reg(mode){
	if(document.frmcontents.lecturer_name.value==''){
		alert('강사명을 입력하셔야 합니다.');
		return false;
	}
	if (!confirm('저장 하시겠습니까?')){
	    return false;    
	}
	if (mode == 'add'){
		frmcontents.mode.value='add';
	}else if(mode == 'edit'){
		frmcontents.mode.value='edit';
	}
	frmcontents.submit();
}

//작품등록&수정
function reg_item(v){
	var reg_item = window.open('/lectureadmin/corner/good_item_list.asp?lecturer_id='+v,'reg_item','width=1024,height=768,scrollbars=yes,resizable=yes');
	reg_item.focus();
}


</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/corner/profileimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="mode" >
	<input type="hidden" name="lecturer_id" value="<%=lecturer_id%>">
	<input type="hidden" name="socname" value="<%=socname%>">
	<input type="hidden" name="islecAdmin" value="Y">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100"><b>작가/강사ID</b><br></td>
		<td><%=lecturer_id%>&nbsp;
			&nbsp;&nbsp;<input type="button" value="작품 등록" class="button" onclick="reg_item('<%= lecturer_id %>')">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>작가/강사명</b><br></td>
		<td><input type="text" name="lecturer_name" value="<%=lecturer_name%>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>브랜드(영문)</b><br></td>
		<td><%=socname%></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>브랜드(한글)</b><br></td>
		<td><input type="text" name="socname_kor" value="<%=socname_kor%>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="30">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>내 한마디</b><br></td>
		<td>
			<input type="text" name="onesentence" size="60" maxlength="60" value="<%=onesentence%>">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>소개글</b><br></td>
		<td>
			<textarea name="history" cols="60" rows="6" class="textarea"><%=history%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>작가/강사활동</b><br></td>
		<td>
			<textarea name="history_act" cols="60" rows="6" class="textarea"><%=history_act%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">
		<b>New 이미지 프로필</b>
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
		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><b>사용여부</b><br></td>
		<td><select name="isusing" class="select">
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
			<input type="button" value="저장" onclick="good_reg('<%=mode%>');" class="button">
		</tr>
</form>
</table>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->