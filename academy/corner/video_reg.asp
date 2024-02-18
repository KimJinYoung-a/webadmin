<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/video_cls.asp"-->

<%
	Dim vIdx, vTitle, vCateCD2, vLecturer, vMakerID, vKeyword, vImage_URL, vImage2_URL, vYoutube_URL, vYoutube_source, vIsUsing, vRegdate

	vIdx = requestcheckvar(request("idx"),32)
	vIsUsing = "N"
	
'// 있는경우에만 쿼리
dim oip
If vIdx <> "" Then
	set oip = new cvideo
		oip.frectidx = vIdx
		oip.video_view()
		
		if oip.ftotalcount > 0 then
			vIdx = oip.foneitem.fidx 
			vTitle = oip.foneitem.ftitle 
			vCateCD2 = oip.foneitem.fcatecd2 
			vLecturer = oip.foneitem.flecturer 
			vMakerID = oip.foneitem.fmakerid
			vKeyword = oip.foneitem.fkeyword 
			vImage_URL = oip.foneitem.fimage_url
			vImage2_URL = oip.foneitem.fimage2_url
			vYoutube_URL = oip.foneitem.fyoutube_url 
			vYoutube_source = oip.foneitem.fyoutube_source 
			vIsUsing = oip.foneitem.fisusing
			vRegdate = oip.foneitem.fregdate
		end if

	set oip = nothing
End IF
%>

<script language="javascript">

	document.domain = "10x10.co.kr";	
	
	//저장
	function video_reg(){

		if(document.frmcontents.title.value==''){
			alert('제목을 입력하셔야 합니다.');
			document.frmcontents.title.focus();
			return false;
		}
		if(document.frmcontents.CateCD2.value==''){
			alert('카테고리를 입력하셔야 합니다.');
			document.frmcontents.CateCD2.focus();
			return false;
		}
//		if(document.frmcontents.lecturer.value==''){
//			alert('강사의 ID를 입력하셔야 합니다.');
//			document.frmcontents.lecturer.focus();
//			return false;
//		}
//		if(document.frmcontents.makerid.value==''){
//			alert('브랜드ID를 입력하셔야 합니다.\n없으면 thefingers01로 입력하세요.');
//			document.frmcontents.makerid.focus();
//			return false;
//		}
		if(document.frmcontents.youtube_url.value==''){
			alert('YouTube URL을 입력하셔야 합니다.');
			document.frmcontents.youtube_url.focus();
			return false;
		}
//		if(document.frmcontents.youtube_source.value==''){
//			alert('YouTube 소스를 입력하셔야 합니다.');
//			document.frmcontents.youtube_source.focus();
//			return false;
//		}
		
		<% If vIdx = "" Then %>
		if(document.frmcontents.list_image.value==''){
			alert('리스트 이미지를 선택하셔야 합니다.');
			return false;
		}
		<% End If %>

		frmcontents.submit();		
	}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/corner/video_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>IDX</b><br></td>
		<td>
			<%= vIdx %><input type="hidden" name="idx" value="<%= vIdx %>">
			<% If vIdx <> "" Then %>&nbsp;&nbsp;&nbsp;등록일:<%=vRegdate%><% End If %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>제 목</b><br></td>
		<td>
			<input type="text" name="title" size="80" value="<%=vTitle%>" maxlength="150">
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>카테고리</b><br></td>
		<td>
			<% Call DrawSelectBoxLecCategoryLarge("CateCD2",vCateCD2,"N")%>
		</td>
	</tr>
<!--
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>강사 ID</b><br></td>
		<td>
			<input type="text" name="lecturer" size="80" value="<%=vLecturer%>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>브랜드ID</b><br></td>
		<td>
			<input type="text" name="makerid" size="80" value="<%=vMakerID%>" maxlength="32">
		</td>
	</tr>
-->
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>내용</b><br></td>
		<td>
			<input type="text" name="keyword" size="80" value="<%=vKeyword%>" maxlength="200">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>YouTube URL</b><br></td>
		<td>
			<input type="text" name="youtube_url" size="80" value="<%=vYoutube_URL%>" maxlength="200"><br>
			<font color="red"> ※ 유튜브 : 소스코드 복사 (예 : http://www.youtube.com/embed/qj4rn1I_dC8 ) ※ 유튜브 동영상 URL복사 아님!</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>YouTube 소스</b><p>width와 height값<br>width="705"<br>height="360"<br></td>
		<td>
			<textarea name="youtube_source" rows="12" cols="80"><%=vYoutube_source%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>리스트 이미지</b>
		<br><font color="red">240x160</font>
		</td>
		<td>
			<% if vImage_URL <> "" then %>
			<img src="<%=vImage_URL%>"><br>
			<% end if %>
			<input type="file" name="list_image" size="80" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>리스트 이미지2</b>
		<br><font color="red">180x120</font>
		</td>
		<td>
			<% if vImage2_URL <> "" then %>
			<img src="<%=vImage2_URL%>"><br>
			<% end if %>
			<input type="file" name="list_image2" size="80" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>사용여부</b><br></td>
		<td><select name="isusing">
				<option value="Y" <% if vIsUsing = "Y" then response.write " selected" end if %>>Y</option>
				<option value="N" <% if vIsUsing = "N" then response.write " selected" end if %>>N</option>
			</select>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//수정
			if vIdx <> "" then 
			%>
				<input type="button" value="수정" onclick="video_reg('');" class="button">
			<% 
			'//신규
			else 
			%>
				<input type="button" value="신규저장" onclick="video_reg('');" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

