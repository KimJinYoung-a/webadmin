<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	dim lp
	dim maxAddImg

	maxAddImg=5		'추가 이미지 최대수 지정
%>
<script language='javascript'>
<!--
	// 이미지 종류 검사
	function checkImageSuffix (fileInput)
	{
	   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
	   if (!suffixPattern.test(fileInput.value)) {
	     alert('GIF,JEPG,PNG 파일만 가능합니다.');
	     fileInput.focus();
	     fileInput.select();
	     return false;
	   }
	   return true;
	}

	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.titleShort.value)
		{
			alert("짧은 제목을 입력해주십시오.");
			frm.titleShort.focus();
			return false;
		}

		if(!frm.titleLong.value)
		{
			alert("상세(긴) 제목을 입력해주십시오.");
			frm.titleLong.focus();
			return false;
		}

		if(!frm.imgIcon.value && !frm.imgFile[0].value)
		{
			alert("목록에 보여질 아이콘을 위해\n목록 이미지 혹은 첨부 이미지중 하나를 반드시 선택해주십시오.\n\n※ 이미지는 JPG, GIF형식으로 선택해주십시요.");
			return false;
		}

		for(var i=0;i<frm.imgFile.length;i++)
		{
			if ((frm.imgFile[i].value.length>0)&&(!checkImageSuffix(frm.imgFile[i]))){
				return false;
			}
		}

		// 폼 전송
		return true;

	}

	// 추가 이미지폼 제어
	function addImgControl(lid, md)
	{
		var frm = document.all;
		var btext = "";

		if(md=="add")
		{
			// 레이어 보이기 및 내용 활성화
			frm["addimg"+(lid+1)].style.display="";
			frm["btn"+(lid)].innerHTML = "";

			lid++;
			
			// 버튼처리
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='이미지 삭제' onClick=addImgControl(" + lid + ",'del') style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='이미지 추가' onClick=addImgControl(" + lid + ",'add') style='cursor:pointer'>";
			
			frm["btn"+lid].innerHTML = btext;
		}
		else if(md=="del")
		{
			// 레이어 숨기기 및 내용 삭제
			frm.imgFile[lid-1].select();
			document.execCommand('Delete');
			frm["addimg"+lid].style.display="none";
			frm["btn"+lid].innerHTML = "";

			lid--;

			// 버튼처리
			if(lid>1)
				btext += "<img src='/images/icon_minus.gif' alt='이미지 삭제' onClick=addImgControl(" + lid + ",'del') style='cursor:pointer'> ";
			if(lid<<%=maxAddImg%>)
				btext += "<img src='/images/icon_plus.gif' alt='이미지 추가' onClick=addImgControl(" + lid + ",'add') style='cursor:pointer'>";

			frm["btn"+lid].innerHTML = btext;
		}
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyStory.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>마디 이야기 신규 등록</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>짧은 제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="titleShort" size="40" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>상세 제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="titleLong" size="80" maxlength="120"></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">목록 이미지</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="imgIcon" size="60"><br>
		<font color=darkred>※ 50px * 50px 크기의 JPG/GIF 이미지 파일을 업로드하십시오.</font>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">내용 설명</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" rows="2" cols="80"></textarea></td>
</tr>
<%
	'// 추가 이미지 폼 작성
	for lp=1 to maxAddImg
%>
<tr id="addimg<%=lp%>" <% if lp>1 then Response.Write "style='display:none;'"%>>
	<td align="center" width="120" bgcolor="#DDDDFF">이미지 #<%=lp%></td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="file" name="imgFile" size="50"></td>
			<td width="36" align="right" valign="bottom" id="btn<%=lp%>">
				<% if lp=1 then %><img src="/images/icon_plus.gif" alt="이미지 추가" onClick="addImgControl(<%=lp%>,'add')" style="cursor:pointer"><% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% next %>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->