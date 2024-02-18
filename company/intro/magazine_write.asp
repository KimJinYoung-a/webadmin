<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCompanyOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/company/board_cls.asp"-->
<%
	Dim brdDiv
	Dim page

	brdDiv = 2					'게시판 구분 (1:언론보도, 2:잡지협찬)
%>
<!-- 상단띠 시작 -->
<script language="javascript">
<!--
	// 폼검사 및 실행
	function submitForm()
	{
		var form = document.frm_upload;

		if(!form.brd_subject.value)
		{
			alert("제목을 입력해주십시오.");
			form.brd_subject.focus();
			return;
		}

		if (sector_1.chk==0){
			form.brd_content.value = editor.document.body.innerHTML;
		}
		else if(sector_1.chk!=3){
			form.brd_content.value = editor.document.body.innerText;
		}
		if(!form.brd_content.value)
		{
			alert("내용을 작성해주십시오.");
			form.brd_content.focus();
			return;
		}

		if(confirm("입력한 내용으로 저장하시겠습니까?"))
		{
			// 서버로 전송 실행
			var UploadFiles = form.TABSFileup.UploadFiles;

		    form.TABSFileup.AddFormValue(form.brdDiv.name, form.brdDiv.value);
		    form.TABSFileup.AddFormValue(form.mode.name, form.mode.value);
		    form.TABSFileup.AddFormValue(form.userid.name, form.userid.value);
		    form.TABSFileup.AddFormValue(form.brd_subject.name, form.brd_subject.value);
		    form.TABSFileup.AddFormValue(form.brd_content.name, form.brd_content.value);
	
		    form.TABSFileup.PostMultipartFormData();
		}
		else
		{		
			return;
		}
	}
//-->
</script>
<script language="JavaScript" src="/js/file_upload.js"></script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm_upload" method="post" action="">
<input type="hidden" name="retURL" value="magazine_list.asp?menupos=<%= menupos %>">
<input type="hidden" name="brdDiv" value="<%=brdDiv%>">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><b>게시물 신규 작성</b></td>
	<td align="right">&nbsp;</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 상단띠 끝 -->
<!-- 메인 내용 시작 -->
<table width="750" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td width="70" bgcolor="#E6E6E6" align="center">제목</td>
	<td bgcolor="#FFFFFF" colspan="5"><input type="text" name="brd_subject" size="80" value=""></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">내용</td>
	<td bgcolor="#FFFFFF" colspan="5">
	<% 
		'에디터의 너비와 높이를 설정
		dim editor_width, editor_height, brd_content
		editor_width = "95%"
		editor_height = "320"
		brd_content = ""
	%>
	<!-- #INCLUDE Virtual="/lib/util/editor.inc" -->
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" rowspan="2" align="center">첨부파일</td>
	<td bgcolor="#FFFFFF" colspan="5">
		<table width="610" class="a" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
			    <script language="javascript">TabsEmbed('write','TABSFileup','100%',120,'<%=uploadUrl%>/linkweb/company/board_process.asp','이미지파일(*.jpg;*.gif;*.png;*.bmp)|*.jpg;*.gif;*.png;*.bmp|문서파일(*.doc;*.hwp;*.ppt;*.txt)|*.doc;*.hwp;*.ppt;*.txt|',1,'#FAFAFF')</script>
			    <SCRIPT FOR="TABSFileup" Event="CompletedPostMultipartFormData(ErrType, ErrCode, ErrText)" language="javascript">
			    	var retURL = document.frm_upload.retURL.value;
			    	OnCompletedPostMultipartFormData(ErrType, ErrCode, ErrText,retURL);
			    </SCRIPT>
				<SCRIPT FOR="TABSFileup" Event="ChangingUploadFile(TotalCount, TotalFileSize)" language="javascript">
					OnChangingUploadFile(TotalCount, TotalFileSize);
				</SCRIPT>
			</td>
		</tr>
		<tr>
			<td>* 파일을 끌어다 놓아도 추가할 수 있습니다.<br>* 화일 1개당 최대 2메가씩, 10개까지 한번에 업로드 가능합니다.</td>
			<td align="right">
				<img src="http://fiximage.10x10.co.kr/images/button_imgup.gif" width="56" height="20" onClick="addFiles()" style="cursor:hand" align="absbottom">
				<img src="http://fiximage.10x10.co.kr/images/button_imgdel.gif" width="56" height="20" hspace="5" onClick="removeFiles()" style="cursor:hand" align="absbottom"><br>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!-- 메인 내용 끝 -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absbottom"></a>
				<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absbottom"></a>
			</td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</form>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->