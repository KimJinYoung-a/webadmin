<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 수정
' History : 2012.02.25 박영운 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
	Dim g_MenuPos, writer, arrFileList, i
	IF application("Svr_Info")="Dev" THEN
		g_MenuPos   = "1288"		'### 메뉴번호 지정.
	Else
		g_MenuPos   = "1304"		'### 메뉴번호 지정.
	End If

	Dim mBoard, bsn
	Dim part_sn, level_sn
	Dim brd_content

	bsn = request("brd_sn")
	set mBoard = new Board
		mBoard.Fbrd_sn = bsn
		mBoard.fnBoardmodify
		arrFileList = mBoard.fnGetFileList
		
%>
<script language="javascript">
function form_check(){
	var frm = document.frm;

//제목 입력 여부//
	if(frm.brd_subject.value == ""){
		alert("제목을 입력하세요");
		frm.brd_subject.focus();
		return false;
	}
//내용 등록 여부//
	var chkCont = oEditor.GetHTML(true);
	if (chkCont == "" || chkCont == "<P>&nbsp;</P>")
	{
		alert("내용을 입력해 주세요!");
		return false;
	}
	
	if (chkCont.indexOf("<form")>=0||chkCont.indexOf("&lt;form")>=0) {
	    alert("내용에 form 테그를 입력할 수 없습니다.\nHTML 버튼을 클릭하셔서 <form테그를 제거해주세요.");
	    return false;
	}
	
	if (chkCont.indexOf("</form")>=0||chkCont.indexOf("&lt;/form")>=0) {
	    alert("내용에 form 테그를 입력할 수 없습니다.\nHTML 버튼을 클릭하셔서 </form>테그를 제거해주세요.");
	    return false;
	}

	frm.action = "board_proc.asp";
	frm.submit();
}
function fileupload()
{
	window.open('board_popupload.asp','worker','width=420,height=200,scrollbars=yes');
}
function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;
	
		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}
function filedownload(idx)
{
	filefrm.file_idx.value = idx;
	filefrm.submit();
}
</script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td><img src="/images/icon_arrow_link.gif"></td>
	<td style="padding-top:3">&nbsp;<b>게시글 수정</b></td>
</tr>
</table>
<form name="frm"  method="post">
<input type = "hidden" name = "mode" value = "modify">
<input type = "hidden" name = "brd_sn" value = "<%=bsn%>">
<input type = "hidden" name="fixed" id="fixed" value="<%=mBoard.Fbrd_fixed%>">
<input type = "hidden" name="isusing" id="isusing" value="<%=mBoard.Fbrd_isusing%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No.<%=bsn%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=mBoard.Fbrd_username%>(<%=mBoard.Fbid%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=mBoard.Fbrd_regdate%></td>
		</tr>	
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="<%= mBoard.Fbrd_subject %>" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
			<!-- ##### TABS EDITOR ##### //-->
			<%
				blnUploadFile = false				'첨부파일 사용여부
				editWidth = "100%"					'Editor 너비
				frmNameCont = "brd_content"			'작성내용 폼이름
				editContent = mboard.Fbrd_content			'Editor 내용
			%>
			<!-- #include virtual="/lib/util/tabsEditor/inc_tabsEditor.asp"-->
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0">
						<input type="button" value="파일업로드" class="button" onclick="fileupload();">
					</td>
					<td width="100%" style="padding:3 0 3 10">
						<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
						<%
						IF isArray(arrFileList) THEN
							For i =0 To UBound(arrFileList,2)
						%>
							<tr>
								<td>
									<input type='hidden' name='doc_file' value='<%=arrFileList(1,i)%>'>
									<input type='hidden' name='doc_realfile' value='<%=arrFileList(2,i)%>'>
									<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
									<span class="a" onClick="filedownload(<%=arrFileList(0,i)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,i),"http://",""),"/")(3)%></span>
								</td>
							</tr>							
						<%
							Next
							Response.Write "<input type='hidden' name='isfile' value='o'>"
						Else
						%>
							<tr>
								<td>
									<%

									%>
								</td>
							</tr>
						<% End If %>
						</table>
					</td>					
				</tr>
				</table>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정 여부</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label><input type="radio" onclick="document.getElementById('fixed').value = 1;"  name="brd_fixed" value="1" <% If mBoard.Fbrd_fixed = "1" Then response.write "checked" End If %>>Y</label>&nbsp;&nbsp;&nbsp;
				<label><input type="radio" onclick="document.getElementById('fixed').value = 2;"  name="brd_fixed" value="2" <% If mBoard.Fbrd_fixed = "2" Then response.write "checked" End If %> >N</label><br>
				<font color = "RED"> ※Y를 선택하시면 게시글의 최상단에 위치하게 됩니다.</font>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">게시글 삭제</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('isusing').value = 'Y';" name="brd_isusing" id="brd_isusing" <% If mBoard.Fbrd_isusing = "Y" Then response.write "checked" End If %> value="Y">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('isusing').value = 'N';" name="brd_isusing" id="brd_isusing" <% If mBoard.Fbrd_isusing = "N" Or mBoard.Fbrd_isusing = "" Then response.write "checked" End If %> value="N">N</label><br>
				<font color = "RED"> ※Y를 선택 후 확인버튼 클릭 시 게시글에서 삭제됩니다.</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left"><img src="/images/icon_list.gif" border="0" onclick="location.href = 'board_list.asp'" style="cursor:hand"></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="image" src="/images/icon_confirm.gif" border="0" onclick="form_check();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/photo_req/photo_req_download.asp" target="fileiframe">
<input type="hidden" name="brd_sn" value="<%=bsn%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src="" width="0" height="0" name="fileiframe" width="0" height="0"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
