<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 뷰
' History : 2011.02.28 김진영 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
Dim writer
Dim bsn, sbrd_Id
bsn 		= request("brd_sn")
sbrd_Id		= session("ssBctId")

Dim vBoard, page, i, arrFileList, cooperateFile, intLoop
page = request("page")

If page = "" Then page = 1

set vBoard = new board
	vBoard.FPageSize = 20
	vBoard.FCurrPage = page
	vBoard.Fbrd_sn = bsn
	vBoard.fnBoardview

	vBoard.Fbrd_team = Replace(vBoard.Fbrd_team, ",", "<BR>")
	If vBoard.Fbrd_fixed = "1" Then
		vBoard.Fbrd_fixed = "고정"
	ElseIF vBoard.Fbrd_fixed = "2" Then
		vBoard.Fbrd_fixed = "비고정"
	End If


set cooperateFile = new board
cooperateFile.Fbrd_sn = bsn
arrFileList = cooperateFile.fnGetFileList
	
%>
<script language="javascript">
function go_modify(str){
	location.href="board_modify.asp?brd_sn="+str;
}
function filedownload(idx)
{
	filefrm.file_idx.value = idx;
	filefrm.submit();
}

</script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<form name="frm" action="cooperate_proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="bsn" value="<%=bsn%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>게시글 내용</b></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=bsn%></td>
			</tr>
			<input type="hidden" name="doc_useyn" value="Y">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fusername%>(<%=vBoard.Fbid%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=vBoard.Fbrd_regdate%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제목</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_subject%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내용</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_content%></td>
			</tr>
			<%
				IF isArray(arrFileList) THEN
			%>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
						<%
						IF isArray(arrFileList) THEN
							For intLoop =0 To UBound(arrFileList,2)
						%>
							<tr>
								<td>
									<input type='hidden' name='doc_file' value='<%=arrFileList(1,intLoop)%>'>
									<input type='hidden' name='doc_realfile' value='<%=arrFileList(2,intLoop)%>'>
									<!--· <a href='<%=arrFileList(0,intLoop)%>' target='_blank'><%'Split(Replace(arrFileList(0,intLoop),"http://",""),"/")(3)%></a>//-->
									<span id="<%=intLoop%>" class="a" onClick="filedownload(<%=arrFileList(0,intLoop)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(4)%></span>
								</td>
							</tr>
						<%
							Next
							Response.Write "<input type='hidden' name='isfile' value='o'>"
						End If
						%>
					</table>
				</td>
			</tr>
			<% End If %>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">고정 여부</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_fixed%></td>
			</tr>
		</table>
	</td>
</tr>
</table>
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left"><img src="/images/icon_list.gif" border="0" onclick="location.href = 'board_list.asp'" style="cursor:hand"></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td style="padding-right:15"></td>
			<td>
				<%
					If (vBoard.Fbid = session("ssBctId") or session("ssBctId") = "teachmeplz") Then
				%>
				<img src="/images/icon_modify.gif" border="0" style="cursor:hand" onclick="go_modify('<%=bsn%>');">
				<%
					End If
				%>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>


<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/photo_req/photo_req_download.asp" target="fileifr2131232ame">
<input type="hidden" name="brd_sn" value="<%=bsn%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src="" width="0" height="0" name="fileiframe" width="0" height="0"></iframe>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->