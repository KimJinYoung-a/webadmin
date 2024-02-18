<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  핑거스 강사 게시판
' History : 2010.03.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->

<%
Dim arrFileList, i , olect , lectFile, sDoc_part_sn
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Type, sDoc_Import, sDoc_Subj, sDoc_Content
Dim sDoc_UseYN, sDoc_Regdate , vParam, s_status, s_type, s_ans_ox , g_MenuPos
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")	
	s_status		= NullFillWith(requestCheckVar(Request("s_status"),10),"")
	s_type			= NullFillWith(requestCheckVar(Request("s_type"),10),"")
	s_ans_ox		= NullFillWith(requestCheckVar(Request("s_ans_ox"),1),"")	
	g_MenuPos = RequestCheckvar(request("menupos"),10)
		
	'읽는 페이지에 내용도 저장이 되어서 혹시나 모를 일에 대비하여 파라메터명을 바꿔서 주고 받았슴
	vParam = "doc_status="&s_status&"&doc_type="&s_type&"&ans_ox="&s_ans_ox
	
If iDoc_Idx = "" Then
	sDoc_Id 		= session("ssBctId")
	sDoc_Name		= session("ssBctCname")
	sDoc_Regdate	= Left(now(),10)
	sDoc_WorkerName	= ""
	sDoc_Worker		= ""
Else
		
	Set olect = New clecturer_list
	olect.FrectDoc_Idx = iDoc_Idx
	olect.fnGetlecturerView

	sDoc_Id 		= olect.FOneItem.FDoc_Id
	sDoc_Name		= olect.FOneItem.FDoc_Name
	sDoc_Status		= olect.FOneItem.FDoc_Status
	if sDoc_Status = "" then sDoc_Status = "K001"	
	sDoc_Type		= olect.FOneItem.FDoc_Type
	sDoc_Import		= olect.FOneItem.FDoc_Import
	sDoc_Subj		= olect.FOneItem.FDoc_Subj
	sDoc_Content	= olect.FOneItem.FDoc_Content
	sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
	sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
	sDoc_part_sn	= olect.FOneItem.fpart_sn

	set lectFile = new clecturer_list
 	lectFile.FrectDoc_Idx = iDoc_Idx
	arrFileList = lectFile.fnGetFileList	
End If

%>

<script type="text/javascript">

function checkform(frm){
	frm.submit();
}

</script>

<form name="frm" action="lecturer_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="mode" value="view">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10"> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<% If iDoc_Idx <> "" Then %>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번호</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%></td>
			</tr>
			<% End If %>
			<input type="hidden" name="doc_useyn" value="Y">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">등록자</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%= getthefingers_staff("", sDoc_part_sn, sDoc_Name) %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ 등록일: <%=sDoc_Regdate%>
				</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">현재 상태</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("w","K000",sDoc_Status)%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">구분</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=CommonCode("v","G000",sDoc_Type)%>
					<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
				</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">중요도</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=CommonCode("v","L000",sDoc_Import)%>						
				</td>
			</tr>
			<input type="hidden" name="doc_difficult" value="2">

			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Subj%>
				</td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">내 용</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=sDoc_Content%>
				</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<table cellpadding="0" cellspacing="0" border="0" class="a">
					<tr>
						<td width="100%" style="padding:3 0 3 10">
							<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
							<%
							IF isArray(arrFileList) THEN
								For i =0 To UBound(arrFileList,2)
							%>
								<tr>
									<td>											
										<a href='<%=arrFileList(0,i)%>' target='_blank'>
										<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
									</td>
								</tr>
							<%
								Next
								Response.Write "<input type='hidden' name='isfile' value='o'>"
							Else
							%>
								<tr>
									<td>
									</td>
								</tr>
							<% End If %>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<% If iDoc_Idx <> "" AND sDoc_Id = session("ssBctId") Then %>	
		<input type="button" onclick="checkform(frm);" value="저장하기" class="button">
		<% end if %>
		<input type="button" value="목록으로" onclick="location.href='lecturer.asp?menupos=<%=g_MenuPos%>'" class="button">		
	</td>	
</tr>
</table>

</form>

<% If iDoc_Idx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<table border="0" width="100%" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td>
		<img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>답변</b>
	</td>
</tr>
</table>
<iframe src="iframe_lecturer_ans.asp?didx=<%=iDoc_Idx%>&registid=<%=sDoc_Id%>" name="iframeDB1" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### 답변쓰기 ####### //-->
<% End If %>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

<%
	'session.codePage = 949
%>