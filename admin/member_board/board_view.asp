<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ��
' History : 2011.02.28 ������ ����
'           2018.07.12 �ѿ�� ����(ISMS���� ����üũ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%
Dim bsn, sbrd_Id, vBoard, page, i, arrFileList, cooperateFile, intLoop, writer, Positsn, sbrd_Type
    bsn 		= requestcheckvar(request("brd_sn"),10)
    sbrd_Id		= requestcheckvar(session("ssBctId"),32)
    page		= requestcheckvar(request("page"),10)
	sbrd_Type	= requestcheckvar(request("brd_type"),3)
    Positsn = session("ssAdminPOSITsn")

If page = "" Then page = 1

set vBoard = new board
	vBoard.FPageSize = 20
	vBoard.FCurrPage = page
	vBoard.Fbrd_sn = bsn
	vBoard.fnBoardview

    if vBoard.FtotalCount < 1 Then
        response.write "<script type='text/javascript'>"
        response.write "    alert('�ش�Ǵ� ���������� �����ϴ�.');"
        response.write "</script>"
        dbget.close() : response.end
    end if

	vBoard.Fbrd_team = Replace(vBoard.Fbrd_team, ",", "<BR>")
	If vBoard.Fbrd_fixed = "1" Then
		vBoard.Fbrd_fixed = "����"
	ElseIF vBoard.Fbrd_fixed = "2" Then
		vBoard.Fbrd_fixed = "�����"
	End If

	''������ ��û �λ��� lovesay999 �̳ɹ� �ΰ�� ������̶� ������ ������ �����ٰ� ��.. (�λ������� �ø������� ���� �ְ�.)
	dim inSaUidARR : inSaUidARR = "wahahajw,lovesay999,jhw7980,icommang"
	if (InStr((","&inSaUidARR&","),","&session("ssBctId")&",")>0) then
	    if (vBoard.FPositsn>8) then
	        Positsn=8 ''��� ���� ������ ������ ������� ..
        end if
    end if

    if vBoard.fposit_sn < Positsn or Positsn="" or isnull(Positsn) then
        response.write "<script type='text/javascript'>"
        response.write "    alert('�ش���� ���� �ִ� ������ �����ϴ�.');"
        response.write "</script>"
        dbget.close() : response.end
    end if

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
<script type="text/javascript" src="/js/jquery-1.11.0.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<form name="frm" action="cooperate_proc.asp" method="post" onSubmit="return checkform(this);">
<input type="hidden" name="bsn" value="<%=bsn%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>�Խñ� ����</b></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=bsn%></td>
			</tr>
			<input type="hidden" name="doc_useyn" value="Y">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fusername%>(<%=vBoard.Fbid%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �����: <%=vBoard.Fbrd_regdate%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=fnBrdType("v", "", vBoard.Fbrd_type, "")%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_team%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��å</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%
					If IsNull(vBoard.FJob_name) = "True" Then
						response.write "�Ϲ�"
					Else
						response.write vBoard.FJob_name&" �̻� ��������"
					End If
					%>&nbsp;
				</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=Replace(vBoard.FPosit_name,"Assistant","������")%>&nbsp;�̻� ��������</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_subject%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ReplaceScript(db2html(vBoard.Fbrd_content))%></td>
			</tr>
			<%
				IF isArray(arrFileList) THEN
			%>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">÷������</td>
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
									<!--�� <a href='<%=arrFileList(0,intLoop)%>' target='_blank'><%'Split(Replace(arrFileList(0,intLoop),"http://",""),"/")(3)%></a>//-->
									�� <span id="<%=intLoop%>" class="a" onClick="filedownload(<%=arrFileList(0,intLoop)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(3)%></span>
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
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=vBoard.Fbrd_fixed%></td>
			</tr>
		</table>
	</td>
</tr>
</table>
<br />
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left"><a href="board_list.asp?menupos=<%=MenuPos%>&brd_type=<%=sbrd_Type%>"><img src="/images/icon_list.gif" border="0"></a></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td style="padding-right:15"></td>
			<td>
				<%
					If (vBoard.Fbid = session("ssBctId") or C_ADMIN_AUTH or C_PSMngPart) Then
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


<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/member_board_admin/member_board_download.asp" target="fileiframe">
<input type="hidden" name="brd_sn" value="<%=bsn%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src="" width="100" height="100" name="fileiframe" frameborder="0" marginheight="0" marginwidth="0"></iframe>



<!-- ####### ��۾��� ####### //-->
<br>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td>
		<img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>���</b>
	</td>
</tr>
</table>
<iframe src="iframe_board_reply.asp?brd_sn=<%=bsn%>&rid=<%=sbrd_Id%>&page=<%=page%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### ��۾��� ####### //-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->