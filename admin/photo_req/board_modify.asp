<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ����
' History : 2012.02.25 �ڿ��� ����
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
		g_MenuPos   = "1288"		'### �޴���ȣ ����.
	Else
		g_MenuPos   = "1304"		'### �޴���ȣ ����.
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

//���� �Է� ����//
	if(frm.brd_subject.value == ""){
		alert("������ �Է��ϼ���");
		frm.brd_subject.focus();
		return false;
	}
//���� ��� ����//
	var chkCont = oEditor.GetHTML(true);
	if (chkCont == "" || chkCont == "<P>&nbsp;</P>")
	{
		alert("������ �Է��� �ּ���!");
		return false;
	}
	
	if (chkCont.indexOf("<form")>=0||chkCont.indexOf("&lt;form")>=0) {
	    alert("���뿡 form �ױ׸� �Է��� �� �����ϴ�.\nHTML ��ư�� Ŭ���ϼż� <form�ױ׸� �������ּ���.");
	    return false;
	}
	
	if (chkCont.indexOf("</form")>=0||chkCont.indexOf("&lt;/form")>=0) {
	    alert("���뿡 form �ױ׸� �Է��� �� �����ϴ�.\nHTML ��ư�� Ŭ���ϼż� </form>�ױ׸� �������ּ���.");
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
	if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?") == true) {
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
	<td style="padding-top:3">&nbsp;<b>�Խñ� ����</b></td>
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
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No.<%=bsn%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=mBoard.Fbrd_username%>(<%=mBoard.Fbid%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �����: <%=mBoard.Fbrd_regdate%></td>
		</tr>	
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="<%= mBoard.Fbrd_subject %>" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
			<!-- ##### TABS EDITOR ##### //-->
			<%
				blnUploadFile = false				'÷������ ��뿩��
				editWidth = "100%"					'Editor �ʺ�
				frmNameCont = "brd_content"			'�ۼ����� ���̸�
				editContent = mboard.Fbrd_content			'Editor ����
			%>
			<!-- #include virtual="/lib/util/tabsEditor/inc_tabsEditor.asp"-->
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">÷������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<table cellpadding="0" cellspacing="0" border="0" class="a">
				<tr>
					<td width="100" valign="top" style="padding:5 0 0 0">
						<input type="button" value="���Ͼ��ε�" class="button" onclick="fileupload();">
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
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label><input type="radio" onclick="document.getElementById('fixed').value = 1;"  name="brd_fixed" value="1" <% If mBoard.Fbrd_fixed = "1" Then response.write "checked" End If %>>Y</label>&nbsp;&nbsp;&nbsp;
				<label><input type="radio" onclick="document.getElementById('fixed').value = 2;"  name="brd_fixed" value="2" <% If mBoard.Fbrd_fixed = "2" Then response.write "checked" End If %> >N</label><br>
				<font color = "RED"> ��Y�� �����Ͻø� �Խñ��� �ֻ�ܿ� ��ġ�ϰ� �˴ϴ�.</font>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�Խñ� ����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('isusing').value = 'Y';" name="brd_isusing" id="brd_isusing" <% If mBoard.Fbrd_isusing = "Y" Then response.write "checked" End If %> value="Y">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_isusing"><input type="radio" onclick="document.getElementById('isusing').value = 'N';" name="brd_isusing" id="brd_isusing" <% If mBoard.Fbrd_isusing = "N" Or mBoard.Fbrd_isusing = "" Then response.write "checked" End If %> value="N">N</label><br>
				<font color = "RED"> ��Y�� ���� �� Ȯ�ι�ư Ŭ�� �� �Խñۿ��� �����˴ϴ�.</font>
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
