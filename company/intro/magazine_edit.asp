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
	Dim page, SearchArea, SearchKeyword, brdsn

	brdDiv = 2					'�Խ��� ���� (1:��к���, 2:��������)
	brdsn = Request("brdsn")
	page = Request("page")
	SearchArea = Request("SearchArea")
	SearchKeyword = Request("SearchKeyword")
	if page="" then page=1


	'// ���� ����
	dim oBoard, lp
	Set oBoard = new CBoard
	oBoard.FRectBrdSn = brdsn
	
	oBoard.getBoardCont
%>
<!-- ��ܶ� ���� -->
<script language="javascript">
<!--
	function OnInitialize()
	{
		// ������ //
		<%
			oBoard.getBoardFile()

			'// ������ ���� ��� ��� ����
			if oBoard.FResultCount>0 then
				for lp=0 to oBoard.FResultCount-1
		%>
		frm_upload.TABSFileup.AddUploadedFile("<%=oBoard.FfileList(lp).Ffile_sn%>", "<%=oBoard.FfileList(lp).Ffile_name %>", <%=oBoard.FfileList(lp).Ffile_size%>, "<%="http://imgstatic.10x10.co.kr/company/pr/" & oBoard.FfileList(lp).Ffile_name %>");
		<%
				next
			end if
		%>
	}

	// ���˻� �� ����
	function submitForm()
	{
		var form = document.frm_upload;

		if(!form.brd_subject.value)
		{
			alert("������ �Է����ֽʽÿ�.");
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
			alert("������ �ۼ����ֽʽÿ�.");
			form.brd_content.focus();
			return;
		}

		if(confirm("�Է��� �������� �����Ͻðڽ��ϱ�?"))
		{
			// ������ ���� ����
			var UploadFiles = form.TABSFileup.UploadFiles;

		    form.TABSFileup.AddFormValue(form.brdDiv.name, form.brdDiv.value);
		    form.TABSFileup.AddFormValue(form.mode.name, "modi");
		    form.TABSFileup.AddFormValue(form.brdsn.name, form.brdsn.value);
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

	//�Խù� ����
	function deleteItem()
	{
		var form = document.frm_upload;
		if(confirm("�� �Խù��� �����Ͻðڽ��ϱ�?\n\n�س����� ������ �����Ǹ� ���� �� �� ���� �˴ϴ�."))
		{
			// ������ ���� ����
			var UploadFiles = form.TABSFileup.UploadFiles;

		    form.TABSFileup.AddFormValue(form.mode.name, "del");
		    form.TABSFileup.AddFormValue(form.brdsn.name, form.brdsn.value);
	
		    form.TABSFileup.PostMultipartFormData();
		}
	}
//-->
</script>
<script language="JavaScript" src="/js/file_upload.js"></script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm_upload" method="post" action="">
<input type="hidden" name="retURL" value="magazine_list.asp?menupos=<%= menupos %>&page=<%=page%>&SearchArea=<%=SearchArea%>&SearchKeyword=<%=SearchKeyword%>">
<input type="hidden" name="brdDiv" value="<%=brdDiv%>">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="brdsn" value="<%=brdsn%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchArea" value="<%=SearchArea%>">
<input type="hidden" name="SearchKeyword" value="<%=SearchKeyword%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><b>�Խù� �󼼺���/����</b></td>
	<td align="right">&nbsp;</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ��ܶ� �� -->
<!-- ���� ���� ���� -->
<table width="750" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td width="70" bgcolor="#E6E6E6" align="center">��ȣ</td>
	<td width="180" bgcolor="#FFFFFF"><b><%=brdSn%></b></td>
	<td width="70" bgcolor="#E6E6E6" align="center">�ۼ���</td>
	<td width="180" bgcolor="#FFFFFF">
		<%=oBoard.FitemList(1).Fuserid%>
		<input type="hidden" name="userid" value="<%=oBoard.FitemList(1).Fuserid%>">
	</td>
	<td width="70" bgcolor="#E6E6E6" align="center">��ȸ��</td>
	<td width="180" bgcolor="#FFFFFF"><%=oBoard.FitemList(1).Fbrd_hit%></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">����</td>
	<td bgcolor="#FFFFFF" colspan="5"><input type="text" name="brd_subject" size="80" value="<%=oBoard.FitemList(1).Fbrd_subject%>"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">����</td>
	<td bgcolor="#FFFFFF" colspan="5">
	<% 
		'�������� �ʺ�� ���̸� ����
		dim editor_width, editor_height, brd_content
		editor_width = "95%"
		editor_height = "320"
		brd_content = oBoard.FitemList(1).Fbrd_content
	%>
	<!-- #INCLUDE Virtual="/lib/util/editor.inc" -->
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" rowspan="2" align="center">÷������</td>
	<td bgcolor="#FFFFFF" colspan="5">
		<table width="610" class="a" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
			    <script language="javascript">TabsEmbed('modi','TABSFileup','100%',120,'<%=uploadUrl%>/linkweb/company/board_process.asp','�̹�������(*.jpg;*.gif;*.png;*.bmp)|*.jpg;*.gif;*.png;*.bmp|��������(*.doc;*.hwp;*.ppt;*.txt)|*.doc;*.hwp;*.ppt;*.txt|',1,'#FAFAFF')</script>
			    <SCRIPT FOR="TABSFileup" Event="CompletedPostMultipartFormData(ErrType, ErrCode, ErrText)" language="javascript">
			    	var retURL = document.frm_upload.retURL.value;
			    	OnCompletedPostMultipartFormData(ErrType, ErrCode, ErrText,retURL);
			    </SCRIPT>
				<SCRIPT FOR="TABSFileup" Event="ChangingUploadFile(TotalCount, TotalFileSize)" language="javascript">
					OnChangingUploadFile(TotalCount, TotalFileSize);
				</SCRIPT>
			    <SCRIPT FOR="TABSFileup" Event="Initialize()" language="javascript">
			        OnInitialize();
			    </SCRIPT>
			</td>
		</tr>
		<tr>
			<td>* ������ ����� ���Ƶ� �߰��� �� �ֽ��ϴ�.<br>* ȭ�� 1���� �ִ� 2�ް���, 10������ �ѹ��� ���ε� �����մϴ�.</td>
			<td align="right">
				<img src="http://fiximage.10x10.co.kr/images/button_imgup.gif" width="56" height="20" onClick="addFiles()" style="cursor:hand" align="absbottom">
				<img src="http://fiximage.10x10.co.kr/images/button_imgdel.gif" width="56" height="20" hspace="5" onClick="removeFiles()" style="cursor:hand" align="absbottom"><br>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!-- ���� ���� �� -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absbottom"></a>&nbsp;
				<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absbottom"></a>&nbsp;
				<a href="javascript:deleteItem();"><img src="/images/icon_delete.gif" width="45" border="0" align="absbottom"></a>
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
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->