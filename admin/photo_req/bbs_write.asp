<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���
' History : 2012.03.16 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
	Dim writer
	Dim cBoard
	Dim sBrd_Id, sBrd_Name, sBrd_Regdate
	Dim brd_content, arrFileList
	sBrd_Id 		= session("ssBctId")
	sBrd_Name		= session("ssBctCname")
	sBrd_Regdate	= Left(now(),10)

	set cBoard = new Board
		cBoard.fnBoardcontent
%>
<script language="javascript">
/*
//������ ����//
function fTeam(str){
	if(str == "all"){
		document.getElementById('brd_team').style.display = 'none';
		for(var j=0; j<frm.part_sn.length; j++) {
			frm.part_sn[j].checked = false;
		}
	}else{
		document.getElementById('brd_team').style.display = 'block';
	}
}*/

function form_check(){
	var frm = document.frm;

//���� �Է� ����//
	if(frm.brd_subject.value == ""){
		alert("������ �Է��ϼ���");
		frm.brd_subject.focus();
		return false;
	}
//���� ��� ����(�������� ������ ������ �����Ͽ� �˻�)//
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

	var chk3 = 0;
	for(var k=0; k<frm.brd_fixed.length; k++) {
		if(frm.brd_fixed[k].checked) chk3++;
	}
	if(chk3 == "0"){
		alert("�������θ� �����ϼ���");
		return false;
	}

	frm.action = "board_proc.asp";
	frm.submit();
}
function fileupload()
{
	window.open('request_popupload.asp','worker','width=420,height=200,scrollbars=yes');
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

</script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td><img src="/images/icon_arrow_link.gif"></td>
	<td style="padding-top:3">&nbsp;<b>�Խñ� �ۼ�</b></td>
</tr>
</table>
<form name="frm"  method="post">
<input type = "hidden" name = "mode" value = "add">
<input type = "hidden" name = "brd_sn" value = "<%=cBoard.Fbrd_sn + 1%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No.<%= cBoard.Fbrd_sn + 1 %></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sBrd_Name%>(<%=sBrd_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �����: <%=sBrd_Regdate%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="brd_subject" value="" size="95" maxlength="128">
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
									<input type='hidden' name='doc_file' value='<%=arrFileList(0,i)%>'>
									<img src='http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
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
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">���� ����</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<label id="brd_useynY"><input type="radio" name="brd_fixed" id="brd_useynY" value="1">Y</label>&nbsp;&nbsp;&nbsp;
				<label id="brd_useynN"><input type="radio" name="brd_fixed" id="brd_useynN" value="2">N</label><br>
				<font color = "RED"> �ذ��� ���� Y�� �����Ͻø� �Խñ��� �ֻ�ܿ� ��ġ�ϰ� �˴ϴ�.</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<table width="813" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left"><a href="board_list.asp"><img src="/images/icon_list.gif" border="0"></a></td>
	<td width="50%" align="right">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="image" src="/images/icon_confirm.gif" border="0" onclick="form_check();"></td>
		</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
