<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/noticefaqcls.asp"-->
<%
Dim mode, page, idx
Dim oboard
Dim gubun, subject, contents, isusing

mode	= request("mode")
menupos	= request("menupos")
idx		= request("idx")

If mode = "" Then mode = "I"

If mode = "U" Then
	SET oboard = new cNoticeFAQ
		oboard.FRectIdx = idx
		oboard.getNoticeModify()

		gubun		= oboard.FItemList(0).FGubun
		subject		= oboard.FItemList(0).FSubject
		contents	= oboard.FItemList(0).FContents
		isusing		= oboard.FItemList(0).FIsusing
	SET oboard = nothing
End If
%>
<!-- �̳���� ��ũ��� JS -->
<script language="javascript" type="text/javascript">
	var g_arrSetEditorArea = new Array();
	g_arrSetEditorArea[0] = "EDITOR_AREA_CONTAINER";
</script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/customize.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/customize_ui.js"></script>
<script language="javascript" type="text/javascript" src="/lib/util/innoditor/js/loadlayer.js"></script>
<script language="javascript" type="text/javascript">
	//�̳���Ϳ��� ���ε� �� URL����
	//Fd�� ����� ������ �Ķ��Ÿ�� �ѱ�� webimage���� ������ ���������Ѵ�.///webimage/innoditor/�Ķ��Ÿ��
	var g_strUploadImageURL = "/lib/util/innoditor/pop_upload_img.asp?Fd=SCM_notice";

	// ũ��, ���� ������
	g_nEditorWidth = 800;
	g_nEditorHeight = 1000;
</script>
<script language="javascript" type="text/javascript">
function brdSubmit(frm){
	<% If mode = "I" Then %>
	if(frm.gubun.value==""){
		alert('������ �����ϼ���');
		frm.gubun.focus();
		return false;
	}
	<% End If %>
	if(frm.subject.value==""){
		alert('������ �Է��ϼ���');
		frm.subject.focus();
		return false;
	}
	// �̳���ͷ� ������ ���� textarea�� �Ҵ� ����
	var strHTMLCode = fnGetEditorHTMLCode(true, 0);
	if(strHTMLCode == ''){
		alert("������ �Է��ϼ���");	
		return false;
	}else{
		frm["contents"].value = strHTMLCode;	
	}
	// �̳���ͷ� ������ ���� textarea�� �Ҵ� ��
}
</script>
<!-- �̳���� ��ũ��� JS �� -->
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td><img src="/images/icon_arrow_link.gif"></td>
	<td style="padding-top:3">&nbsp;<b>�Խñ� �ۼ�</b></td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<form name="frm"  method="post" action="board_process.asp" onSubmit="return brdSubmit(this);" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<textarea name="contents" rows="0" cols="0" style="display:none"><%= ChkIIF(mode="U", contents, "") %></textarea> <!-- ���� �̳���� �������� ���� ����Ǵ� �κ�(�����Ϳ� ������ ���� textarea�� stlye:none���� ���� -->
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<% If mode = "U" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
			<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= idx %></td>
		</tr>
		<% End If %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<% If mode = "I" Then %>
				<select name="gubun" class="select">
					<option value="">-Choice-</option>
					<option value="1">��������</option>
					<option value="2">FAQ</option>
				</select>
				<%
					Else 
						Select Case gubun
							Case "1" response.write "��������"
							Case "2" response.write "FAQ"
						End Select
					End If
				%>
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="subject" value="<%= subject %>" size="95" maxlength="128">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
			<td bgcolor="#FFFFFF" style="padding: 5 4 2 5">
				<div id="EDITOR_AREA_CONTAINER"></div>
			</td>
		</tr>
		<% If mode = "U" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�������</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="radio" name="isusing" value="Y" <%= ChkIIF(isusing = "Y", "checked", "") %>>Y
				<input type="radio" name="isusing" value="N" <%= ChkIIF(isusing = "N", "checked", "") %>>N
			</td>
		</tr>
		<% End If %>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="2" align="right">
				<input type="image" src="/images/icon_save.gif">
				<a href="notice_list.asp?menupos=<%=menupos%>"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<script>
	var strHTMLCode = document.frm["contents"].value;
	fnSetEditorHTMLCode(strHTMLCode, false, 0);
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->