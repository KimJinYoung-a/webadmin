<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/DiaryEnjoyCls.asp"-->
<%
'###############################################
' PageName : evnet_enjoyWrite.asp
' Discription : �۰����� �׷��� ���/����
' History : 2009.09.30 ������ ����
'###############################################

dim denjSn,mode,i
mode=request("mode")
denjSn=request("denjSn")
%>
<script language="javascript">
<!--
function subcheck(){
	var frm=document.inputfrm;

	if(!frm.makerid.value) {
		alert("�귣��ID�� �Է����ּ���!");
		frm.makerid.focus();
		return;
	}

	if(!frm.subject.value) {
		alert("������ �Է����ּ���!");
		frm.subject.focus();
		return;
	}

	if(!frm.videoSn.value) {
		alert("������ �������� [������ �˻�]�� �̿��Ͽ� �������ּ���!");
		return;
	}

	if(!frm.smallImage.value&&!frm.denjSn.value) {
		alert("�޴��� ǥ���� ���� �̹����� �������ּ���!");
		frm.smallImage.focus();
		return;
	}

	if(!frm.listImage.value&&!frm.denjSn.value) {
		alert("��Ͽ� ǥ���� �̹����� �������ּ���!");
		frm.listImage.focus();
		return;
	}

	if(!frm.introImage.value&&!frm.denjSn.value) {
		alert("�Ұ��� �̹����� �������ּ���!");
		frm.introImage.focus();
		return;
	}

	//if(!frm.bestImage.value&&!frm.denjSn.value) {
	//	alert("����Ʈ ��Ͽ� ǥ���� �̹����� �������ּ���!");
	//	frm.bestImage.focus();
	//	return;
	//}

	frm.submit();
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('�� �Խù��� �����Ͻðڽ��ϱ�?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="<%= uploadImgUrl %>/linkweb/Diary/doDiaryEnjoyProcess.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>�۰����� �׷��� ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<input type="hidden" name="denjSn" value="">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
	<td bgcolor="#FFFFFF">
	    <input type="text" class="text" name="makerid" value="" size="20" >
	    <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="subject" value="" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="videoSn" value="" size="3" readonly>
		<input type="button" class="button" value="������ �˻�" onclick="jsSearchVideoSn(this.form.name,'videoSn','dia');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ �Ⱓ �� ��ǥ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="eventday" value="�Ⱓ:2009.11.16 ~ 11.18 / ��ǥ:11.19" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="smallImage" value="" size="40"> (�� JPG,GIF �̹���, 154px �� 104px, �ִ� 200KB ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="listImage" value="" size="40"> (�� JPG,GIF �̹���, 200px �� 134px, �ִ� 300KB ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�Ұ��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="introImage" value="" size="40"> (�� JPG,GIF �̹���, 276px �� 239x, �ִ� 500KB ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ʈ �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="bestImage" value="" size="40"> (�� JPG,GIF �̹���, 120px �� 120px, �ִ� 200KB ����)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ver.2 ���� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="v2mainImage" value="" size="40"> (�� JPG,GIF �̹���, 186px �� 195px, �ִ� 300KB ����)
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New CEnjoy
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectEnSN=denjSn
	fmainitem.GetDiaryEnjoyList
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FdenjSn%></b>
		<input type="hidden" name="denjSn" value="<%=fmainitem.FItemList(0).FdenjSn%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
	<td bgcolor="#FFFFFF">
	    <input type="text" class="text" name="makerid" value="<%=fmainitem.FItemList(0).Fmakerid%>" size="20" >
	    <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="subject" value="<%=fmainitem.FItemList(0).Fsubject%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="videoSn" value="<%=fmainitem.FItemList(0).FvideoSn%>" size="3" readonly>
		<input type="button" class="button" value="������ �˻�" onclick="jsSearchVideoSn(this.form.name,'videoSn','dia');" >
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ �Ⱓ �� ��ǥ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="eventday" value="<%=fmainitem.FItemList(0).Feventday%>" size="60">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="smallImage" value="" size="40"> (�� JPG,GIF �̹���, 165px �� 115px, �ִ� 200KB ����)
		<%
			if Not(fmainitem.FItemList(0).FsmallImage="" or isNull(fmainitem.FItemList(0).FsmallImage)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FsmallImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="listImage" value="" size="40"> (�� JPG,GIF �̹���, 200px �� 134px, �ִ� 300KB ����)
		<%
			if Not(fmainitem.FItemList(0).FlistImage="" or isNull(fmainitem.FItemList(0).FlistImage)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FlistImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�Ұ��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="introImage" value="" size="40"> (�� JPG,GIF �̹���, 270px �� 246px, �ִ� 500KB ����)
		<%
			if Not(fmainitem.FItemList(0).FintroImage="" or isNull(fmainitem.FItemList(0).FintroImage)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FintroImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ʈ �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="bestImage" value="" size="40"> (�� JPG,GIF �̹���, 120px �� 120px, �ִ� 200KB ����)
		<%
			if Not(fmainitem.FItemList(0).FbestImage="" or isNull(fmainitem.FItemList(0).FbestImage)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FbestImage & ")"
			end if
		%>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ver.2 ���� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" class="text" name="v2mainImage" value="" size="40"> (�� JPG,GIF �̹���, 186px �� 195px, �ִ� 300KB ����)
		<%
			if Not(fmainitem.FItemList(0).Fv2mainimage="" or isNull(fmainitem.FItemList(0).Fv2mainimage)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).Fv2mainimage & ")"
			end if
		%>
	</td>
</tr>
<input type="hidden" name="imgname" value="<%=fmainitem.FItemList(0).FsmallImage%>">
<input type="hidden" name="imgname_1" value="<%=fmainitem.FItemList(0).Fv2mainimage%>">
<% end if %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isUsing" value="Y" <% If mode = "add" Then Response.Write "checked" Else If fmainitem.FItemList(0).Fisusing = "Y" Then Response.Write "checked" End If End If %>> Y
		<input type="radio" name="isUsing" value="N" <% If mode = "edit" Then If fmainitem.FItemList(0).Fisusing = "N" Then Response.Write "checked" End If End If %>> N
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><!--<input type="button" value=" ���� " class="button" onclick="delitems();"> &nbsp;&nbsp;//--><% end if %>
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
