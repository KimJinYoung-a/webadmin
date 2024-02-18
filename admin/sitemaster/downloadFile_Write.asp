<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/downloadFileCls.asp"-->
<%
'###############################################
' PageName : downloadFile_Write.asp
' Discription : ���ϴٿ�ε� ���� ���/����
' History : 2010.05.13 ������ ����
'###############################################

dim fileSn,mode,i
mode		= requestCheckvar(request("mode"),10)
fileSn	= requestCheckvar(request("fileSn"),10)


%>
<script type="text/javascript">
<!--
function editcont(){
    //���µ��� ���� ������ ���;;
    var frm=document.inputfrm;
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.sale_code.value="";
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.fileTitle.value) {
		alert("���� ������ �Է����ּ���!");
		frm.fileTitle.focus();
		return;
	}

	if(!frm.fileDownNm.value) {
		alert("�ٿ�޾����� ����� ���ϸ��� �Է����ּ���!");
		frm.fileDownNm.focus();
		return;
	}

	if(!frm.fName.value&&!frm.fileSn.value) {
		alert("JPG, GIF, ZIP ������ ������ �������ּ���!");
		frm.fName.focus();
		return;
	}

	frm.submit();
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('�� ������ �����Ͻðڽ��ϱ�?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//���Ͼ��ε�
function jsGetFile(){
	var winFile = window.open("/lib/popRegFile.asp?sAL=/linkweb/sitemaster/doDownFileProcess.asp&iML=3","popFile","width=400, height=300");
	winFile.focus();
} 

//���ε� ���ϸ� ��������
function jsSetFile(sfilename, sfilepath, slocation,sfilesize){
		document.inputfrm.fName.value = sfilename;
		document.inputfrm.fSize.value = sfilesize
}
//-->
</script>
<form name="inputfrm" method="post" action="dodownloadFile.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>�ٿ�ε� ���� ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="iEC" value="" size="10" maxlength="10">
	</td>
</tr>
<input type="hidden" name="fileSn" value="">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="fileTitle" value="" size="60" maxlength="32">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� ���ϸ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="fileDownNm" value="" size="60" maxlength="32"> (ex. imagefile_1024.jpg)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="button" class="button" value="����÷��" onClick="jsGetFile();">
		<input type="hidden" name="fSize" value="">
		<input type="text" class="text" name="fName" value="" size="40" readonly> (�� JPG, GIF, ZIP����, �ִ� 3MB ����)
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New cDownFile
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectFSN=fileSn
	fmainitem.GetfileList
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FfileSn%></b>
		<input type="hidden" name="fileSn" value="<%=fmainitem.FItemList(0).FfileSn%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="iEC" value="<%=fmainitem.FItemList(0).Fevt_code%>" size="10" maxlength="10">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="fileTitle" value="<%=fmainitem.FItemList(0).FfileTitle%>" size="60" maxlength="32">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� ���ϸ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="fileDownNm" value="<%=fmainitem.FItemList(0).FfileDownNm%>" size="60" maxlength="32">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="button" class="button" value="����÷��" onClick="jsGetFile();">
		<input type="hidden" name="fSize" value="<%=fmainitem.FItemList(0).Ffilesize%>">
		<input type="text" class="text" name="fName" value="<%=fmainitem.FItemList(0).FfileName%>" size="40" readonly> (�� JPG, GIF, ZIP����, �ִ� 3MB ����)
		<%
			if Not(fmainitem.FItemList(0).FfileName="" or isNull(fmainitem.FItemList(0).FfileName)) then
				Response.Write "<br>(����:" & fmainitem.FItemList(0).FfileName & ")"
			end if
		%>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" ���� " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr> 
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
