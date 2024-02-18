<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/artistboard_cls.asp"-->
<%
Dim CBCont
Dim lecuserid, idx, userid, title, content, regdate, imgurl1, imgurl2,ithread, idepth
idx = requestCheckVar(request("idx"),10)
lecuserid = requestCheckVar(request("lecuserid"),32)
IF idx <> "" THEN
Set CBCont = new CArtistRoomBoard
	CBCont.Fidx = idx
	CBCont.FLecuserid = lecuserid
	CBCont.fnGetContent	
	title   = CBCont.FTitle
	content = CBCont.FContent
	ithread = CBCont.FThread
	idepth  = CBCont.FDepth + 1
Set CBCont = nothing
END IF
%>
<script language='javascript'>
<!--
	
	function jsSubmit(frm){
	if(fnChkBlank(frm.sT.value)) {
	 	alert("������ �Է��� �ּ���");
	 	frm.sT.focus();
	 	return false;
	}
	
	 if(fnChkBlank(frm.tC.value)) {
		alert("������ �Է��� �ּ���");
		frm.tC.focus();
		return false;
	}
	
	if (frm.sImg1.value) {
		 if (!checkFile(frm.sImg1))  return false;
	 }
	 
	if (frm.sImg2.value) {
		 if (!checkFile(frm.sImg2))  return false;
	 }
}

function fnChkBlank(str)
{
    if (str == "" || str.split(" ").join("") == ""){
        return true;
	}
    else{
        return false;
	}
}	
	
function checkFile(sfile){
  //���� Ȯ���� Ȯ��
   var suffixPattern = /(gif|jpg|jpeg|png)$/i;
	if (!suffixPattern.test(sfile.value)) {
     alert('GIF,JPG ���ϸ� �����մϴ�.');
     sfile.focus();
     sfile.select();
     return false;
   }
   
  //���� ������ Ȯ�� 
  var img = new Image();
  img.dynsrc = sfile.value;
  var filesize = img.fileSize;
	if(filesize > 1024000){
	 alert('����ũ�� �ʰ��Դϴ�. �ִ� 1MB���� �����մϴ�.');
     sfile.focus();
     sfile.select();
     return false;
	}
	
	return true;
}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmReg" method="post" action="<%=imgFingers%>/linkweb/artist/procboard.asp" enctype="multipart/form-data" onSubmit="return jsSubmit(this);">      	
<input type="hidden" name="sM" value="I">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sUID" value="thefingers">
<input type="hidden" name="iPT" value="<%=ithread%>">
<input type="hidden" name="iDepth" value="<%=idepth%>">
<input type="hidden" name="retUrl" value="http://webadmin.10x10.co.kr/academy/board/artistboard_list.asp?menupos=978">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxLecturer "lecuserid",lecuserid  %></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">���̵�</td>
	<td bgcolor="#FFFFFF">thefingers</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sT" size="40" maxlength="50" value="<%=title%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF">
	<textarea name="tC" rows="14" cols="80">	

<%IF idx <> "" THEN%>	
************************************************************
> <%=db2html(content)%><%END IF%>
	</textarea>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�̹���÷��1</td>
	<td bgcolor="#FFFFFF">����ũ��� 1MB����,JPG�Ǵ� GIF������ ���ϸ� �����մϴ�.<br>
		������� WIDTH - 400���Ϸ� ������ �ֽñ� �ٶ��ϴ�.<br>
		<input type="file" name="sImg1">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�̹���÷��2</td>
	<td bgcolor="#FFFFFF">����ũ��� 1MB����,JPG�Ǵ� GIF������ ���ϸ� �����մϴ�.<br>
		������� WIDTH - 400���Ϸ� ������ �ֽñ� �ٶ��ϴ�.<br>
		<input type="file" name="sImg2">
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
