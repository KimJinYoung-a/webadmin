<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2008.03.13 create
' Description :  �̹��� ÷��
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
 Dim sImgURL, arrImg
 
 sImgURL 	= requestCheckVar(request("sIU"),100)
 
 '//�̹��� �� ����
 IF sImgURL <> "" THEN
 arrImg 	= split(sImgURL,"/")
 sImgURL	= arrImg(Ubound(arrImg))
END IF
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("ã�ƺ��� ��ư�� ���� ���ε��� �̹����� ������ �ּ���.");			
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ������� </div>
�� ���� ������� 300 x 400 (����x����) �� �����մϴ�.
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/sitemaster/uploadEmpImage.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
  <tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�̹�����</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
	</tr>	
	<%IF sImgURL <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">���� �̹����� : <%=sImgURL%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->