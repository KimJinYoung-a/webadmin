<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̼��� ���ڰ�꼭 ���
' History : 2012.02.07 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim  iMaxLength
	IF iMaxLength = "" THEN iMaxLength = 10
%>

	<script language="javascript">
	<!--
		function jsSumbit(){
			var frm = document.frmFile;

			arrFileExt = new Array();
			arrFileExt[arrFileExt.length]  = "XLS";
			arrFileExt[arrFileExt.length]  = "XLSX";

			//���� Ȯ��
			if( frm.sFile.value =="") {
				alert("������ �Է��Ͻʽÿ�.");
				frm.sFile.focus();
				return;
			}

			//������ȿ�� üũ
			if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
				alert("������ <%=iMaxLength%>MB������ XLS,XLSX ���ϸ� ���ε� �����մϴ�.");
				return;
			}

			frm.submit();
		}

		  function fnChkFile(sFile, sMaxSize, arrExt){
    //���� ���ε� ����Ȯ��
   	 if (!sFile){
    	 return true;
    	}

    var blnResult = false;

   	//���� �뷮 Ȯ��
   	var maxsize = sMaxSize * 1024 * 1024;

 	 //	var img = new Image();
	//	img.dynsrc = sFile;
	//var fSize = img.fileSize ;
		//if (fSize > maxsize){
			//alert("����ũ��� "+sMaxSize+"MB���ϸ� �����մϴ�.");
			//return false;
		//}

   	//���� Ȯ���� Ȯ��
   		var pPoint = sFile.lastIndexOf('.');
		var fPoint = sFile.substring(pPoint+1,sFile.length);
		var fExet = fPoint.toLowerCase();

		for (var i = 0; i < arrExt.length; i++)
		   	{
		    	if (arrExt[i].toLowerCase() == fExet)
		    	{
		   			blnResult =  true;
		   		}
			}

		return blnResult;
   }
	//-->
	</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><strong>�鼼 ���ݰ�꼭 ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
			<form name="frmFile" method="post" action="<%=uploadImgUrl%>/linkweb/tax/procNoTaxUpload.asp"  enctype="MULTIPART/FORM-DATA">
			<input type="hidden" name="iML" value="<%=iMaxLength%>">
				<tr>
					<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>"> ����/����  </td>
					<td bgcolor="#FFFFFF"><input type="radio" name="iTST" value="0" checked>���� <input type="radio" name="iTST" value="1" >����</td>
				</tr>
				<tr>
					<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>">���ϸ� </td>
					<td bgcolor="#FFFFFF"><input type="file" name="sFile" class="button"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2"><a href="javascript:jsSumbit();"><font class="text_blue">���</font></a> | <a href="javascript:self.close();">���</a></td>
	</tr>
	</form>
	<tr>
		<td>
			 - ����(.XLS,.XLSX) ���ϸ� ��ϰ����մϴ�.<br>
			 - ��Ʈ���� �⺻��Ʈ����  "sheet1"���� ���ּ���.<br>
			 - ��Ʈ�� ù��°������ �ʵ��(��:�ۼ�����,���ι�ȣ),<br> �ι�°���� �����ͳ���(��:20111203,1234564889)�� ������ ���� �������ּ��� <br>
		</td>
	</tr>
</table>
</body>
</html>

