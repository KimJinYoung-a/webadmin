<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'==========================================================================
'	Description: ���� ���
'	History: 2011.03.17		
'==========================================================================
	Dim iMaxLength , sPosition ,sActionLink
	 
	iMaxLength 	= requestCheckVar(Request("iML"),10)		'�ִ�ũ�� 
	sPosition		=  requestCheckVar(Request("sP"),10)		'������ġ
	sActionLink	=requestCheckVar(Request("sAL"),100)		'ó���ּ�
	IF iMaxLength = "" THEN iMaxLength = 10

%>
	<script language="javascript">
	<!-- 
		function jsSubmit(){
			var frm = document.frmImg;
		
			arrFileExt = new Array();
			arrFileExt[arrFileExt.length]  = "XLS";
			arrFileExt[arrFileExt.length]  = "PPT";
			arrFileExt[arrFileExt.length]  = "DOC";
			arrFileExt[arrFileExt.length]  = "RTF";
			arrFileExt[arrFileExt.length]  = "RTF";
			arrFileExt[arrFileExt.length]  = "XLSX";
			arrFileExt[arrFileExt.length]  = "PPTX";
			arrFileExt[arrFileExt.length]  = "DOCX";
			arrFileExt[arrFileExt.length]  = "HWP";
			arrFileExt[arrFileExt.length]  = "PDF";
			arrFileExt[arrFileExt.length]  = "TXT";
			arrFileExt[arrFileExt.length]  = "ZIP";
			arrFileExt[arrFileExt.length]  = "RAR";
			arrFileExt[arrFileExt.length]  = "7Z";
			arrFileExt[arrFileExt.length]  = "CAB";
			arrFileExt[arrFileExt.length]  = "ALZ";
			arrFileExt[arrFileExt.length]  = "XML";
			arrFileExt[arrFileExt.length]  = "GIF";
			arrFileExt[arrFileExt.length]  = "JPG";
			arrFileExt[arrFileExt.length]  = "JPEG";
		
			//���� Ȯ��
			if( frm.sFile.value =="") {
				alert("������ �Է��Ͻʽÿ�.");
				frm.sFile.focus();
				return;
			}
						
			//������ȿ�� üũ
			if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
				alert("�̹����� <%=iMaxLength%>MB������  �����Ǵ� ������ ���ϸ� ���ε� �����մϴ�.\n\n �����Ǵ� ���������� �����ڿ��� �������ּ���");
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
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%" class="a">
	<tr>
		<td height="30"  >���ϵ��<br><hr width="100%"> </td>		
	</tr>
	<tr>
		<td valign="top">		
			<form name="frmImg" method="post" action="<%= uploadImgUrl&sActionLink %>"  enctype="MULTIPART/FORM-DATA">
			<input type="hidden" name="iML" value="<%=iMaxLength%>">  
			<input type="hidden" name="sP" value="<%=sPosition%>">  			
			<table width="100%" border="0" cellpadding="5" cellspacing="10" class="a"> 	
				<tr>
					<td valign="top">���ϸ�:</td>
					<td><input type="file" name="sFile" size="30" class="input"> <br><br>
						<font size="1">(�����Ǵ� ���� ���� : XLS,PPT,DOC,RTF,XLSX,PPTX,DOCX<br>,HWP,PDF,TXT,ZIP,RAR,7Z,CAB,ALZ,JPG,JPEG,GIF,XML)</font>
					</td>
				</tr>				
				<tr>
					<td align="center" colspan="2"><input type="button" class="button" value="���" onclick="jsSubmit();"></td>
				</tr> 
			</table>
			</form>	
		</td>
	</tr>	 
</table>	
</body>
</html>
			
			