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
	Dim iMaxLength , sPosition 
	 
	iMaxLength 	= requestCheckVar(Request("iML"),10)		'�ִ�ũ�� 
	sPosition		=  requestCheckVar(Request("sP"),10)		'�ִ�ũ�� 
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
			arrFileExt[arrFileExt.length]  = "CSV";
		
			//���� �Է�Ȯ�� 
			var chkinput = 0; 
			
			for(i=0;i<4;i++){
				if( frm.sFile[i].value !="") {
					chkinput = 1;
				}
			}
		 
		if(chkinput==0){
				alert("������ �Ѱ� �̻� �Է����ּ���");
				frm.sFile[0].focus();
				return;
		}	 
						
			//������ȿ�� üũ
			if (!fnChkFile(frm.sFile[i].value, <%=iMaxLength%>, arrFileExt)){
				alert("�̹����� <%=iMaxLength%>MB������  �����Ǵ� ������ ���ϸ� ���ε� �����մϴ�.\n\n �����Ǵ� ���������� �����ڿ��� �������ּ���");
				return;
			}
	 
		
			frm.submit(); 
			document.all.dvLoad.style.display = "";
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
<form name="frmImg" method="post" action="<%=uploadImgUrl%>/linkweb/board/procUpload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">  
<input type="hidden" name="sP" value="<%=sPosition%>">  				
<table border="0" cellpadding="5" cellspacing="0" width="100%"   class="a">
	<tr>
		<td height="30" >���ϵ��<br><hr width="100%"> </td>		
	</tr>
	<tr>
		<td valign="top">	 
			<table width="100%" border="0" cellpadding="0" cellspacing="5" class="a">  
				<tr>
					<td valign="top">���ϸ�:</td>
					<td><input type="file" name="sFile" ><br>
						<input type="file" name="sFile" ><br>
						<input type="file" name="sFile" ><br>
						<input type="file" name="sFile" ><br>
						<input type="file" name="sFile" > 
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>	
			<div style="font-size:11px;">- �����Ǵ� ��������: XLS,PPT,DOC,RTF,XLSX,PPTX,DOCX<br>
				&nbsp;&nbsp;&nbsp;,HWP,PDF,TXT,ZIP,RAR,7Z,CAB,ALZ,JPG,JPEG,GIF,XML,CSV  </div>
			<div style="padding-top:5px;">- �ִ� <font color="red">10,240KB</font>���� ��ϰ����մϴ�.</div>
		</td>
	</tr>				
	<tr>
		<td align="center" colspan="2"><input type="button" class="button" value="���" onclick="jsSubmit();"></td>
	</tr> 
</table>	
</form>	 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<div id="dvLoad" style="display:none;top:100px;left:50;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">���ε� ó�����Դϴ�. ��ø� ��ٷ��ּ���~~</font></td>
		</tr>
	</table>
</div>
 
			