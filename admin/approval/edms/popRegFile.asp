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
'	History: 2011.02.25
'==========================================================================
	Dim   iMaxLength, iedmsidx
	Dim page, icateidx1, icateidx2,menupos 
	iMaxLength 	= requestCheckVar(Request("iML"),10)		'�ִ�ũ�� 
	iedmsidx	= requestCheckVar(Request("ieidx"),10)	 
	
	menupos = requestCheckvar(Request("menupos"),10) 
	page = requestCheckvar(Request("page"),10) 
	icateidx1 = requestCheckvar(Request("icateidx1"),10)
	icateidx2 = requestCheckvar(Request("icateidx2"),10)
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
		<td height="30"  style="padding:0 0 0 10">���ϵ��<br><hr width="100%"> </td>		
	</tr>
	<tr>
		<td valign="top">					
			<table width="100%" border="0" cellpadding="5" cellspacing="10" class="a">
			<form name="frmImg" method="post" action="<%=uploadImgUrl%>/linkweb/edms/procUpload.asp"  enctype="MULTIPART/FORM-DATA">
			<input type="hidden" name="iML" value="<%=iMaxLength%>"> 
			<input type="hidden" name="ieidx" value="<%=iedmsidx%>"> 
			<input type="hidden" name="menupos" value="<%=menupos%>">
			<input type="hidden" name="page" value="<%=page%>">
			<input type="hidden" name="icateidx1" value="<%=icateidx1%>">
			<input type="hidden" name="icateidx2" value="<%=icateidx2%>">  
				<tr>
					<td valign="top">���ϸ�:</td>
					<td><input type="file" name="sFile" ><br>
						<font size="1">(�����Ǵ� ���� ���� : XLS,PPT,DOC,RTF,XLSX,PPTX,DOCX<br>,HWP,PDF,TXT,ZIP,RAR,7Z,CAB,ALZ)</font>
					</td>
				</tr>				
				<tr>
					<td align="center" colspan="2"><input type="button" class="button" value="���" onclick="jsSubmit();"></td>
				</tr>
			</form>	
			</table>
		</td>
	</tr>	
	<tr>
		<td bgcolor="#F8F8F8" width="100%" height="10"></td>
	</tr>	
</table>	
</body>
</html>
			
			