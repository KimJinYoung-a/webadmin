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
'	Description: 파일 등록
'	History: 2011.02.25
'==========================================================================
	Dim   iMaxLength, iedmsidx
	Dim page, icateidx1, icateidx2,menupos 
	iMaxLength 	= requestCheckVar(Request("iML"),10)		'최대크기 
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
		
			//파일 확인
			if( frm.sFile.value =="") {
				alert("파일을 입력하십시오.");
				frm.sFile.focus();
				return;
			}
						
			//파일유효성 체크
			if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
				alert("이미지는 <%=iMaxLength%>MB이하의  지원되는 형식의 파일만 업로드 가능합니다.\n\n 지원되는 파일형식은 관리자에게 문의해주세요");
				return;
			}
			
			frm.submit();
		}
		
		  function fnChkFile(sFile, sMaxSize, arrExt){   
    //파일 업로드 유무확인
   	 if (!sFile){
    	 return true;
    	}
   
    var blnResult = false;
        
   	//파일 용량 확인
   	var maxsize = sMaxSize * 1024 * 1024;
   	
 	 //	var img = new Image();
	//	img.dynsrc = sFile;
	//var fSize = img.fileSize ;		
		//if (fSize > maxsize){
			//alert("파일크기는 "+sMaxSize+"MB이하만 가능합니다.");
			//return false;
		//}
		
   	//파일 확장자 확인
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
		<td height="30"  style="padding:0 0 0 10">파일등록<br><hr width="100%"> </td>		
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
					<td valign="top">파일명:</td>
					<td><input type="file" name="sFile" ><br>
						<font size="1">(지원되는 파일 포맷 : XLS,PPT,DOC,RTF,XLSX,PPTX,DOCX<br>,HWP,PDF,TXT,ZIP,RAR,7Z,CAB,ALZ)</font>
					</td>
				</tr>				
				<tr>
					<td align="center" colspan="2"><input type="button" class="button" value="등록" onclick="jsSubmit();"></td>
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
			
			