<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이세로 전자계산서 등록
' History : 2012.02.07 정윤정  생성
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

			//파일 확인
			if( frm.sFile.value =="") {
				alert("파일을 입력하십시오.");
				frm.sFile.focus();
				return;
			}

			//파일유효성 체크
			if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
				alert("파일은 <%=iMaxLength%>MB이하의 XLS,XLSX 파일만 업로드 가능합니다.");
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
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><strong>면세 세금계산서 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
			<form name="frmFile" method="post" action="<%=uploadImgUrl%>/linkweb/tax/procNoTaxUpload.asp"  enctype="MULTIPART/FORM-DATA">
			<input type="hidden" name="iML" value="<%=iMaxLength%>">
				<tr>
					<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>"> 매출/매입  </td>
					<td bgcolor="#FFFFFF"><input type="radio" name="iTST" value="0" checked>매입 <input type="radio" name="iTST" value="1" >매출</td>
				</tr>
				<tr>
					<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>">파일명 </td>
					<td bgcolor="#FFFFFF"><input type="file" name="sFile" class="button"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2"><a href="javascript:jsSumbit();"><font class="text_blue">등록</font></a> | <a href="javascript:self.close();">취소</a></td>
	</tr>
	</form>
	<tr>
		<td>
			 - 엑셀(.XLS,.XLSX) 파일만 등록가능합니다.<br>
			 - 시트명은 기본시트명인  "sheet1"으로 해주세요.<br>
			 - 시트의 첫번째라인은 필드명(예:작성일자,승인번호),<br> 두번째부터 데이터내용(예:20111203,1234564889)이 들어가도록 폼을 변경해주세요 <br>
		</td>
	</tr>
</table>
</body>
</html>

