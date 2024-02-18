<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 카드청구내역 파일  등록
' History : 2012.05.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<% 
Dim intY, intM, sYear, sMonth,iOpExpPartIdx,arrPart
Dim clsPart,ipartsn,sadminid,iPartTypeIdx,blnAdmin
Dim  iMaxLength	
	IF iMaxLength = "" THEN iMaxLength = 10 
		
		sYear = Year(date())
		sMonth = month(date())
	 	iOpExpPartIdx = requestCheckvar(Request("selP"),10) 
	 	iPartTypeIdx	= 4
	 	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '어드민권한	
	 	
	 	IF not blnAdmin THEN  '리스트 권한을 가진 사람을 제외하고 담당자와 담당부서  view 가능
			ipartsn  =  session("ssAdminPsn")
	 		sadminid = 	session("ssBctId")
	 	END IF
	 	
	 	Set clsPart = new COpExpPart
	 	clsPart.FRectPartsn = ipartsn
		clsPart.FRectUserid = sadminid
	 	clsPart.FPartTypeidx 	= iPartTypeIdx   
		arrPart = clsPart.fnGetOpExppartAllList   
	 	Set clsPart = nothing
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
   
   function jsChDisp(iType){
   	if (iType==1){
   		document.all.disDate.style.display="";
   	}else{
   		document.all.disDate.style.display="none";
   	}
  }
	//-->
	</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td><strong>법인카드 파일 등록</strong><br><hr width="100%"></td>
</tr> 
<tr>
	<td> 
		<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" > 
			<form name="frmFile" method="post" action="<%=uploadImgUrl%>/linkweb/expense/procOpExp.asp"  enctype="MULTIPART/FORM-DATA">
			<input type="hidden" name="iML" value="<%=iMaxLength%>">
			<input type="hidden" name="sRID" value="<%=session("ssBctId")%>">
				<tr>
					<td width="100" style="padding:3px;" align="center" bgcolor="<%= adminColor("tabletop") %>"> 	청구일 입력여부 </td>
					<td style="padding:3px;" bgcolor="#FFFFFF"> <input type="radio" name="rdoD" value="1" onClick="jsChDisp(1);"> 입력 <input type="radio" name="rdoD" value="0" checked onClick="jsChDisp(0);"> 미입력</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#FFFFFF">
						<div style="display:none;" id="disDate">
						<table border="0" class="a" cellpadding="3" cellspacing="0">
						<tr>	
							<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>" style="border-right:1px <%= adminColor("tablebg") %> solid;"> 청구일 </td>
							<td bgcolor="#FFFFFF">
								<select name="selY">
								<%For intY=2012 To Year(date()) STEP -1%>	
								<option value="<%=intY%>" <%IF intY = sYear THEN%>selected<%END IF%>><%=intY%></option>
								<%Next%>
							</select>년	
							<select name="selM">
								<%For intM = 1 To 12%>
								<option value="<%=intM%>" <%IF intM = sMonth THEN%>selected<%END IF%>><%=intM%></option>	
								<%Next%>
							</select>월 
							</td>
						</tr>
					</table>
					</div>
					</td>
				</tr>				
				<tr>
					<td style="padding:3px;" align="center" bgcolor="<%= adminColor("tabletop") %>">파일명 </td>
					<td style="padding:3px;" bgcolor="#FFFFFF"><input type="file" name="sFile" class="button"></td>
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
			 - 엑셀(.XLS) 파일만 등록가능합니다.&nbsp;<font color="red">엑셀통합문서 (.XLSX)은 등록불가</font><br>
			 - Excel 97-2003 통합문서로 파일을 생성해주세요<br>	
			 - 시트명은 기본시트명인  "sheet1"으로 해주세요.<br>
			 - 시트의 첫번째라인은 필드명(예:작성일자,승인번호),<br> 두번째부터 데이터내용(예:20111203,1234564889)이 들어가도록 폼을 변경해주세요 <br>
		</td>
	</tr>
</table>	
</body>
</html>
			
			