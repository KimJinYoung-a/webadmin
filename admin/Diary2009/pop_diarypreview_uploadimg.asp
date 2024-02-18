<%@ language=vbscript %>
<% 
	option explicit 
	Response.Expires = -1440
'###########################################################
' Description :  다이어리_프리뷰_이미지업로드
' History : 2018.08.16 이종화 생성 - 멀티 이미지 변경
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<%
Dim sFolder, sName, slen, arrImg, mode, idx
Dim iMaxLength : iMaxLength = 10
	sFolder = Request.Querystring("sF") 
	mode 	= Request.Querystring("mode")
	idx 	= Request.Querystring("idx")
	sName 	= Request.Querystring("sName")
%>
<script type="text/javascript">
<!--  
	function jsUpload(){
		var frm = document.frmImg;
	
		arrFileExt = new Array();
		arrFileExt[arrFileExt.length]  = "PNG";
		arrFileExt[arrFileExt.length]  = "GIF";
		arrFileExt[arrFileExt.length]  = "JPG";
		arrFileExt[arrFileExt.length]  = "JPEG";	
		//파일 입력확인
		var chkinput = 0; 
		
		for(i=0;i<4;i++){
			if( frm.sfImg[i].value !="") {
				chkinput = 1;
			}
		}
		 
		if(chkinput==0){
			alert("파일을 한개 이상 입력해주세요");
			frm.sfImg[0].focus();
			return;
		}	 
						
		//파일유효성 체크
		if (!fnChkFile(frm.sfImg[i].value, <%=iMaxLength%>, arrFileExt)){
			alert("이미지는 <%=iMaxLength%>MB이하의  지원되는 형식의 파일만 업로드 가능합니다.\n\n 지원되는 파일형식은 관리자에게 문의해주세요");
			return;
		}
		
		frm.submit(); 
		document.all.dvLoad.style.display = "";
	}
		
	function fnChkFile(sfImg, sMaxSize, arrExt){   
    	//파일 업로드 유무확인
   	 	if (!sfImg){
    		return true;
    	}
    	var blnResult = false;
        
		//파일 용량 확인
		var maxsize = sMaxSize * 1024 * 1024;
		
		//파일 확장자 확인
		var pPoint = sfImg.lastIndexOf('.');
		var fPoint = sfImg.substring(pPoint+1,sfImg.length);
		var fExet = fPoint.toLowerCase();

		for (var i = 0; i < arrExt.length; i++){
			if (arrExt[i].toLowerCase() == fExet) 
			{ 
				blnResult =  true;
			}
		}
		
		return blnResult; 
   }
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle">이미지 업로드 처리</div><br/>
<form name="frmImg" method="post" action="<%= uploadUrl %>/linkweb/diary/diarydetailimagesUpload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sFsub" value="detail">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<tr bgcolor="#FFFFFF">
	<td valign="top">	 
		<table width="100%" border="0" cellpadding="0" cellspacing="5" class="a">  
			<tr>
				<td valign="top">파일명:</td>
				<td><input type="file" name="sfImg" ><br>
					<input type="file" name="sfImg" ><br>
					<input type="file" name="sfImg" ><br>
					<input type="file" name="sfImg" ><br>
					<input type="file" name="sfImg" > 
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>	
		<div style="font-size:11px;">- 지원되는 파일포맷: JPG,JPEG,GIF,PNG</div>
		<div style="padding-top:5px;">- 최대 <font color="red">10,240KB</font>까지 등록가능합니다.</div>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</table>
</form>	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<div id="dvLoad" style="display:none;top:100px;left:50;position:absolute;background-color:gray;">
	<table border="0" class="a" cellpadding="5" cellspacing="5">
		<tr>
			<td> <font color="#FFFFFF">업로드 처리중입니다. 잠시만 기다려주세요~~</font></td>
		</tr>
	</table>
</div>