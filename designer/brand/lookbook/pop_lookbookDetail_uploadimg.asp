<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, mode, idx
	sFolder = requestCheckVar(Request.Querystring("sF"),100)
	sImg 	= requestCheckVar(Request.Querystring("sImg"),300)
	mode 	= requestCheckVar(Request.Querystring("mode"),50)
	idx 	= requestCheckVar(Request.Querystring("idx"),10)
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
	sName = requestCheckVar(Request.Querystring("sName"),300)
	sSpan = requestCheckVar(Request.Querystring("sSpan"),300)
%>

<script language="javascript">

	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return false;
		}
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle">이미지 업로드 처리</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadUrl %>/linkweb/street/doLookbookDetailupload_designer.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sFsub" value="detail">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
	<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
</tr>	
<%IF sImg <> "" THEN%>
<tr>
	<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName%></td>
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