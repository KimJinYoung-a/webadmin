<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : showbanner_upload
' Hieditor : 2014.03.13 이종화 생성
'			 2017.03.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
Dim sImg, sName,sSpan, slen, arrImg, sImgName	
	sImg = Request.Querystring("sImg")

IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	

sName = Request.Querystring("sName")
sSpan = Request.Querystring("sSpan")
%>

<script type="text/javascript">

	document.domain = '10x10.co.kr';

	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return false;
		}
	}
	

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이미지 업로드 처리</div>
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/showbanner/showbanner_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();" style="margin:0px;" >
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
	<td bgcolor="#FFFFFF"><input type="file" name="sfImg" class="file"></td>
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
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->