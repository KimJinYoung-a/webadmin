<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 히치하이커 어드민 프리뷰 Iframe_이미지 등록 페이지
' History : 2014.08.04 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, mode, idx, device
	sFolder = Request.Querystring("sF") 
	sImg 	= Request.Querystring("sImg")
	mode 	= Request.Querystring("mode")
	idx 	= Request.Querystring("idx")

	device 	= Request.Querystring("device")
IF sImg <> "" THEN
	arrImg = split(sImg,"/")
	slen = ubound(arrImg)
	sImgName = arrImg(slen)
END IF	
	sName = Request.Querystring("sName")
	sSpan = Request.Querystring("sSpan")
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
<form name="frmImg" method="post" action="<%= uploadUrl %>/linkweb/hitchhiker/dohitchhikerDetailupload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sFsub" value="detail">
<input type="hidden" name="sImg" value="<%=sImg%>">
<input type="hidden" name="sName" value="<%=sName%>">
<input type="hidden" name="sSpan" value="<%=sSpan%>">
<input type="hidden" name="device" value="<%=device%>">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
	<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
</tr>	
<% IF sImg <> "" THEN %>
<tr>
	<td colspan="2" bgcolor="#FFFFFF">현재 파일명 : <%=sImgName%></td>
</tr>	
<% END IF %>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>	
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->