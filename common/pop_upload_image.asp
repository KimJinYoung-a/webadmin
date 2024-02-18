<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 이미지 등록
' History : 2010.06.16 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

''참조 : /admin2009scm/admin/lib/popoffshopinfo.asp
''function popUploadShopimage(frm)


dim mode, imagekind, pk, img50x50

mode 		= request("mode")
imagekind 	= requestCheckVar(request("imagekind"),32)
pk 			= requestCheckVar(request("pk"),32)
img50x50 	= request("50x50")

%>

<script language="javascript">

	document.domain ="10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.imagefile.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");
			return false;
		}
	}


</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이미지 업로드 처리</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/common/upload_image_process.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="imagekind" value="<%= imagekind %>">
<input type="hidden" name="pk" value="<%= pk %>">
<input type="hidden" name="50X50" value="<%= img50x50 %>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
	<td bgcolor="#FFFFFF"><input type="file" name="imagefile" class="file"></td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->