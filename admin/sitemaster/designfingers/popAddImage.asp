<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2008.03.13 create
' Description :  이미지 첨부
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
 Dim sFolder, sImgID, chkIcon,sImgURL, arrImg
 sFolder 	=  requestCheckVar(request("sF"),10)
 sImgID 	=  requestCheckVar(request("sID"),4)
 chkIcon   	=  requestCheckVar(request("chkI"),1)
 sImgURL 	= requestCheckVar(request("sIU"),100)
 
 '//이미지 명만 추출
 IF sImgURL <> "" THEN
 arrImg 	= split(sImgURL,"/")
 sImgURL	= arrImg(Ubound(arrImg))
END IF
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.sfImg.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");			
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이미지 첨부</div>
<table width="350" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/sitemaster/uploadDesignfingers.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="sF" value="<%=sFolder%>">
<input type="hidden" name="sID" value="<%=sImgID%>">
<input type="hidden" name="chkI" value ="<%=chkIcon%>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지명</td>
		<td bgcolor="#FFFFFF"><input type="file" name="sfImg"></td>
	</tr>	
	<%IF sImgURL <> "" THEN%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">현재 이미지명 : <%=sImgURL%></td>
	</tr>	
	<%END IF%>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">+ 최대 파일사이즈 1MB(1,024KB) 이하<br>
			+ gif,jpg,png 타입의 파일만 등록가능.
			</td>
	</tr>
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->