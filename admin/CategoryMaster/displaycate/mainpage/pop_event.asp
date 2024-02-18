<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear, vEventID, vCateCode, vType, vPage, vStartDate
vType = Request.Querystring("type")
vEventID = Request.Querystring("eventid")
If vEventID = "0" Then
	vEventID = ""
End IF
vCateCode = Request.Querystring("catecode")
vPage = Request.Querystring("page")
vStartDate = Request.Querystring("startdate")


%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.eventid.value){
			alert("이벤트코드는 꼭 넣어주세요.");
			document.frmImg.eventid.focus();
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> <b><%=vType%></b> 이벤트등록</div>
<table width="360" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="pop_event_proc.asp" onSubmit="return jsUpload();">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="type" value="<%=vType%>">
<input type="hidden" name="page" value="<%=vPage%>">
<input type="hidden" name="startdate" value="<%=vStartDate%>">
	<tr>
		<td bgcolor="#FFFFFF0" colspan="2">
			* <b><font color="red">[필독]</font></b><b>등록 후</b> 해당 이벤트의 <b><font color="blue-green">제목, 이미지, 카피, 타입, 링크가 변경이 되면</font></b> 이 팝업창에서 <b><font color="blue">다시 확인 버튼을 눌러</font></b>주셔야 <b><font color="green">변경된 내용으로 적용</font></b>됩니다.
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="eventid" value="<%=vEventID%>">
		</td>
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