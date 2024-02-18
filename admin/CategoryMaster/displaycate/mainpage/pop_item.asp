<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim sFolder, sImg, sName,sSpan, slen, arrImg, sImgName, vYear, vItemID, vCateCode, vType, vPage, vStartDate
vType = Request.Querystring("type")
vItemID = Request.Querystring("itemid")
If vItemID = "0" Then
	vItemID = ""
End IF
vCateCode = Request.Querystring("catecode")
vPage = Request.Querystring("page")
vStartDate = Request.Querystring("startdate")

%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.itemid.value){
			alert("상품코드는 꼭 넣어주세요.");
			document.frmImg.itemid.focus();
			return false;
		}
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> <b><%=vType%></b> 상품등록</div>
<table width="360" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="pop_item_proc.asp" onSubmit="return jsUpload();">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="type" value="<%=vType%>">
<input type="hidden" name="page" value="<%=vPage%>">
<input type="hidden" name="startdate" value="<%=vStartDate%>">
	<tr>
		<td bgcolor="#FFFFFF" colspan="2">
			* <b>상품코드</b>는 <b>필수</b><br>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemid" value="<%=vItemID%>">
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