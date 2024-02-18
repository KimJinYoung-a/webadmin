<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim vQuery, i, vYear, vMonth, vItemID, vCateCode, vType, vIdx, vPage, vStartDate, vImg(2), vLink(2), vTitle, vSubcopy
vStartDate = Request.Querystring("startdate")
vType = Request.Querystring("type")
vItemID = Request.Querystring("itemid")
vCateCode = Request.Querystring("catecode")
vPage = Request.Querystring("page")
vYear = Year(now)
vMonth = Month(now)
if len(vMonth) = 1 then vMonth = "0"&vMonth end if


vQuery = "select imgurl, linkurl, title, subcopy from [db_sitemaster].[dbo].[tbl_display_catemain_detail] "
vQuery = vQuery & "where startdate = '" & vStartDate & "' and catecode = '" & vCateCode & "' and page = '" & vPage & "' "
If vType = "multi" Then
vQuery = vQuery & " and type in ('multiimg1','multiimg2','multiimg3') "
Else
vQuery = vQuery & " and type = '" & vType & "' "
End If
vQuery = vQuery & "order by idx asc"
rsget.Open vQuery, dbget, 1
If Not rsget.Eof Then
	i = 0
	Do Until rsget.Eof
		vImg(i) = rsget("imgurl")
		vLink(i) = rsget("linkurl")
		vTitle = db2html(rsget("title"))
		vSubcopy = db2html(rsget("subcopy"))
		i = i + 1
	rsget.Movenext
	Loop
End If
rsget.close()
%>
<script language="javascript">
<!--
document.domain = "10x10.co.kr";

	function jsUpload(){
		if(!document.frmImg.multilink1.value){
			alert("이미지1링크를 넣어주세요.");
			document.frmImg.multilink1.focus();
			return false;
		}
		<% If vType = "multi" Then %>
		if(!document.frmImg.multilink2.value){
			alert("이미지2링크를 넣어주세요.");
			document.frmImg.multilink2.focus();
			return false;
		}
		if(!document.frmImg.multilink3.value){
			alert("이미지3링크를 넣어주세요.");
			document.frmImg.multilink3.focus();
			return false;
		}
		<% End If %>
	}
	
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> <b><%=vType%></b> 이미지 등록</div>
<table width="560" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/category/img_upload.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<input type="hidden" name="startdate" value="<%=vStartDate%>">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="type" value="<%=vType%>">
<input type="hidden" name="page" value="<%=vPage%>">
<input type="hidden" name="yr" value="<%=vYear%>">
<input type="hidden" name="mm" value="<%=vMonth%>">
<input type="hidden" name="reguserid" value="<%=session("ssBctId")%>">
<input type="hidden" name="regusername" value="<%=session("ssBctCname")%>">
	<tr>
		<td bgcolor="#FFFFFF" colspan="2">
			* 이미지 총 <b> <%=CHKIIF(vType="multi","3","1")%>개 필수</b>. 사이즈(<%=CHKIIF(vType="multi","444x444","444x212")%>).<br>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지1</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="file1_now" value="<%=vImg(0)%>">
			<table border="0" class="a" width="100%">
			<tr>
				<td><input type="file" name="file1"></td>
				<% If vImg(0) <> "" Then %><td align="right"><a href="<%=vImg(0)%>" target="_blank"><img src="<%=vImg(0)%>" width="40" border="0"></a></td><% End If %>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지1링크</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="multilink1" value="<%=vLink(0)%>" size="62">
		</td>
	</tr>
	<% If vType = "recipe" Then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">타이틀</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="title" value="<%=vTitle%>" size="62">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">서브카피</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="subcopy" value="<%=vSubcopy%>" size="62">
		</td>
	</tr>
	<% End If %>
	<% If vType = "multi" Then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지2</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="file2_now" value="<%=vImg(1)%>">
			<table border="0" class="a" width="100%">
			<tr>
				<td><input type="file" name="file2"></td>
				<% If vImg(1) <> "" Then %><td align="right"><a href="<%=vImg(1)%>" target="_blank"><img src="<%=vImg(1)%>" width="40" border="0"></a></td><% End If %>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지2링크</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="multilink2" value="<%=vLink(1)%>" size="62">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지3</td>
		<td bgcolor="#FFFFFF">
			<input type="hidden" name="file3_now" value="<%=vImg(2)%>">
			<table border="0" class="a" width="100%">
			<tr>
				<td><input type="file" name="file3"></td>
				<% If vImg(2) <> "" Then %><td align="right"><a href="<%=vImg(2)%>" target="_blank"><img src="<%=vImg(2)%>" width="40" border="0"></a></td><% End If %>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">이미지3링크</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="multilink3" value="<%=vLink(2)%>" size="62">
		</td>
	</tr>
	<% End If %>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<input type="image" src="/images/icon_confirm.gif">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>	
</form>	
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->