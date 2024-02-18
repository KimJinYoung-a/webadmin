<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 아티스트 브랜드 관리 페이지   
' History : 2012.03.27 김진영 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->

<%
'// 변수 선언
dim oGallery
dim mode, idx, lp
dim page, isusing, gal_div, designerid
mode = request("mode")
If mode = "" Then mode = "add"

idx = request("idx")
page = request("page")
isusing = request("isusing")
gal_div = request("gal_div")
designerid = request("designerid")
%>
<script language="javascript">
function subcheck(){
	var frm=document.inputfrm;

	if (frm.designerid.value.length< 1 ){
		alert('업체를 선택 해주세요');
		frm.designerid.focus();
		return;
	}
	frm.submit();
}

function popSearchBrand()
{
	window.open("popBrandSearch.asp","popBrand","width=338,height=350,scrollbars=yes");
}
</script>
<p>
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/doArtist_Brand.asp" enctype="multipart/form-data">
	<input type="hidden" name="mode" value="<% =mode %>">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="orgUsing" value="<%= isusing %>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">

	<% if mode="add" then %>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">브랜드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="designerid" size="15" readonly value="<%=designerid%>">
			<input type="button" value="검색" onClick="popSearchBrand()">
			<span name="designerName" id="designerName"></span>
		</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">배너 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="file2" value="" size="55">
		</td>
	</tr>
	<tr>
		<td height="30" align="center" bgcolor="#F0F0FD">사용유무</td>
		<td bgcolor="#FFFFFF">
			<label><input type="radio" name="isusing" value="Y" checked>Y</label>
			<label><input type="radio" name="isusing" value="N">N</label>
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
				<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
				<input type="button" value=" 취소 " onclick="history.back();">
		</td>
	</tr>
	<% elseif mode="edit" then
		'// 목록 접수
		set oGallery = New cposcode_list
		oGallery.FRectIdx = idx
		oGallery.FArtistBrand_oneitem
	%>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0FD" height="30">번호</td>
		<td bgcolor="#FFFFFF"><%=idx%></td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD" height="30">브랜드</td>
		<td bgcolor="#FFFFFF">
			<%=oGallery.FOneItem.fsocname & " (" & oGallery.FOneItem.Fsocname_kor & ")"%>
			<input type="hidden" name="designerid" value="<%=oGallery.FOneItem.Fdesignerid%>">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">배너 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="file2" value="" size="55">
			<% if oGallery.FOneItem.ffile2<>"" then %>
			<br><img src="<%=uploadUrl%>/artist/brandbanner/<%=oGallery.FOneItem.ffile2%>">
			<br>Filename : <%=oGallery.FOneItem.ffile2%>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">사용유무</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y"<% if oGallery.FOneItem.Fisusing="Y" then Response.Write " checked" %>>Y
			<input type="radio" name="isusing" value="N"<% if oGallery.FOneItem.Fisusing="N" then Response.Write " checked" %>>N
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
				<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
				<input type="button" value=" 취소 " onclick="history.back();">
		</td>
	</tr>
	<% end if %>
	<%set oGallery = nothing %>
	</form>
</table>
<!-- 액션 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->