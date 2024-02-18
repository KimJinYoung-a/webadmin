<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.14 한용민 생성
'	Description : artist gallery
'#######################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->

<%
'// 변수 선언
dim mode, gal_sn, lp
dim page, isusing, gal_div, designerid
	mode = request("mode")
	gal_sn = request("gal_sn")
	
	page = request("page")
	isusing = request("isusing")
	gal_div = request("gal_div")
	designerid = request("designerid")
%>

<script language="javascript">

	function subcheck(){
		var frm=document.inputfrm;
	
		if (!frm.gal_div.value) {
			alert('이미지 구분을 선택해 주세요..');
			frm.gal_div.focus();
			return;
		}
	
		if (frm.designerid.value.length< 1 ){
			 alert('업체를 선택 해주세요');
		frm.designerid.focus();
		return;
		}
		if (!frm.gal_sortNo.value){
			 alert('표시순서를 입력해주세요');
		frm.gal_sortNo.focus();
		return;
		}
	
		frm.submit();
	}
	
	function popSearchBrand()
	{
		window.open("popBrandSearch.asp","popBrand","width=338,height=350,scrollbars=yes");
	}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" value=" 저장 " onclick="subcheck();" class="button"> 
			<input type="button" value=" 취소 " onclick="history.back();" class="button">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/doArtistGallery.asp" enctype="multipart/form-data">
	<input type="hidden" name="mode" value="<% =mode %>">
	<input type="hidden" name="gal_sn" value="<%= gal_sn %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="orgUsing" value="<%= isusing %>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">

	<% if mode="add" then %>
	<tr>
		<td width="100" bgcolor="#F0F0FD" align="center">이미지 구분</td>
		<td bgcolor="#FFFFFF">
			<select name="gal_div">
				<option value=""<% if gal_div="" then Response.Write " selected" %>>선택</option>
				<option value="W"<% if gal_div="W" then Response.Write " selected" %>>Work</option>
				<option value="D"<% if gal_div="D" then Response.Write " selected" %>>Drawing</option>
				<option value="P"<% if gal_div="P" then Response.Write " selected" %>>Photo</option>
			</select>
		</td>
	</tr>
	
	<tr>
		<td align="center" bgcolor="#F0F0FD">브랜드</td>
		<td bgcolor="#FFFFFF">
			<% Call DrawSelectBoxUseBrand("designerid",designerid) %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">갤러리 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="gal_imgorg" value="" size="55">
			<br>(1MB이하의 JPG 혹은 GIF형식의 가급적 정사각형 이미지로 업로드해주세요.)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">이미지 설명</td>
		<td bgcolor="#FFFFFF">
			<textarea class="textarea" name="gal_desc" cols="60" rows="3"></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">표시순서</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="gal_sortNo" value="0" size="3">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">사용유무</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y" checked>Y
			<input type="radio" name="isusing" value="N">N
		</td>
	</tr>

	<% elseif mode="edit" then

		'// 목록 접수
		dim oGallery
		set oGallery = New CGallery
		oGallery.FRectgal_sn = gal_sn
		oGallery.GetGalleryInfo
	%>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0FD">번호</td>
		<td bgcolor="#FFFFFF"><%=gal_sn%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#F0F0FD" align="center">이미지 구분</td>
		<td bgcolor="#FFFFFF">
			<select name="gal_div">
				<option value=""<% if oGallery.FItemList(1).Fgal_div="" then Response.Write " selected" %>>선택</option>
				<option value="W"<% if oGallery.FItemList(1).Fgal_div="W" then Response.Write " selected" %>>Work</option>
				<option value="D"<% if oGallery.FItemList(1).Fgal_div="D" then Response.Write " selected" %>>Drawing</option>
				<option value="P"<% if oGallery.FItemList(1).Fgal_div="P" then Response.Write " selected" %>>Photo</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">브랜드</td>
		<td bgcolor="#FFFFFF">
			<%=oGallery.FItemList(1).Fsocname & " (" & oGallery.FItemList(1).Fsocname_kor & ")"%>
			<input type="hidden" name="designerid" value="<%=oGallery.FItemList(1).Fdesignerid%>">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">갤러리 이미지</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="gal_imgorg" value="" size="55">
			<br>(1MB이하의 JPG 혹은 GIF형식의 가급적 정사각형 이미지로 업로드해주세요.)
			<% if oGallery.FItemList(1).Fgal_img400<>"" then %>
			<br><img src="<%=oGallery.FItemList(1).Fgal_img400%>" border="0">
			<br>Filename : <%=oGallery.FItemList(1).Fgal_imgorg%>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">이미지 설명</td>
		<td bgcolor="#FFFFFF">
			<textarea class="textarea" name="gal_desc" cols="60" rows="3"><%=oGallery.FItemList(1).Fgal_desc%></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">표시순서</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="gal_sortNo" value="<%=oGallery.FItemList(1).Fgal_sortNo%>" size="3">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">사용유무</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y"<% if oGallery.FItemList(1).Fgal_isusing="Y" then Response.Write " checked" %>>Y
			<input type="radio" name="isusing" value="N"<% if oGallery.FItemList(1).Fgal_isusing="N" then Response.Write " checked" %>>N
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
				<input type="button" value=" 저장 " onclick="subcheck();"> &nbsp;&nbsp;
				<input type="button" value=" 취소 " onclick="history.back();">
		</td>
	</tr>
	
	<% end if %>
	
	</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
