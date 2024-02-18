<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 코너관리
' History : 2009.09.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->

<%
dim idx,lecturer_id,image_400x400,image_50x50,image_80x80,regdate,isusing
	lecturer_id = requestcheckvar(request("lecturer_id"),32)
	idx = requestcheckvar(request("idx"),8)
	
'// 있는경우에만 쿼리
dim oip
set oip = new cgood_onelist
oip.frectidx = idx
oip.frectlecturer_id = 	lecturer_id
	if lecturer_id <> "" then
	oip.fgood_item_edit()
	
		if oip.ftotalcount > 0 and idx <> "" then
			image_400x400 = oip.foneitem.fimage_400x400 
			image_50x50 = oip.foneitem.fimage_50x50
			image_80x80 = oip.foneitem.fimage_80x80
			regdate = oip.foneitem.fregdate 
			isusing = oip.foneitem.fisusing 																	
		end if
	end if

%>

<script language="javascript">
	
	//저장
	function good_reg(mode){			
		frmcontents.submit();		
	}
	
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<form name="frmcontents" method="post" action="<%=imgFingers%>/linkweb/corner/lectureritem_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="idx" value="<%=idx%>">	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>강사ID</b><br></td>
		<td>
			<%=lecturer_id%>
			<input type="hidden" name="lecturer_id" value="<%=lecturer_id%>">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>기본 이미지</b>
		<br><font color="red">400x400</font>
		</td>
		<td>
			<% if image_400x400 <> "" then %>
			<img src="<%=image_400x400%>"><br>
			<% end if %>
			<input type="file" name="image_400x400" size="32" maxlength="32" class="file">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>자동생성 이미지</b>
		<br><font color="red">image_50x50</font>
		</td>
		<td>
			<% if image_50x50 <> "" then %>
			<img src="<%=image_50x50%>"><br>			
			<% end if %>			
			※400x400 이미지 등록시 자동으로 50x50 이미지가 생성됩니다.
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>자동생성 이미지</b>
		<br><font color="red">image_80x80</font>
		</td>
		<td>
			<% if image_80x80 <> "" then %>
			<img src="<%=image_80x80%>"><br>			
			<% end if %>			
			※400x400 이미지 등록시 자동으로 80x80 이미지가 생성됩니다.
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>사용여부</b><br></td>
		<td><select name="isusing">
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//수정
			if idx <> "" then 
			%>
				<input type="button" value="수정" onclick="good_reg();" class="button">
			<% 
			'//신규
			else 
			%>
				<input type="button" value="신규저장" onclick="good_reg();" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

