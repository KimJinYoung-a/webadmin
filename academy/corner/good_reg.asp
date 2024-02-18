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
dim lecturer_id, lecturer_name ,history ,history_act, catecd2, socname , image_profile_75x75, newImage_profile, newProfileFolder
dim socname_kor ,image_profile, image_top, regdate, isusing ,homepage , image_list , best, twitter, onesentence
	lecturer_id = requestcheckvar(request("lecturer_id"),32)
	
'// 있는경우에만 쿼리
dim oip
set oip = new cgood_onelist
oip.frectlecturer_id = 	lecturer_id
	if lecturer_id <> "" then
	oip.fgood_edit()
	
		if oip.ftotalcount > 0 then
			lecturer_name = oip.foneitem.flecturer_name 
			history = oip.foneitem.fhistory
			history_act = oip.foneitem.fhistory_act												
			catecd2 = oip.foneitem.fcatecd2 
			socname = oip.foneitem.fsocname 
			socname_kor = oip.foneitem.fsocname_kor 
			image_profile = oip.foneitem.fimage_profile 
			image_profile_75x75 = oip.foneitem.fimage_profile_75x75 
			image_top = oip.foneitem.fimage_top 
			image_list = oip.foneitem.fimage_list
			regdate = oip.foneitem.fregdate 
			isusing = oip.foneitem.fisusing 
			homepage = oip.foneitem.fhomepage 	
			best = oip.foneitem.fbest 
			twitter = oip.foneitem.ftwitter
		end if
	end if

'// 핑거스이미지서버에서 강사id변경후 리로드(도메인문제) 안하고 팝업이 
'// 또 뜰경우 강제로 부모창(리로드)과 현재페이지(신규등록) 상태로 바꾼다
if lecturer_id <> "" and lecturer_name = "" then
	response.write "<script>"
	response.write "opener.location.reload();"
	response.write "location.href='/academy/corner/good_reg.asp';"
	response.write "</script>"
end if
%>

<script language="javascript">

	function FnLecturerApp(str){
		var varArray;
		varArray = str.split(',');
	
		document.frmcontents.lecturer_id.value = varArray[0];
		document.frmcontents.lecturer_name.value = varArray[1];
		document.frmcontents.socname.value = varArray[2];
		document.frmcontents.socname_kor.value = varArray[3];
			
	}

	document.domain = "10x10.co.kr";	
	
	//저장
	function good_reg(mode){
		
		if(document.frmcontents.temp_lec_id.value==''){
			alert('강사 구분을 입력하셔야 합니다.');
			return false;
		}
		if(document.frmcontents.CateCD2.value==''){
			alert('카테고리를 입력하셔야 합니다.');
			return false;
		}
		if(document.frmcontents.lecturer_name.value==''){
			alert('강사명을 입력하셔야 합니다.');
			return false;
		}				
		if (mode == 'add'){
			frmcontents.mode.value='add';
		}else if(mode == 'edit'){
			frmcontents.mode.value='edit';
		}
		
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
	<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/corner/lecturerimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="mode" >
	<input type="hidden" name="tmplecturer_id" value="<%=lecturer_id%>">	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>강사</b><br></td>
		<td>
			<% SelectLecturerId(lecturer_id) %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>강사ID</b><br></td>
		<td>
			<input type="text" name="lecturer_id" value="<%=lecturer_id%>">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>강사명</b><br></td>
		<td>
			<input type="text" name="lecturer_name" size="50" value="<%=lecturer_name%>">
		</td>
	</tr>	
		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>브랜드(영문)</b><br></td>
		<td>
			<input type="text" name="socname" size="50" value="<%=socname%>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>브랜드(한글)</b><br></td>
		<td>
			<input type="text" name="socname_kor" size="50" value="<%=socname_kor%>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>카테고리</b><br></td>
		<td>
			<%=makeCateSelectBox("CateCD2",catecd2)%>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>소개글</b><br></td>
		<td>
			<textarea name="history" cols="60" rows="6"><%=history%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>작가활동</b><br></td>
		<td>
			<textarea name="history_act" cols="60" rows="6"><%=history_act%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>BEST</b><br></td>
		<td>
			<select name="best">
				<option value="Y" <% if best = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if best = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>ICON 이미지</b>
		<br><font color="red">120x120</font>
		</td>
		<td>
			<% if image_profile <> "" then %>
			<img src="<%=image_profile%>"><br>
			<% end if %>
			<% if image_profile_75x75 <> "" then %>
			<img src="<%=image_profile_75x75%>"><br>
			<% end if %>
			※120x120 등록시 75x75 자동생성됩니다.<br>
			<input type="file" name="image_profile" size="32" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<b>TOP 이미지</b>
			<br><font color="red">210x146</font>
		</td>
		<td>
			<% if image_top <> "" then %>
			<img src="<%=image_top%>"><br>
			<% end if %>
			<input type="file" name="image_top" size="32" maxlength="32" class="file">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<b>list 이미지</b>
			<br><font color="red">180x120</font>			
		</td>
		<td>
			<% if image_list <> "" then %>
			<img src="<%=image_list%>"><br>
			<% end if %>
			<input type="file" name="image_list" size="32" maxlength="32" class="file">
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
		<td align="center"><b>트위터</b><br></td>
		<td>
			<input type="text" name="twitter" size="50" value="<%=twitter%>"><br>
			<font color="red">※ex) http://twitter.com/Yunaaaa</font>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>홈페이지</b><br></td>
		<td>
			<input type="text" name="homepage" size="50" value="<%=homepage%>"><br>
			<font color="red">※ex) http://www.10x10.co.kr</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//수정			
			if lecturer_id <> "" and lecturer_name <> "" then 
			%>
				<input type="button" value="수정" onclick="good_reg('edit');" class="button">
			<% 
			'//신규
			else 
			%>
				<input type="button" value="신규저장" onclick="good_reg('add');" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

