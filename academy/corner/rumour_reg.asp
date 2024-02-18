<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 코너관리
' History : 2009.09.14 한용민 생성
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
dim rumour_id,rumour_title,rumour_userid,startdate,enddate
dim list_image,main_image1,main_image2,regdate,isusing
	rumour_id = requestcheckvar(request("rumour_id"),32)
	
'// 있는경우에만 쿼리
dim oip
set oip = new crumour_one_list
	oip.frectrumour_id = rumour_id
	if rumour_id <> "" then
	oip.frumour_edit()
	
		if oip.ftotalcount > 0 then
			rumour_id = oip.foneitem.frumour_id 
			rumour_title = oip.foneitem.frumour_title 
			rumour_userid = oip.foneitem.frumour_userid 
			startdate = oip.foneitem.fstartdate 
			enddate = oip.foneitem.fenddate
			list_image = oip.foneitem.flist_image 
			main_image1 = oip.foneitem.fmain_image1 
			main_image2 = oip.foneitem.fmain_image2 
			regdate = oip.foneitem.fregdate 
			isusing = oip.foneitem.fisusing 
		end if
	end if

%>

<script language="javascript">

	document.domain = "10x10.co.kr";	
	
	//저장
	function rumour_reg(){
		
		if(document.frmcontents.rumour_title.value==''){
			alert('제목을 입력하셔야 합니다.');
			return false;
		}
		if(document.frmcontents.rumour_userid.value==''){
			alert('전시자를 입력하셔야 합니다.');
			return false;
		}
		if(document.frmcontents.startdate.value==''){
			alert('시작일을 입력하셔야 합니다.');
			return false;
		}
		if(document.frmcontents.enddate.value==''){
			alert('종료일을 입력하셔야 합니다.');
			return false;
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
	<form name="frmcontents" method="post" action="<%=imgFingers%>/linkweb/corner/rumourimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>ID</b><br></td>
		<td>
			<%= rumour_id %><input type="hidden" name="rumour_id" value="<%= rumour_id %>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>전시제목</b><br></td>
		<td>
			<input type="text" name="rumour_title" size="50" value="<%=rumour_title%>">
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>전시자</b><br></td>
		<td>
			<input type="text" name="rumour_userid" size="50" value="<%=rumour_userid%>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>전시 기간</b><br></td>
		<td>
			<input type="text" name="startdate" size=10 value="<%= startdate %>">			
			<a href="javascript:calendarOpen3(frmcontents.startdate,'시작일',frmcontents.startdate.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
			<input type="text" name="enddate" size=10  value="<%= left(enddate,10) %>">
			<a href="javascript:calendarOpen3(frmcontents.enddate,'마지막일',frmcontents.enddate.value)">
			<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>	
		</td>
	</tr>			

	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>리스트 이미지</b>
		<br><font color="red">55x55</font>
		</td>
		<td>
			<% if list_image <> "" then %>
			<img src="<%=list_image%>"><br>
			<% end if %>
			<input type="file" name="list_image" size="32" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<b>메인 이미지1</b>
			<br><font color="red">760</font>			
		</td>
		<td>
			<% if main_image1 <> "" then %>
			<img src="<%=main_image1%>"><br>
			<% end if %>
			<input type="file" name="main_image1" size="32" maxlength="32" class="file">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<b>메인 이미지2</b>
			<br><font color="red">760</font>			
		</td>
		<td>
			<% if main_image2 <> "" then %>
			<img src="<%=main_image2%>"><br>
			<% end if %>
			<input type="file" name="main_image2" size="32" maxlength="32" class="file">
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
			if rumour_id <> "" then 
			%>
				<input type="button" value="수정" onclick="rumour_reg('');" class="button">
			<% 
			'//신규
			else 
			%>
				<input type="button" value="신규저장" onclick="rumour_reg('');" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

