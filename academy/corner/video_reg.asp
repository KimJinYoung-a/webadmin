<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/video_cls.asp"-->

<%
	Dim vIdx, vTitle, vCateCD2, vLecturer, vMakerID, vKeyword, vImage_URL, vImage2_URL, vYoutube_URL, vYoutube_source, vIsUsing, vRegdate

	vIdx = requestcheckvar(request("idx"),32)
	vIsUsing = "N"
	
'// �ִ°�쿡�� ����
dim oip
If vIdx <> "" Then
	set oip = new cvideo
		oip.frectidx = vIdx
		oip.video_view()
		
		if oip.ftotalcount > 0 then
			vIdx = oip.foneitem.fidx 
			vTitle = oip.foneitem.ftitle 
			vCateCD2 = oip.foneitem.fcatecd2 
			vLecturer = oip.foneitem.flecturer 
			vMakerID = oip.foneitem.fmakerid
			vKeyword = oip.foneitem.fkeyword 
			vImage_URL = oip.foneitem.fimage_url
			vImage2_URL = oip.foneitem.fimage2_url
			vYoutube_URL = oip.foneitem.fyoutube_url 
			vYoutube_source = oip.foneitem.fyoutube_source 
			vIsUsing = oip.foneitem.fisusing
			vRegdate = oip.foneitem.fregdate
		end if

	set oip = nothing
End IF
%>

<script language="javascript">

	document.domain = "10x10.co.kr";	
	
	//����
	function video_reg(){

		if(document.frmcontents.title.value==''){
			alert('������ �Է��ϼž� �մϴ�.');
			document.frmcontents.title.focus();
			return false;
		}
		if(document.frmcontents.CateCD2.value==''){
			alert('ī�װ��� �Է��ϼž� �մϴ�.');
			document.frmcontents.CateCD2.focus();
			return false;
		}
//		if(document.frmcontents.lecturer.value==''){
//			alert('������ ID�� �Է��ϼž� �մϴ�.');
//			document.frmcontents.lecturer.focus();
//			return false;
//		}
//		if(document.frmcontents.makerid.value==''){
//			alert('�귣��ID�� �Է��ϼž� �մϴ�.\n������ thefingers01�� �Է��ϼ���.');
//			document.frmcontents.makerid.focus();
//			return false;
//		}
		if(document.frmcontents.youtube_url.value==''){
			alert('YouTube URL�� �Է��ϼž� �մϴ�.');
			document.frmcontents.youtube_url.focus();
			return false;
		}
//		if(document.frmcontents.youtube_source.value==''){
//			alert('YouTube �ҽ��� �Է��ϼž� �մϴ�.');
//			document.frmcontents.youtube_source.focus();
//			return false;
//		}
		
		<% If vIdx = "" Then %>
		if(document.frmcontents.list_image.value==''){
			alert('����Ʈ �̹����� �����ϼž� �մϴ�.');
			return false;
		}
		<% End If %>

		frmcontents.submit();		
	}
	
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<form name="frmcontents" method="post" action="<%=UploadImgFingers%>/linkweb/corner/video_image_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>IDX</b><br></td>
		<td>
			<%= vIdx %><input type="hidden" name="idx" value="<%= vIdx %>">
			<% If vIdx <> "" Then %>&nbsp;&nbsp;&nbsp;�����:<%=vRegdate%><% End If %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>�� ��</b><br></td>
		<td>
			<input type="text" name="title" size="80" value="<%=vTitle%>" maxlength="150">
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>ī�װ�</b><br></td>
		<td>
			<% Call DrawSelectBoxLecCategoryLarge("CateCD2",vCateCD2,"N")%>
		</td>
	</tr>
<!--
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>���� ID</b><br></td>
		<td>
			<input type="text" name="lecturer" size="80" value="<%=vLecturer%>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>�귣��ID</b><br></td>
		<td>
			<input type="text" name="makerid" size="80" value="<%=vMakerID%>" maxlength="32">
		</td>
	</tr>
-->
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>����</b><br></td>
		<td>
			<input type="text" name="keyword" size="80" value="<%=vKeyword%>" maxlength="200">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>YouTube URL</b><br></td>
		<td>
			<input type="text" name="youtube_url" size="80" value="<%=vYoutube_URL%>" maxlength="200"><br>
			<font color="red"> �� ��Ʃ�� : �ҽ��ڵ� ���� (�� : http://www.youtube.com/embed/qj4rn1I_dC8 ) �� ��Ʃ�� ������ URL���� �ƴ�!</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>YouTube �ҽ�</b><p>width�� height��<br>width="705"<br>height="360"<br></td>
		<td>
			<textarea name="youtube_source" rows="12" cols="80"><%=vYoutube_source%></textarea>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>����Ʈ �̹���</b>
		<br><font color="red">240x160</font>
		</td>
		<td>
			<% if vImage_URL <> "" then %>
			<img src="<%=vImage_URL%>"><br>
			<% end if %>
			<input type="file" name="list_image" size="80" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>����Ʈ �̹���2</b>
		<br><font color="red">180x120</font>
		</td>
		<td>
			<% if vImage2_URL <> "" then %>
			<img src="<%=vImage2_URL%>"><br>
			<% end if %>
			<input type="file" name="list_image2" size="80" maxlength="32" class="file">
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>��뿩��</b><br></td>
		<td><select name="isusing">
				<option value="Y" <% if vIsUsing = "Y" then response.write " selected" end if %>>Y</option>
				<option value="N" <% if vIsUsing = "N" then response.write " selected" end if %>>N</option>
			</select>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//����
			if vIdx <> "" then 
			%>
				<input type="button" value="����" onclick="video_reg('');" class="button">
			<% 
			'//�ű�
			else 
			%>
				<input type="button" value="�ű�����" onclick="video_reg('');" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

