<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڳʰ���
' History : 2009.09.15 �ѿ�� ����
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
dim life_id,life_title,plusitem,commentyn
dim list_image,main_image1,main_image2,regdate,isusing
	life_id = requestcheckvar(request("life_id"),32)
	
'// �ִ°�쿡�� ����
dim oip
set oip = new clife_list
	oip.frectlife_id = life_id
	if life_id <> "" then
	oip.flife_edit()
	
		if oip.ftotalcount > 0 then
			life_id = oip.foneitem.flife_id 
			life_title = oip.foneitem.flife_title 
			list_image = oip.foneitem.flist_image 
			main_image1 = oip.foneitem.fmain_image1 
			main_image2 = oip.foneitem.fmain_image2 
			regdate = oip.foneitem.fregdate 
			isusing = oip.foneitem.fisusing 
			plusitem = oip.foneitem.fplusitem 
			commentyn = oip.foneitem.fcommentyn 
		end if
	end if

%>

<script language="javascript">
	
	document.domain = "10x10.co.kr";	
	
	//����
	function life_reg(){
		
		if(document.frmcontents.life_title.value==''){
			alert('������ �Է��ϼž� �մϴ�.');
			return false;
		}					
				
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
	<form name="frmcontents" method="post" action="<%=imgFingers%>/linkweb/corner/lifeimage_proc.asp" onsubmit="return false;" enctype="multipart/form-data">
			
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>ID</b><br></td>
		<td>
			<%= life_id %><input type="hidden" name="life_id" value="<%= life_id %>">
		</td>
	</tr>	
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>��������</b><br></td>
		<td>
			<input type="text" name="life_title" size="50" value="<%=life_title%>">
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>�ڸ�Ʈ<br>��뿩��</b><br></td>
		<td><select name="commentyn">
				<option value="Y" <% if commentyn = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if commentyn = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center"><b>�÷�������</b><br></td>
		<td>
			<font color="red">�����ǻ���)  , �� ���� �ǳ��� , ����</font><br>
			<textarea rows="5" cols="70" name="plusitem"><%= plusitem %></textarea>
		</td>
	</tr>			
	<tr bgcolor="#FFFFFF">
		<td align="center">
		<b>����Ʈ �̹���</b>
		<br><font color="red">180x55</font>
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
			<b>���� �̹���1</b>
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
			<b>���� �̹���2</b>
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
		<td align="center"><b>��뿩��</b><br></td>
		<td><select name="isusing">
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
		</td>
	</tr>		
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan="2">
			<% 
			'//����
			if life_id <> "" then 
			%>
				<input type="button" value="����" onclick="life_reg('');" class="button">
			<% 
			'//�ű�
			else 
			%>
				<input type="button" value="�ű�����" onclick="life_reg('');" class="button">
			<% end if %>
		</tr>
</form>	
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

