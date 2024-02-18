<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'			2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/artistGalleryCls.asp" -->

<%
'// ���� ����
dim mode, gal_sn, lp
dim page, isusing, gal_div, designerid
	mode = requestCheckVar(request("mode"),32)
	gal_sn = requestCheckVar(request("gal_sn"),10)
	
	page = requestCheckVar(request("page"),10)
	isusing = requestCheckVar(request("isusing"),1)
	gal_div = requestCheckVar(request("gal_div"),10)
	designerid = requestCheckVar(request("designerid"),32)

designerid = session("ssBctID")	
%>

<script type='text/javascript'>

	function subcheck(){
		var frm=document.inputfrm;
	
		if (!frm.gal_div.value) {
			alert('�̹��� ������ ������ �ּ���..');
			frm.gal_div.focus();
			return;
		}
	
		if (frm.designerid.value.length< 1 ){
			 alert('��ü�� ���� ���ּ���');
		frm.designerid.focus();
		return;
		}
		if (!frm.gal_sortNo.value){
			 alert('ǥ�ü����� �Է����ּ���');
		frm.gal_sortNo.focus();
		return;
		}
	
		frm.submit();
	}

</script>

<!-- #include virtual="/designer/brand/inc_streetHead.asp"-->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" value=" ���� " onclick="subcheck();" class="button"> 
			<input type="button" value=" ��� " onclick="history.back();" class="button">
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="inputfrm" method="post" action="<%=uploadUrl%>/linkweb/street/doArtistGallery_designer.asp" enctype="multipart/form-data">
	<input type="hidden" name="mode" value="<% =mode %>">
	<input type="hidden" name="gal_sn" value="<%= gal_sn %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="orgUsing" value="<%= isusing %>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">

	<% if mode="add" then %>
	<tr>
		<td width="100" bgcolor="#F0F0FD" align="center">�̹��� ����</td>
		<td bgcolor="#FFFFFF">
			<select name="gal_div">
				<option value=""<% if gal_div="" then Response.Write " selected" %>>����</option>
				<option value="W"<% if gal_div="W" then Response.Write " selected" %>>Work</option>
				<option value="D"<% if gal_div="D" then Response.Write " selected" %>>Drawing</option>
				<option value="P"<% if gal_div="P" then Response.Write " selected" %>>Photo</option>
			</select>
		</td>
	</tr>
	
	<tr>
		<td align="center" bgcolor="#F0F0FD">�귣��</td>
		<td bgcolor="#FFFFFF">
			<%= designerid %>
			<input type="hidden" name="designerid" value="<%= designerid %>">		
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">������ �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="gal_imgorg" value="" size="55">
			<br>(1MB������ JPG Ȥ�� GIF������ ������ ���簢�� �̹����� ���ε����ּ���.)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�̹��� ����</td>
		<td bgcolor="#FFFFFF">
			<textarea class="textarea" name="gal_desc" cols="60" rows="3"></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">ǥ�ü���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="gal_sortNo" value="0" size="3">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�������</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y" checked>Y
			<input type="radio" name="isusing" value="N">N
		</td>
	</tr>

	<% elseif mode="edit" then

		'// ��� ����
		dim oGallery
		set oGallery = New CGallery
		oGallery.FRectgal_sn = gal_sn
		oGallery.GetGalleryInfo
	%>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0FD">��ȣ</td>
		<td bgcolor="#FFFFFF"><%=gal_sn%></td>
	</tr>
	<tr>
		<td width="100" bgcolor="#F0F0FD" align="center">�̹��� ����</td>
		<td bgcolor="#FFFFFF">
			<select name="gal_div">
				<option value=""<% if oGallery.FItemList(1).Fgal_div="" then Response.Write " selected" %>>����</option>
				<option value="W"<% if oGallery.FItemList(1).Fgal_div="W" then Response.Write " selected" %>>Work</option>
				<option value="D"<% if oGallery.FItemList(1).Fgal_div="D" then Response.Write " selected" %>>Drawing</option>
				<option value="P"<% if oGallery.FItemList(1).Fgal_div="P" then Response.Write " selected" %>>Photo</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�귣��</td>
		<td bgcolor="#FFFFFF">
			<%=oGallery.FItemList(1).Fsocname & " (" & oGallery.FItemList(1).Fsocname_kor & ")"%>
			<input type="hidden" name="designerid" value="<%=oGallery.FItemList(1).Fdesignerid%>">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">������ �̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="gal_imgorg" value="" size="55">
			<br>(1MB������ JPG Ȥ�� GIF������ ������ ���簢�� �̹����� ���ε����ּ���.)
			<% if oGallery.FItemList(1).Fgal_img400<>"" then %>
			<br><img src="<%=oGallery.FItemList(1).Fgal_img400%>" border="0">
			<br>Filename : <%=oGallery.FItemList(1).Fgal_imgorg%>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�̹��� ����</td>
		<td bgcolor="#FFFFFF">
			<textarea class="textarea" name="gal_desc" cols="60" rows="3"><%=oGallery.FItemList(1).Fgal_desc%></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">ǥ�ü���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="gal_sortNo" value="<%=oGallery.FItemList(1).Fgal_sortNo%>" size="3">
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0FD">�������</td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" value="Y"<% if oGallery.FItemList(1).Fgal_isusing="Y" then Response.Write " checked" %>>Y
			<input type="radio" name="isusing" value="N"<% if oGallery.FItemList(1).Fgal_isusing="N" then Response.Write " checked" %>>N
		</td>
	</tr>
	<tr bgcolor="#DDDDFF" >
		<td colspan="2" align="center">
				<input type="button" value=" ���� " onclick="subcheck();"> &nbsp;&nbsp;
				<input type="button" value=" ��� " onclick="history.back();">
		</td>
	</tr>
	
	<% end if %>
	
	</form>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->