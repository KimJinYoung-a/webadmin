<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/video_cls.asp"-->

<%
Dim oip, i,page , rumour_id , isusing, vGubun, vParam, vCateCD2
	menupos = RequestCheckvar(request("menupos"),10)
	page = RequestCheckvar(request("page"),10)
	isusing = requestcheckvar(request("isusing"),1)
	vCateCD2 = RequestCheckvar(request("CateCD2"),10)
		
	if page = "" then page = 1
		
	vParam = "&menupos="&menupos&"&isusing="&isusing&"&CateCD2="&vCateCD2&""
				
	set oip = new cvideo
	oip.frectisusing = isusing
	oip.frectcate2 = vCateCD2
	oip.FPageSize = 10
	oip.FCurrPage = page
	oip.fvideo_list
%>

<script language="javascript">

document.domain = "10x10.co.kr";

// ������&����
function reg_video(video_id){
	var reg_video = window.open('/academy/corner/video_reg.asp?idx='+video_id,'reg_video','width=800,height=600,scrollbars=yes,resizable=yes');
	reg_video.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select name="isusing">
				<option value="">��뿩��</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			ī�װ� : <% Call DrawSelectBoxLecCategoryLarge("CateCD2",vCateCD2,"N")%>
		</td>	
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">				
			<input type="button" class="button" value="video���" onclick="reg_video('');">				
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" >IDX</td>
		<td align="center">����</td>	
		<td align="center">�̹���</td>
		<td align="center">��뿩��</td>
		<td align="center">���</td>
    </tr>
    
	<% 
	if oip.FresultCount>0 then    
	
	for i=0 to oip.FresultCount-1
	%>
		<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><%= oip.FItemList(i).fidx %></td>
			<td align="center"><%= oip.FItemList(i).ftitle %></td>
			<td align="center"><img src="<%= oip.FItemList(i).fimage_url %>" height="50"></td>
			<td align="center"><%= oip.FItemList(i).fisusing %></td>
			<td align="center"><input type="button" class="button" value="����" onclick="reg_video('<%= oip.FItemList(i).fidx %>');"></td>
		</tr>
	<%
	next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %><%=vParam%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %><%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %><%=vParam%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
</table>

<%
set oip = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
