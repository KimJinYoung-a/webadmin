<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.09 �ѿ�� 2008����Ʈ �̵�/�߰�/����
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
'// ���� ����
dim page, isusing, gal_div, designerid, lp
	page = request("page")
	isusing = request("isusing")
	gal_div = request("gal_div")
	designerid = request("designerid")
	
	if page="" then page=1
	if isusing="" then isusing="Y"

'// ��� ����
dim oGallery
	set oGallery = New CGallery
	oGallery.FCurrPage = page
	oGallery.FPageSize=20
	oGallery.FRectGal_div = gal_div
	oGallery.FRectDesignerId = designerid
	oGallery.FRectIsusing = isusing
	oGallery.GetGalleryList

'//���������� ��� 6�� ����Ʈ
dim oGalleryitem
	set oGalleryitem = New CGallery
	oGalleryitem.getgalleryitem
%>

<script language="javascript">

	//���ι�� ��ϻ�ǰ ��ǰã��
	function popItemWindow(tgf){
		var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
		popup_item.focus();
	}
	
	//���ι�� ��ǰ ���
	function regmainbanneritem()
	{
		if (searchForm.itemid.value==''){
			alert('��ǰ�ڵ带�Է��ϼ���');
			searchForm.itemid.focus();
		}else{
			moveForm.action="/admin/artist/artist_process.asp";
			moveForm.mode.value="mainbanneritem";
			moveForm.itemid.value = searchForm.itemid.value;
			moveForm.submit();
		}
	}

	function goPage(pg)
	{
		frm = document.moveForm;
		frm.action="";
		frm.page.value=pg;
		frm.submit();
	}

	function addItem()
	{
		frm = document.moveForm;
		frm.action="artist_gallery_edit.asp";
		frm.mode.value="add";
		frm.submit();
	}
	
	//��������
	function inquiry(){
		var inquiry = window.open('/admin/artist/artist_inquiry.asp','inquiry','width=1024,height=768,scrollbars=yes,resizable=yes');
		inquiry.focus();
	}	
	
	//��Ƽ��Ʈ��õ����
	function recommend(){
		var recommend = window.open('/admin/artist/artist_recommend.asp','recommend','width=1024,height=768,scrollbars=yes,resizable=yes');
		recommend.focus();
	}	

	function editItem(sn)
	{
		frm = document.moveForm;
		frm.action="artist_gallery_edit.asp";
		frm.mode.value="edit";
		frm.page.value="<%=page%>";
		frm.gal_sn.value=sn;
		frm.submit();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="searchForm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���м��� :
			<select name="gal_div">
				<option value=""<% if gal_div="" then Response.Write " selected" %>>����</option>
				<option value="W"<% if gal_div="W" then Response.Write " selected" %>>Work</option>
				<option value="D"<% if gal_div="D" then Response.Write " selected" %>>Drawing</option>
				<option value="P"<% if gal_div="P" then Response.Write " selected" %>>Photo</option>
			</select>&nbsp; &nbsp;
			�귣�弱�� :  
			<% Call DrawSelectBoxUseBrand("designerid",designerid) %>
			&nbsp; &nbsp;
			������� : <select name="isusing"><option value="Y">Yes</option><option value="N">No</option></select>
			<script language="javascript">
				document.searchForm.isusing.value="<%=isusing%>";
			</script>
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="searchForm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		</td>
	</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<!--
<tr>
	<td align="left">
		<font color="red">��Mainpage �ϴܹ��6��(�ֱٵ�ϻ�ǰ��ù��°�γ���)<br></font>
		<% for lp = 0 to oGalleryitem.ftotalcount - 1 %>
		<img src="<%= oGalleryitem.fitemlist(lp).flistimage120 %>" border=0 width=40 height=40>
		<% next %>
		��ǰ�ڵ� : <input type="text" name="itemid" size=10>
		<input type="button" class="button" value="ã��" onClick="popItemWindow('searchForm')">			
		<input type="button" class="button" value="����" onClick="regmainbanneritem()">					
	</td>
	<td align="right">	
	</td>
</tr>
-->
<tr>
	<td align="left">	
	</td>
	<td align="right">	
		<input type="button" value="������ �߰�" onclick="addItem()" class="button">
		<input type="button" value="��������" onclick="inquiry()" class="button">
		<input type="button" value="Artist��õ����" onclick="recommend()" class="button">
	</td>
</tr>	
</form>	
</table>
<!-- �׼� �� -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oGallery.FResultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oGallery.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oGallery.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center">��ȣ</td>
		<td width="100" align="center">����</td>
		<td width="250" align="center">��ü��</td>
		<td align="center">�̹���</td>
		<td width="50" align="center">�������</td>
		<td width="80" align="center">�����</td>
    </tr>
    
	<% for lp=0 to oGallery.FResultCount-1 %>
	    <tr align="center" bgcolor="#FFFFFF">
			<td align="center"><%= oGallery.FItemList(lp).Fgal_sn %></td>
			<td align="center"><%= oGallery.FItemList(lp).getGalDivName %></td>
			<td align="center"><%= oGallery.FItemList(lp).Fsocname_kor & "(" & oGallery.FItemList(lp).Fsocname & ")" %></td>
			<td align="center">
				<a href="javascript:editItem(<%= oGallery.FItemList(lp).Fgal_sn %>)">
				<img src="<%= oGallery.FItemList(lp).Fgal_img400 %>" width=50 height="50" border="0">
				</a>
			</td>
			<td align="center"><%= oGallery.FItemList(lp).Fgal_isusing %></td>
			<td align="center"><%= FormatDateTime(oGallery.FItemList(lp).Fgal_regdate,2) %></td>
	    </tr>   
	<% next %>
	
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oGallery.HasPreScroll then %>
				<a href="javascript:goPage(<%= oGallery.StartScrollPage-1 %>)">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for lp=0 + oGallery.StartScrollPage to oGallery.FScrollCount + oGallery.StartScrollPage - 1 %>
				<% if lp>oGallery.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(lp) then %>
				<font color="red">[<%= lp %>]</font>
				<% else %>
				<a href="javascript:goPage(<%= lp %>)">[<%= lp %>]</a>
				<% end if %>
			<% next %>
		
			<% if oGallery.HasNextScroll then %>
				<a href="javascript:goPage(<%= lp %>)">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<form name="moveForm" method="GET">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="">
<input type="hidden" name="gal_sn" value="">
<input type="hidden" name="isusing" value="<%=isusing%>">
<input type="hidden" name="gal_div" value="<%=gal_div%>">
<input type="hidden" name="designerid" value="<%=designerid%>">
<input type="hidden" name="itemid" size=10>
</form>
<%
	set oGallery = Nothing
	set oGalleryitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
