<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڳʰ���
' History : 2009.09.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->

<%
Dim oip,i,page , lecturer_id , isusing
	menupos = requestCheckvar(request("menupos"),10)
	page = requestCheckvar(request("page"),10)
	lecturer_id = requestcheckvar(request("lecturer_id"),32)
	isusing = requestcheckvar(request("isusing"),1)
		
	if page = "" then page = 1
				
'// �̺�Ʈ ����Ʈ
set oip = new cgood_onelist
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectisusing = isusing
	oip.frectlecturer_id = lecturer_id
	oip.fgood_item_list()

	If oip.FTotalCount < 1 Then
	'	response.write "<script>alert('���� �������� ����ϼ���.'); self.close();</script>"
	End If
%>

<script language="javascript">

// ��ǰ���&����
function reg_item(lecturer_id,idx){
	frm_corner.lecturer_id.value = lecturer_id;
	frm_corner.idx.value = idx;
	frm_corner.action = '/lectureadmin/corner/good_item_reg.asp';
	frm_corner.submit();
}

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}	

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<select name="isusing">
				<option value="">��뿩��</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp;����ID: <%=lecturer_id%><input type="hidden" name="lecturer_id" value="<%=lecturer_id%>">
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		
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
			<input type="button" class="button" value="��ǰ���" onclick="reg_item('<%=lecturer_id%>','');">				
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
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center" >�̹���</td>
		<td align="center" >����id</td>		
		<td align="center">��뿩��</td>		
		<td align="center">���</td>
    </tr>
	<% 
	if oip.FresultCount>0 then    
	
	for i=0 to oip.FresultCount-1 
	%>
	<form action="" name="frmBuyPrc<%=i%>" method="get">		   
    <% if oip.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa">
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center"><img src="<%= oip.FItemList(i).fimage_400x400 %>" width=40 height=40></td>
		<td align="center"><%= oip.FItemList(i).flecturer_id %></td>		
		<td align="center"><%= oip.FItemList(i).fisusing %></td>
		<td align="center">
			<input type="button" class="button" value="����" onclick="reg_item('<%= oip.FItemList(i).flecturer_id %>',<%= oip.FItemList(i).fidx %>);">			
		</td>			
    </tr>   
	</form>
	<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>&lecturer_id=<%=lecturer_id%>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&lecturer_id=<%=lecturer_id%>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&lecturer_id=<%=lecturer_id%>&isusing=<%=isusing%>">[next]</a></span>
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
<form name="frm_corner" method="get">
	<input type="hidden" name="idx" >
	<input type="hidden" name="lecturer_id" >
</form>
<%
	set oip = nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->