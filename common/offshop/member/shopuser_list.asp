<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����������α��Ѽ���
' Hieditor : 2011.01.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->

<%
dim omember , i , part_sn , page , adminyn ,SearchKey ,SearchString
	part_sn = request("part_sn")
	page = request("page")
	adminyn = request("adminyn")
	SearchKey = request("SearchKey")
	SearchString = request("SearchString")

if page="" then page=1
if adminyn = "" then adminyn = "Y"
			
set omember = new cshopuser_list
	omember.FPageSize = 50
	omember.FCurrPage = page
	omember.frectpart_sn = part_sn
	omember.frectadminyn = adminyn
	omember.frectSearchKey = SearchKey
	omember.frectSearchString = SearchString
	omember.getshopuser_list()
%>

<script language="javascript">
	
	function reg(page){
		frm.page.value = page;
		frm.submit();
	}
	
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �μ�:
		<%=printPartOption("part_sn", part_sn)%>
		&nbsp;&nbsp;	
		* ���λ�뿩�� : <% Call drawSelectBoxUsingYN("adminyn",adminyn) %>
		&nbsp;&nbsp;
		* <select name="SearchKey" class="select">
			<option value="" <% if SearchKey = "" then response.write " selected" %>>::����::</option>
			<option value="1" <% if SearchKey = "1" then response.write " selected" %>>���̵�</option>
			<option value="2" <% if SearchKey = "2" then response.write " selected" %>>����ڸ�</option>
			<option value="3" <% if SearchKey = "3" then response.write " selected" %>>���</option>
		</select>
		<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
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
<br>		
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

�� �������� ���������� ������ ��� �Ͻø�, ������ �α����� ������ ��� �Ŵ����� �ش� ������� �����Ͽ� �����ϽǼ� �ֽ��ϴ�.
<br>��ǥ�������� ù �α��ν� �ڵ����� ���õǴ� ������ ���մϴ�, �ݵ�� ���� ��Ź�帳�ϴ�.
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= omember.FTotalCount %></b>
		&nbsp;
		������ : <b><%= Page %> / <%= omember.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�μ�</td>
	<td>�����ȣ</td>
	<td>ID</td>
	<td>�̸�</td>
	<td>��ǥ����(���������)</td>
	<td>���</td>
</tr>
<% if omember.fresultcount > 0 then %>
<% for i=0 to omember.fresultcount - 1 %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td>
		<%= omember.FItemList(i).fpart_name %>
	</td>
	<td>
		<%= omember.FItemList(i).fempno %>
	</td>
	<td>
		<%= omember.FItemList(i).fid %>
	</td>		
	<td>
		<%= omember.FItemList(i).fcompany_name %>
	</td>
	<td align="left">
		<%
		if omember.FItemList(i).fshopfirst = "" or isnull(omember.FItemList(i).fshopfirst) then
			response.write "��������"
		else
			response.write omember.FItemList(i).fshopfirst&"/"&omember.FItemList(i).fshopname
		end if
		%>
		(<%= omember.FItemList(i).fshopcount %>��)
	</td>
	<td width=70>
		<input type="button" onclick="shopreg('<%= omember.FItemList(i).fempno %>');" value="����" class="button">
	</td>	
</tr>   
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
	<% if omember.HasPreScroll then %>
		<a href="javascript:reg('<%= omember.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omember.StartScrollPage to omember.FScrollCount + omember.StartScrollPage - 1 %>
		<% if i>omember.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:reg('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omember.HasNextScroll then %>
		<a href="javascript:reg('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set omember = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->