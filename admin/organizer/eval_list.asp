<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->
<%
'#######################################################
'	History	:  2008.10.23 �ѿ�� ����
'	Description : ���ų�����
'#######################################################
%>
<%
dim page , i , itemid , page_eval
	page = requestCheckVar(request("page"),5)
	page_eval = requestCheckVar(request("page_eval"),5)
	itemid = requestCheckVar(request("itemid"),10)	
	if page = "" then page = 1
	if page_eval = "" then page_eval =1
		
dim oip_eval
set	oip_eval = new organizerCls
	oip_eval.FPageSize = 4
	oip_eval.FCurrPage = page_eval
	oip_eval.geteval_list

dim oip
set oip = new organizerCls
	oip.FPageSize = 50
	oip.FCurrPage = page	
	oip.frectitemid = itemid
	oip.feval_list()

%>

<script language="javascript">

//��ǰ�ı� �˾�
function eval_reg(organizerid,idx,mode){
	var popRegeval = window.open('/admin/organizer/eval_list_process.asp?organizerid='+organizerid+'&idx='+idx+'&mode='+mode,'popRegeval','width=1024,height=768,scrollbars=yes,resizable=yes')
}

//��ǰ�ı� ����Ʈ������ ���� �˾�
function eval_process(organizerid,idx,mode){
	var eval_process = window.open('/admin/organizer/eval_list_process.asp?organizerid='+organizerid+'&idx='+idx+'&mode='+mode,'eval_process','width=1024,height=768,scrollbars=yes,resizable=yes')
}
	

</script>

<!-- ����Ʈ ����Ʈ ���� -->
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip_eval.FResultCount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�� 1 �������� �ִ� 4���� ���ų���������Ʈ�������� ����˴ϴ�
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>��ID</td>
		<td>�����</td>
		<td>��ǰ�ı�</td>
		<td>����</td>
	</tr>

	<% For i =0 To  oip_eval.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oip_eval.fitemlist(i).fuserid %></td>
		<td><%= left(oip_eval.fitemlist(i).fregdate,10) %></td>
		<td><%= oip_eval.fitemlist(i).fcontents %></td>
		<td>
			<select  onchange="javascript:eval_process(this.value,<%=oip_eval.fitemlist(i).fidx %>,'update');">
				<option>����</option>
				<option value="Y" <% if oip_eval.fitemlist(i).fisusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if oip_eval.fitemlist(i).fisusing = "N" then response.write " selected" %>>N</option>				
			</select>
		</td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
<% End IF %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">
		
		<!-- ������ ���� -->
	    	<a href="?page_eval=1&itemid=<%=itemid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
			<% if oip_eval.HasPreScroll then %>
				<span class="list_link"><a href="?page_eval=<%= oip_eval.StartScrollPage-1 %>&itemid=<%=itemid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<% for i = 0 + oip_eval.StartScrollPage to oip_eval.StartScrollPage + oip_eval.FScrollCount - 1 %>
				<% if (i > oip_eval.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip_eval.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
				<% else %>
				<a href="?page_eval=<%= i %>&itemid=<%=itemid%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
				<% end if %>
			<% next %>
			<% if oip_eval.HasNextScroll then %>
				<span class="list_link"><a href="?page_eval=<%= i %>&itemid=<%=itemid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<a href="?page_eval=<%= oip_eval.FTotalpage %>&itemid=<%=itemid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
		<!-- ������ �� -->
		
		</td>
	</tr>
</table><br>
<!-- ����Ʈ ����Ʈ �� -->

<!-- ��ǰ�ı� ����Ʈ ���� -->
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip.FResultCount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			��ǰ�ڵ� <b><%= itemid %></b> &nbsp;�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>��ID</td>
		<td>�����</td>
		<td>��ǰ�ı�</td>
		<td>���</td>
	</tr>

	<% For i =0 To  oip.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oip.fitemlist(i).fuserid %></td>
		<td><%= left(oip.fitemlist(i).fregdate,10) %></td>
		<td><%= oip.fitemlist(i).fcontents %></td>
		<td><input type="button" class="button" value="����" onclick="javascript:eval_reg(<%= oip.fitemlist(i).forganizerid %>,<%= oip.fitemlist(i).fidx %>,'insert');"></td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
<% End IF %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" bgcolor="<%=adminColor("green")%>">
		
		<!-- ������ ���� -->
	    	<a href="?page=1&itemid=<%=itemid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
			<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>&itemid=<%=itemid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
				<% else %>
				<a href="?page=<%= i %>&itemid=<%=itemid%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&itemid=<%=itemid%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
			<% else %>
			&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
			<% end if %>
			<a href="?page=<%= oip.FTotalpage %>&itemid=<%=itemid%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
		<!-- ������ �� -->
		
		</td>
	</tr>
</table>
<!-- ��ǰ�ı� ����Ʈ �� -->

<% Set oip = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->