<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������Ʈ
' Hieditor : 2014.03.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/tenbyten/invalid/invalid_user_cls.asp"-->

<%
Dim oinvalid, isusing, page, gubun, reload, i, userid, adminuserid
	isusing = requestcheckvar(request("isusing"),1)
	menupos = requestcheckvar(request("menupos"),10)
	page = requestcheckvar(request("page"),10)
	gubun = requestcheckvar(request("gubun"),12)
	reload = requestcheckvar(request("reload"),10)
	userid = requestcheckvar(request("userid"),32)

adminuserid=session("ssBctId")

if page = "" then page = 1
if reload="" and gubun = "" then gubun = "ONEVT"
if reload="" and isusing = "" then isusing = "Y"

'// �̺�Ʈ ����Ʈ
set oinvalid = new cinvalid_list
	oinvalid.FPageSize = 20
	oinvalid.FCurrPage = page
	oinvalid.frectisusing = isusing
	oinvalid.frectgubun = gubun
	oinvalid.frectuserid = userid
	oinvalid.getinvalid_list()
%>

<script type="text/javascript">

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

function edit_invalid(idx){
	<% if adminuserid="bseo" or adminuserid="boyishP" or C_ADMIN_AUTH then %>
		alert('�����ڱ����Դϴ�.');
		var edit_invalid = window.open('/admin/member/tenbyten/invalid/invalid_user_edit.asp?idx='+idx,'edit_invalid','width=800,height=400,scrollbars=yes,resizable=yes');
		edit_invalid.focus();
	<% else %>
		alert('��ϱ����� �����ϴ�.');
		return;
	<% end if %>
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* Ư�������� : <% Drawinvalidgubun "gubun", gubun, " onchange='frmsubmit(""1"");'" %>
		&nbsp;&nbsp;
		* ��뿩�� : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit(""1"");'" %>
		&nbsp;&nbsp;
		* ���̵� : <input type="text" name="userid" value="<%= userid %>" size=32 maxlength=32>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('1');">
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
		<input type="button" class="button" value="�űԵ��" onclick="edit_invalid('');">
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oinvalid.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oinvalid.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>Ư��������</td>
	<td>���̵�</td>
	<td>�ڸ�Ʈ</td>
	<td>��뿩��</td>
	<td>��������</td>
	<td>���</td>
</tr>
<% if oinvalid.FresultCount>0 then %>
	<% for i=0 to oinvalid.FresultCount-1 %>

	<% if oinvalid.FItemList(i).fisusing = "Y" then %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<% else %>
		<tr align="center" bgcolor="#c1c1c1" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#c1c1c1';>
	<% end if %>

		<td>
			<%= oinvalid.FItemList(i).fidx %>
		</td>
		<td>
			<%= getinvalidgubun(oinvalid.FItemList(i).fgubun) %>
		</td>
		<td>
			<%= printUserId(oinvalid.FItemList(i).finvaliduserid, 2, "*") %>
		</td>
		<td width=400>
			<%= chrbyte(oinvalid.FItemList(i).fcomment,100,"Y") %>
		</td>
		<td>
			<%= oinvalid.FItemList(i).fisusing %>
		</td>
		<td>
			<% if oinvalid.FItemList(i).flastupdate<>"" then %>
				<%= oinvalid.FItemList(i).flastupdate %>
			<% end if %>

			<% if oinvalid.FItemList(i).flastuserid<>"" then %>
				<Br>(<%= oinvalid.FItemList(i).flastuserid %>)
			<% end if %>
		</td>
		<td>
			<input type="button" onclick="edit_invalid('<%= oinvalid.FItemList(i).fidx %>'); return false;" value="����" class="button">
		</td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oinvalid.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oinvalid.StartScrollPage-1 %>&isusing=<%=isusing%>&gubun=<%=gubun%>&userid=<%=userid%>&reload=<%=reload%>&menupos=<%=menupos%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oinvalid.StartScrollPage to oinvalid.StartScrollPage + oinvalid.FScrollCount - 1 %>
				<% if (i > oinvalid.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oinvalid.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&gubun=<%=gubun%>&userid=<%=userid%>&reload=<%=reload%>&menupos=<%=menupos%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oinvalid.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&gubun=<%=gubun%>&userid=<%=userid%>&reload=<%=reload%>&menupos=<%=menupos%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oinvalid=nothing
%>
