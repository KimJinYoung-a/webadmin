<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ [�ȳ�����] �⺻ ī�װ�
' History : �̻� ����
'			2021.09.10 �ѿ�� ����(�̹����̻�Կ�û �ڻ�� �ʵ��߰�, �ҽ�ǥ��ȭ, ���Ȱ�ȭ)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%
' ��Ź��ü ���� ����� ��Ź��ü �Ϲ������� ���� ����.
if C_CSOutsourcingUser then
	if not(C_CSOutsourcingPowerUser) then
		response.write "������ �����ϴ�."
		response.end : dbget.close()
	end if
end if

dim gubunCode : gubunCode = "0001"
dim masterUseYN, detailUseYN, masterIdx, i, page, sitename
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	masterUseYN = requestcheckvar(request("masterUseYN"),1)
	detailUseYN = requestcheckvar(request("detailUseYN"),1)
	masteridx = requestcheckvar(getNumeric(request("masterIdx")),10)
	sitename = requestcheckvar(request("sitename"),32)

if (page = "") then
	page = 1
	masterUseYN = "Y"
	detailUseYN = "Y"
end if

dim oCReply
Set oCReply = new CReply
	oCReply.FPageSize = 50
	oCReply.FCurrPage = page
	oCReply.FRectMasterIDX = masterIdx
	oCReply.FRectMasterUseYN = masterUseYN
	oCReply.FRectDetailUseYN = detailUseYN
	oCReply.FRectGubunCode = gubunCode
	oCReply.FRectsitename = sitename
	oCReply.GetReplyDetailList()

%>

<script type="text/javascript">

function fnRegReplyDetail() {
	document.location.href = "/cscenter/board/cs_replydetail_view.asp?menupos=<%= menupos %>&masterIdx=<%= masterIdx %>&gubunCode=<%= gubunCode %>&masterUseYN=<%= masterUseYN %>";
}

function fnModiReplyDetail(idx) {
	document.location.href = "/cscenter/board/cs_replydetail_view.asp?menupos=<%= menupos %>&idx=" + idx;
}

function jsGotoPage(page) {
	var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� : <% Drawreplysitename "sitename", sitename, "" %>
		&nbsp;
		* �⺻ ī�װ�:
		<% Call drawSelectBoxReplyMaster("masterIdx", masterIdx, gubunCode, masterUseYN) %>
		&nbsp;
		* �⺻ ī�װ� ��뱸��:
		<select class="select" name="masterUseYN">
			<option></option>
			<option value="Y" <% if (masterUseYN = "Y") then %>selected<% end if %> >�����</option>
			<option value="N" <% if (masterUseYN = "N") then %>selected<% end if %> >������</option>
		</select>
		&nbsp;
		* �� ī�װ� ��뱸��:
		<select class="select" name="detailUseYN">
			<option></option>
			<option value="Y" <% if (detailUseYN = "Y") then %>selected<% end if %> >�����</option>
			<option value="N" <% if (detailUseYN = "N") then %>selected<% end if %> >������</option>
		</select>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" onclick="fnRegReplyDetail();" value="�űԵ��" class="button">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oCReply.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oCReply.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="60">IDX</td>
		<td width="40">����</td>
		<td align="left" width="150">�⺻ ī�װ���</td>
		<td align="left" width="250">�� ī�װ���</td>
		<td align="left" width="350">�ȳ�����</td>
		<td width="60">ǥ�ü���</td>
		<td width="30">���</td>
		<td width="150">����������</td>
		<td width="100">���</td>
    </tr>
	<% if oCReply.FresultCount>0 then %>
	<% for i=0 to oCReply.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">

    <% if oCReply.FItemList(i).FuseYN = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="<%= adminColor("tabletop") %>"; onmouseout=this.style.background='#FFFFFF';>
    <% else %>
    <tr align="center" bgcolor="#CCCCCC">
	<% end if %>
		<td height="25">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Fidx %></a>
		</td>
		<td><%= replysitename(oCReply.FItemList(i).fsitename) %></td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Ftitle %></a>
		</td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= oCReply.FItemList(i).Fsubtitle %></a>
		</td>
		<td align="left">
			<a href="javascript:fnModiReplyDetail(<%= oCReply.FItemList(i).Fidx %>)"><%= Left(oCReply.FItemList(i).Fcontents, 30) %>...</a>
		</td>
		<td>
			<%= oCReply.FItemList(i).FdispOrderNo %>
		</td>
		<td>
			<%= oCReply.FItemList(i).FuseYN %>
		</td>
		<td>
			<%= oCReply.FItemList(i).Flastupdate %>
		</td>
		<td></td>
    </tr>
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oCReply.HasPreScroll then %>
				<span class="list_link"><a href="javascript:jsGotoPage(<%= i %>)">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oCReply.StartScrollPage to oCReply.StartScrollPage + oCReply.FScrollCount - 1 %>
				<% if (i > oCReply.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oCReply.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:jsGotoPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oCReply.HasNextScroll then %>
				<span class="list_link"><a href="javascript:jsGotoPage(<%= i %>)">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oCReply = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
