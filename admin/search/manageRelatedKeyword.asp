<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim orgkeyword, relatedKeyword, modiType, useYN, page
dim i
dim research

research 		= request("research")
orgkeyword 		= Trim(request("orgkeyword"))
relatedKeyword 	= Trim(request("relatedKeyword"))
modiType 		= Trim(request("modiType"))
useYN 			= Trim(request("useYN"))
page			= requestCheckvar(request("page"),10)

if (research = "") then
	useYN = "Y"
end if

if (page="") then page = 1


'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FPageSize = 50
osearchKeyword.FCurrPage = page

osearchKeyword.FRectOrgKeyword		= orgkeyword
osearchKeyword.FRectRelatedKeyword	= relatedKeyword
osearchKeyword.FRectModiType		= modiType
osearchKeyword.FRectUseYN			= useYN

osearchKeyword.getRelatedKeywordModi_Paging

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsPopRelatedKeywordAdd() {
    var popwin = window.open('popRelatedKeywordAdd.asp','jsPopRelatedKeywordAdd','width=330,height=220,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function jsDelRelatedKeyword(idx) {
	var ret = confirm("�����Ͻðڽ��ϱ�?");
	if(ret){
		var frm = document.frmAct;
		frm.mode.value = "del";
		frm.idx.value = idx;
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			���˻��� : <input type="text" class="text" name="orgkeyword" value="<%= orgkeyword %>">
			&nbsp;
			�����˻��� : <input type="text" class="text" name="relatedKeyword" value="<%= relatedKeyword %>">
			&nbsp;
			���� :
			<select class="select" name="modiType">
				<option value=""></option>
				<option value="A" <% if (modiType = "A") then %>selected<% end if %> >�߰�</option>
				<option value="D" <% if (modiType = "D") then %>selected<% end if %> >����</option>
			</select>
			��뿩�� :
			<select class="select" name="useYN">
				<option value=""></option>
				<option value="Y" <% if (useYN = "Y") then %>selected<% end if %> >Y</option>
				<option value="N" <% if (useYN = "N") then %>selected<% end if %> >N</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			* �˻����� �ݿ��� 07, 11, 15, 19 �ÿ� �̷�����ϴ�.
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<input type="button" class="button" value=" ��� " onClick="jsPopRelatedKeywordAdd()">

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			�˻���� : <b><%= osearchKeyword.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= osearchKeyword.FTotalPage %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">IDX</td>
		<td width="150">���˻���</td>
		<td width="150">�����˻���</td>
		<td width="80">����ġ</td>
		<td width="50">����</td>
		<td width="100">�����</td>
		<td width="50">��뿩��</td>
		<td width="150">�����</td>
		<td>���</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Fidx %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).ForgKeyword %></td>
		<td align="center"><%= osearchKeyword.FItemList(i).FrelatedKeyword %></td>
		<td align="center">
			<% if (osearchKeyword.FItemList(i).FmodiType = "A") then %>
			<%= osearchKeyword.FItemList(i).FsearchCount %>
			<% end if %>
		</td>
		<td align="center">
			<% if (osearchKeyword.FItemList(i).FmodiType = "D") then %><font color="red"><% end if %>
			<%= osearchKeyword.FItemList(i).GetModiTypeName %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Freguserid %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).FuseYN %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Fregdate %>
		</td>
		<td align="left">
			<input type="button" class="button" value=" ���� " onClick="jsDelRelatedKeyword(<%= osearchKeyword.FItemList(i).Fidx %>)">
		</td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="9">
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21" align="center">
			<% if osearchKeyword.HasPreScroll then %>
			<a href="javascript:NextPage('<%= osearchKeyword.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for i=0 + osearchKeyword.StartScrollPage to osearchKeyword.FScrollCount + osearchKeyword.StartScrollPage - 1 %>
				<% if i>osearchKeyword.FTotalPage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if osearchKeyword.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>
<form name="frmAct" method="post" action="manageRelatedKeyword_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
