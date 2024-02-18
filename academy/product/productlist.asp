<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_productcls.asp"-->

<%

dim oacademyprd
dim page
dim makerid, sellyn, isusing, selBest

page = RequestCheckvar(request("page"),10)
if page="" then page=1
selBest = RequestCheckvar(request("selBest"),1)

set oacademyprd = new CAcademyProduct
oacademyprd.FCurrPage = page
oacademyprd.FPageSize = 20
oacademyprd.FRectMakerid = makerid
oacademyprd.FRectSellYn = sellyn
oacademyprd.FRectIsUsing = isusing
oacademyprd.FRectBest	= selBest

oacademyprd.GetProductList


dim i
%>
<script language='javascript'>
function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function popitemsearch(frm){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist.asp?designerid=" + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function AddIttems(){
	var ret=confirm('���� ��ǰ�� �߰� �Ͻðڽ��ϱ�?');

	if(ret){
		frmbuf.submit();
	}
}

function QuickAdd(frm){
	if (frm.itemidarr.value.length<1){
		alert('���� �Է��ϼ���.');
		frm.itemidarr.focus();
		return;
	}

	var ret=confirm('��ǰ�� �߰� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.submit();
	}
}

function DellItems(frm, stype){
	var ret=confirm('���� ��ǰ�� ���� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.mode.value = stype;
		frm.submit();
	}
}

//����Ʈ ���
function BestItems(frm, stype){
	var ret=confirm('���� ��ǰ�� ����Ʈ�� ��� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.mode.value = stype;
		frm.submit();
	}
}

//����Ʈ ���
function BestCancel(frm,stype,sId){
var ret=confirm('����Ʈ�� ��� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.mode.value = stype;
		frm.bestId.value = sId;
		frm.submit();
	}
}

//�Ǹż���
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit_aca','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// �̹�������
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage_aca','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" >
<form name="frmbuf" method="post" action="doacademyproduct.asp">
<input type="hidden" name="mode" value="addArr">
<tr>
	<td bgcolor="#FFFFFF" colspan="3">
	<input type="text" name="itemidarr" size="90" maxlength="90">
	<input type="button" value="��ǰ�ڵ���߰�" onclick="QuickAdd(frmbuf)">
	</td>
</tr>
<tr>
	<td width="50">
		<input type="button" value="���û�ǰ����" onclick="DellItems(frmlist,'dellarr');">
	</td>
	<td width="75%">
		<input type="button" value="���û�ǰ ����Ʈ��� " onclick="BestItems(frmlist,'bestarr');">
	</td>
	<td width="50" bgcolor="#FFFFFF" align="right">
		<input type="button" value="��Ͽ��������߰�" onclick="popitemsearch('frmbuf.itemidarr');">
	</td>
</tr>
</form>
</table>

<br>

<!-- ��� �˻��� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" >
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
       	����Ʈ ���� : 
       	<select name="selBest" onchange="javascript:document.frm.submit();">
       	<option value="">��ü</option>
       	<option value="1" <%IF selBest = "1" THEN%>selected<%END IF%>>����Ʈ</option>
       	<option value="2" <%IF selBest = "2" THEN%>selected<%END IF%>>����Ʈ ����</option>
       	</select>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- ��� �˻��� �� -->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmlist" method=post action="doacademyproduct.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="bestId" value="">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="13" align="right">�˻��Ǽ� : <%= oacademyprd.FTotalCount %> �� Page : <%= page %>/<%= oacademyprd.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="20"></td>
		<td align="center" width="50">��ǰ��ȣ</td>
		<td align="center" width="50">�̹���</td>
		<td align="center" width="70">�귣��</td>
		<td align="center">��ǰ��</td>
		<td align="center" width="40">����<br>����</td>
		<td align="center" width="50">�ǸŰ�</td>
		<td align="center" width="50">���԰�</td>
		<td align="center" width="50">�Ǹ�</td>
		<td align="center" width="50">����</td>
		<td align="center" width="50">���</td>
		<td align="center" width="50">����Ʈ</td>
	</tr>
	<% for i=0 to oacademyprd.FResultCount -1 %>
	<tr <% if oacademyprd.FITemList(i).FisBest = "Y" then%>bgcolor="#F3F3FF"<%else%>bgcolor="#FFFFFF"<%end if%> align="center">
		<td><input type="checkbox" name="itemidarr" value="<%= oacademyprd.FITemList(i).FItemID %>" onClick="AnCheckClick(this);"></td>
		<td><a href="javascript:PopItemSellEdit('<%= oacademyprd.FITemList(i).FItemID %>')"><%= oacademyprd.FITemList(i).FItemID %></a></td>
		<td><a href="javascript:editItemImage('<%= oacademyprd.FITemList(i).FItemID %>')"><img src="<%= oacademyprd.FITemList(i).FSmallImage %>" width="50" border="0"></a></td>
		<td align="left"><%= oacademyprd.FITemList(i).FMakerid %></td>
		<td align="left"><a href="/admin/itemmaster/itemmodify.asp?itemid=<%= oacademyprd.FITemList(i).FItemID %>&menupos=594" target="_blank"><%= oacademyprd.FITemList(i).FItemName %></a></td>
		<td><%= oacademyprd.FITemList(i).GetMWdivStr %></td>
		<td align="right"><%= FormatNumber(oacademyprd.FITemList(i).FSellcash,0) %></td>
		<td align="right"><%= FormatNumber(oacademyprd.FITemList(i).FBuycash,0) %></td>
		<td><%= oacademyprd.FITemList(i).FSellyn %></td>
		<td><%= oacademyprd.FITemList(i).GetLimitStr %></td>
		<td>
			<% if oacademyprd.FITemList(i).IsSoldOut then %>
			<font color="red">ǰ��</font>
			<% end if %>
		</td>
		<td>
			<%if oacademyprd.FITemList(i).FisBest = "Y" then%>
				<font color="red">����Ʈ</font><br>
				<a href="javascript:BestCancel(frmlist,'unbest',<%= oacademyprd.FITemList(i).FItemID %>);">[x���]</a>
			<%END IF%>
		</td>
	</tr>
	<% next %>
	<tr>
		<td align="center" colspan="13" bgcolor="#F0F0FD">
			<!-- ������ ���� -->
				<%
				if oacademyprd.HasPreScroll then
					Response.Write "<a href='javascript:NextPage(" & oacademyprd.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oacademyprd.StarScrollPage to oacademyprd.FScrollCount + oacademyprd.StarScrollPage - 1

					if i>oacademyprd.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:NextPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oacademyprd.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:NextPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- ������ �� -->
		</td>
	</tr>
</form>
</table>
<%
set oacademyprd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
