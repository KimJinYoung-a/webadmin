<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfighterCls.asp"-->
<%
dim page,idx

page = request("page")
if page = "" then page=1
idx = request("idx")

dim ocate
set ocate = New CDesignFighter
ocate.FCurrPage = page
ocate.FPageSize=20
ocate.FRectidx = idx
ocate.GetDesignFighterList

dim i
%>
<script language='javascript'>
function popItemWindow(iid,frm){
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
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

function delitems(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idxarr.value = upfrm.idxarr.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

function MakeCommonUpdate(){
	if (confirm("������Ʈ�� �Ͻðڽ��ϱ�??????")){
	var popwin=window.open('/admin/sitemaster/lib/dofighterupdate.asp','fighterfresh','width=100,height=100');
	popwin.focus();
	}
}

</script>
<!-- �׼� ���� -->
<table width="800"   cellpadding="0" cellspacing="5" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="���ε��" onClick="javascript:location.href='design_fighter_write.asp?mode=add&menupos=<%=menupos%>';">			
		</td>
		<td align="right">
			<input type=button class="button" value="���� �� ������Ʈ" onclick="MakeCommonUpdate()">&nbsp;(�ʹ� ���� ������ ������)			
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<table width="800"   cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=ocate.FResultCount%></b>
			&nbsp;
			������ : <b><%=page%> / <%=ocate.FTotalpage%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100" align="center">idx</td>
		<td width="200" align="center">itemid</td>
		<td width="300" align="center">Image</td>
		<td width="80" align="center">�������</td>
		<td width="120" align="center">�����</td>
	</tr>
<% if ocate.FResultCount < 1 then %>
<% else %>
<% for i=0 to ocate.FResultCount-1 %>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= ocate.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
<td align="center"><a href="http://www.10x10.co.kr/designfighter/design_fighter_preview.asp?idx=<%= ocate.FItemList(i).Fidx %>" target="_blank"><%= ocate.FItemList(i).Fidx %><br>[�̸�����]</a></td>
	<td align="center"><a href="design_fighter_write.asp?mode=edit&idx=<% = ocate.FItemList(i).Fidx %>"><% = ocate.FItemList(i).Fitemid1 %> vs <% = ocate.FItemList(i).Fitemid2 %></a></td>
	<td align="center"><a href="design_fighter_write.asp?mode=edit&idx=<% = ocate.FItemList(i).Fidx %>"><img src="<%= ocate.FItemList(i).Ficon2 %>" border="0"> vs <img src="<%= ocate.FItemList(i).Ficon1 %>" border="0"></a></td>
	<td align="center"><% if ocate.FItemList(i).Fisusing = "Y" then %>Y<% else %><font color="red">N</font><% end if %></td>
	<td align="center"><%= FormatDate(ocate.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
</form>
<% next %>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">
	<% if ocate.HasPreScroll then %>
		<a href="?page=<%= ocate.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ocate.StartScrollPage to ocate.FScrollCount + ocate.StartScrollPage - 1 %>
		<% if i>ocate.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ocate.HasNextScroll then %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set ocate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->