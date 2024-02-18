<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/sitemaster/diy_main_diykitcls.asp"-->
<%

dim i
dim page, malltype
dim isusing, research
dim itemid

page = RequestCheckvar(request("page"),10)
isusing = RequestCheckvar(request("isusing"),1)
research = RequestCheckvar(request("research"),2)
itemid = RequestCheckvar(request("itemid"),10)
if (page = "") then
        page = "1"
end if

if research="" and isusing="" then isusing="Y"

'==============================================================================
dim mdchoicerotate
set mdchoicerotate = new CMainMdChoiceRotate

mdchoicerotate.FCurrPage = CInt(page)
mdchoicerotate.FPageSize = 20
mdchoicerotate.FRectIsUsing = isusing
mdchoicerotate.FRectItemID = itemid
mdchoicerotate.list

%>
<script language='javascript'>
function RefreshMainMdChoiceRotateEventRec(){
	if (confirm('���� �������� ���� �Ͻðڽ��ϱ�?')){
		 var popwin = window.open('','refreshFrm','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm";
		 refreshFrm.action = "<%=wwwFingers%>/chtml/diymain_diykit_make_xml.asp" ;
		 refreshFrm.submit();
	}
}

function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

function frmChange()
{
	var vfm = document.vfrm;
	if(confirm("���ü��� �� ��������� ���� ��Ͽ� ���̴� ���� �״�� ��� ����˴ϴ�.\n��ü ���� �Ͻðڽ��ϱ�?"))
	{
		vfm.action="doMainMdChoiceChange.asp";
		vfm.submit()
	}
	else
		return;
}

var chkUsing="<%=isusing%>";
function usingAllChange()
{
	if(chkUsing=="Y") { chkUsing = "N"; }
	else { chkUsing = "Y"; }

	for (var i=0;i<document.vfrm.isusing.length;i++){
		document.vfrm.isusing[i].value=chkUsing;
	}
}
</script>

<form name="refreshFrm" method="post">
</form>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		����Ʈ�� ���� <a href="javascript:RefreshMainMdChoiceRotateEventRec();"><img src="/images/refreshcpage.gif" width=19 align="absmiddle" border="0" alt="html�����"></a>
		&nbsp;&nbsp;
		��뱸�� :
		<select name="isusing" >
		<option value="" >��ü
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
		<option value="N" <% if isusing="N" then response.write "selected" %> >������
		</select>
		&nbsp;
		��ǰ��ȣ :
		<input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=9>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<form name="vfrm" method="POST" action="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sUsing" value="<%= isusing %>">
	<tr bgcolor="<%= adminColor("topbar") %>">
		<td colspan="7" align="right" height="30">
		<!--
			<a href="javascript:frmChange()"><img src="/images/icon_change.gif" width="45" border="0" align="absmiddle"></a> &nbsp;
		//-->
			<a href="diy_main_diykit_write.asp?mode=write&menupos=<%= menupos %>"><img src="/images/icon_new_registration.gif" width="75" border="0" align="absmiddle"></a>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("topbar") %>">
		<td align="center">ID</td>
		<td align="center">��ǰ�̹���</td>
		<td align="center">��ǰ��ȣ</td>
		<td align="center">link����</td>
		<td align="center">���ü���</td>
		<td align="center">�����</td>
		<td align="center"><a href="javascript:usingAllChange()">�������</a></td>
		<!--<td align="center">ǰ������</td>//-->
	</tr>
<% for i=0 to mdchoicerotate.FResultcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td height="50" align="center">
			<input type="hidden" name="idx" value="<%= mdchoicerotate.FItemList(i).Fidx %>">
			<%= mdchoicerotate.FItemList(i).Fidx %>
		</td>
		<!--<td align="center"><a href="diy_main_diykit_write.asp?mode=modify&idx=<%= mdchoicerotate.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= mdchoicerotate.FItemList(i).Fphotoimg %>" border=0 width="56"></a></td>//-->
		<td align="center"><a href="diy_main_diykit_write.asp?mode=modify&idx=<%= mdchoicerotate.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= mdchoicerotate.FItemList(i).Fsmallimage %>" border=0></a></td>
		<td align="center"><%= mdchoicerotate.FItemList(i).Flinkitemid%></td>
		<td height="50" align="left">
			<a href="diy_main_diykit_write.asp?mode=modify&idx=<%= mdchoicerotate.FItemList(i).Fidx %>&menupos=<%= menupos %>"><%= mdchoicerotate.FItemList(i).Flinkinfo %></a>
			&nbsp;&nbsp;&nbsp;
			[<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=mdchoicerotate.FItemList(i).Flinkitemid%>" target="_blank">��ǰ������������</a>]
		</td>
		<td align="center">
			<!--<input type="text" name="disporder" value="<%= mdchoicerotate.FItemList(i).FDisporder %>" size="3" style="text-align:right">//-->
			<%= mdchoicerotate.FItemList(i).FDisporder %>
		</td>
		<td align="center">
			<%= FormatDateTime(mdchoicerotate.FItemList(i).Fregdate,2) %>
		</td>
		<td align="center">
			<!--
			<select name="isusing">
				<option value="Y" <% if mdchoicerotate.FItemList(i).Fisusing="Y" then Response.Write "selected"%>>���</option>
				<option value="N" <% if mdchoicerotate.FItemList(i).Fisusing="N" then Response.Write "selected"%>>�Ұ�</option>
			</select>
			//-->
			<% if mdchoicerotate.FItemList(i).Fisusing="Y" then Response.Write "���" else Response.Write "�Ұ�" end if %>
		</td>
		<!--
		<td align="center">
			<% if mdchoicerotate.FItemList(i).IsSoldOut then %>
			<font color="red">ǰ��</font>
			<% end if %>
		</td>
		//-->
	</tr>
<% next %>
	<tr>
		<td colspan="7" height="1" bgcolor="#AAAAAA"></td>
	</tr>
	<tr>
		<td colspan="7" align="center">
			<% if mdchoicerotate.HasPreScroll then %>
				<a href="javascript:NextPage('<%= mdchoicerotate.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + mdchoicerotate.StarScrollPage to mdchoicerotate.FScrollCount + mdchoicerotate.StarScrollPage - 1 %>
				<% if i>mdchoicerotate.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if mdchoicerotate.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</form>
</table>
<%
set mdchoicerotate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->