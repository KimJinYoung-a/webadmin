<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 �ѿ�� ����
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/othermall_main_event_rotationcls.asp"-->
<%

dim i
dim page, malltype
dim isusing, research
dim itemid

page = request("page")
isusing = request("isusing")
research = request("research")
itemid = request("itemid")
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
		 //refreshFrm.action = "http://uploadmain.10x10.co.kr/flash/link/MakeMainMdChoiceRotateFlash.asp" ;
		 refreshFrm.action = "<%=othermall%>/chtml/othermall_MakeMainMdChoiceRotateFlash.asp" ;
		 refreshFrm.submit();
	}
}

function NextPage(page){
	document.frm.page.value=page;
	document.frm.submit();
}

</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="refreshFrm" method="post">
	</form>
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�ܺθ� ���� ��õ ��ǰ</strong></font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			����Ʈ�� ���� <a href="javascript:RefreshMainMdChoiceRotateEventRec();">
			<img src="/images/refreshcpage.gif" width=19 align="absmiddle" border="0" alt="html�����"></a>
			��뱸�� :
			<select name="isusing" >
			<option value="" >��ü
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
			<option value="N" <% if isusing="N" then response.write "selected" %> >������
			</select>
			��ǰ��ȣ :
			<input type="text" name="itemid" value="<%= itemid %>" size=6 maxlength=6>			
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
			<a href="othermall_main_md_recommend_flash_write.asp?mode=write&menupos=<%= menupos %>">
			<p align="right"><img src="/images/icon_new_registration.gif" width="75" border="0" align="absmiddle"></a></p>	
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
	</form>
</table>
<!--ǥ ��峡-->

<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#BABABA">
<form name="vfrm" method="POST" action="">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="sUsing" value="<%= isusing %>">
	<tr bgcolor="#DDDDFF">
		<td width="30" align="center">ID</td>
		<td align="center" width="250">�̹���</td>
		<td align="center">link����</td>
		<td width="50" align="center">���ü���</td>
		<td width="100" align="center">�����</td>
		<td width="50" align="center">�������</td>
		<td width="50" align="center">ǰ������</td>
	</tr>
<% for i=0 to mdchoicerotate.FResultcount -1 %>
	<tr bgcolor="#FFFFFF">
		<td height="50" align="center">
			<input type="hidden" name="idx" value="<%= mdchoicerotate.FItemList(i).Fidx %>">
			<%= mdchoicerotate.FItemList(i).Fidx %>
		</td>
		<td align="center"><a href="othermall_main_md_recommend_flash_write.asp?mode=modify&idx=<%= mdchoicerotate.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= mdchoicerotate.FItemList(i).Fphotoimg %>" border=0 width="56"></a></td>
		<td height="50" align="left">
			<%= mdchoicerotate.FItemList(i).Flinkinfo %>
		</td>
		<td align="center">
			<input type="text" name="disporder" value="<%= mdchoicerotate.FItemList(i).FDisporder %>" size="3" style="text-align:right">
		</td>
		<td align="center">
			<%= FormatDateTime(mdchoicerotate.FItemList(i).Fregdate,2) %>
		</td>
		<td align="center">
			<select name="isusing">
				<option value="Y" <% if mdchoicerotate.FItemList(i).Fisusing="Y" then Response.Write "selected"%>>���</option>
				<option value="N" <% if mdchoicerotate.FItemList(i).Fisusing="N" then Response.Write "selected"%>>�Ұ�</option>
			</select>
		</td>
		<td align="center">
			<% if mdchoicerotate.FItemList(i).IsSoldOut then %>
			<font color="red">ǰ��</font>
			<% end if %>
		</td>
	</tr>
<% next %>
</form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set mdchoicerotate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->