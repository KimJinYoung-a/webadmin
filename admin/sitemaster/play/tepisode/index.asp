<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : T-Episode
' Hieditor : ����ȭ ����
'			 2022.07.07 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
'//�̺�Ʈ ����Ʈ
Dim page, isusing, viewtitle, playcate
playcate = 7
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
viewtitle	= request("viewtitle")
	isusing = requestCheckVar(request("isusing"),1)

If page = "" then page = 1

Dim oplaypick, i
Set oplaypick = new CPlayContents
	oplaypick.FPageSize			= 20
	oplaypick.FCurrPage			= page
	oplaypick.FRPlaycate		= playcate
	oplaypick.FRectIsusing		= isusing
	oplaypick.FRectViewTitle	= viewtitle
	oplaypick.sbGetphotopickList()
%>
<script type="text/javascript">
function AddNewContents(idx){
	location.href="/admin/sitemaster/play/tepisode/photopickEdit.asp?idx=" + idx;
}


function ItemIM(idx)
{
	var popitem;
	popitem = window.open('pop_itemReg.asp?idx='+idx,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
	popitem.focus();

}

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="index.asp";
	frm.submit();
}
function NextPage(page){
	frm.page.value = page;
	frm.submit();
}
</script>
<form name="frm" method="post" style="margin:0px;">
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		Ÿ��Ʋ : <input type="text" size="70"  name="viewtitle" size=20 value="<%=viewtitle%>" />&nbsp;&nbsp;
		��� : 
		<select name="isusing" class="select">
			<option value="">��ü</option>
			<option value="Y" <%= chkiif(isusing="Y","selected","") %> >Y</option>
			<option value="N" <%= chkiif(isusing="N","selected","") %> >N</option>
		</select>
	</td>
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSerach();">
	</td>
</tr>
</table>
</form>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="AddNewContents('0');">
	</td>
</tr>
</table>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= oplaypick.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> / <%=  oplaypick.FTotalpage %></b>
			</td>
			<td align="right"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%">��ȣ</td>
	<td width="5%">�̹���</td>
	<td width="10%">Ÿ��Ʋ</td>
	<td width="10%">TAG</td>
	<td width="10%">���</td>
	<td width="10%">�����</td>
	<td width="10%">���</td>
</tr>
<% If oplaypick.FResultCount > 0 then %>
<% For i = 0 to oplaypick.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oplaypick.FItemList(i).Fidx %></td>
	<td align="center"><img src="<%= oplaypick.FItemList(i).FPPimg %>" width=70 border=0></td>
	<td align="center"><%= ReplaceBracket(oplaypick.FItemList(i).FViewTitle) %></td>
	<td align="center"><%= Chkiif(oplaypick.FItemList(i).FTagCnt = "0","�������", "��ϿϷ�") %></td>
	<td align="center"><%= oplaypick.FItemList(i).FIsusing %></td>
	<td align="center"><%= Left(oplaypick.FItemList(i).FRegdate,10) %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="AddNewContents('<%= oplaypick.FItemList(i).Fidx %>');"/>
		<input type="button" class="button" value="��ǰ���/����[<%=oplaypick.FItemList(i).FitemCnt%>]" onclick="ItemIM('<%= oplaypick.FItemList(i).Fidx %>');" />
	</td>
</tr>
<% Next %>
<tr>
	<td colspan="20" align="center" bgcolor="#FFFFFF">
	 	<% if oplaypick.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oplaypick.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oplaypick.StartScrollPage to oplaypick.FScrollCount + oplaypick.StartScrollPage - 1 %>
			<% if i>oplaypick.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oplaypick.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% Else %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% End If %>
</table>
<% Set oplaypick = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
