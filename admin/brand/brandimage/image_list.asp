<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  �귣�� �̹��� ���
' History : 2018-04-16 ����ȭ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/street/brandmainCls.asp" -->
<%
Dim page, lbrand, i
Dim makerid, isUsing, mode, frmName

mode = requestCheckVar(request("mode"),6)
frmName= requestCheckVar(request("frmName"),32)
if frmName="" then frmName="frm"
page    = requestCheckVar(request("page"),6)
makerid = requestCheckVar(request("makerid"),32)
isUsing = requestCheckVar(request("isusing"),1)
if isUsing="" then isUsing="1"

Response.write makerid

If page = "" Then page = 1

'// ��� ����
Set lbrand = New cBrandMain
	lbrand.FCurrPage = page
	lbrand.FRectMakerid = makerid
	lbrand.FRectIsUsing = chkIIF(isUsing="A","",isUsing)
	lbrand.FPageSize=20
	lbrand.sBrandImageGetList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
//�̹����űԵ�� & ����
function AddNewMainContents(idx){
	var AddNewMainContents = window.open('/admin/brand/brandimage/image_insert.asp?idx='+ idx,'AddNewMainContents','width=1250,height=650,scrollbars=yes,resizable=yes');
	AddNewMainContents.focus();
}

//���� ������ ���� ����
function SaveSelectedContents() {
	var selCnt = $("#frmList input:checkbox[name='idx']:checked").length;
	if(selCnt==0) {
		alert("���õ� �̹����� �����ϴ�.");
		return false;
	}

	if(confirm("�����Ͻ� "+selCnt+"���� �̹����� �����Ͻðڽ��ϱ�?")) {
		document.frmList.submit();
	}
}

$(function(){
	$("#frmList input:checkbox[name='idx']").click(function(){
		var ival = $(this).attr("data-idx");
		var iUs = $("#frmList input:radio[name='isus"+ival+"']:checked").val()
		$(this).val(ival+"/"+iUs);
	});
});

function fnSelectIMG(brandimage){
	opener.<%= frmName %>.mainimg.value = brandimage;
	$("#mainimg",opener.document).attr('src', brandimage);
	$("#imgurl",opener.document).html(brandimage);
	self.close();
}
</script>
<img src="/images/icon_arrow_link.gif"> <b>�귣�� �̹��� ����</b>
<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<% drawSelectBoxDesignerwithName "makerid",makerid %>
		/ ��뿩��
		<select name="isusing" class="select">
			<option value="A" <%=chkIIF(isUsing="A","selected","")%>>��ü</option>
			<option value="1" <%=chkIIF(isUsing="1","selected","")%>>���</option>
			<option value="0" <%=chkIIF(isUsing="0","selected","")%>>������</option>
		</select>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit()">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<% if mode="img" then %>
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0;">
<tr>
	<td align="left">
		<font style="color:red">�̹����� Ŭ���Ͻø� ���� �˴ϴ�.</font>
	</td>
</tr>
</table>
<% else %>
<table width="100%" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0;">
<tr>
	<td align="left">
		<input type="button" value="��������" class="button_auth" onclick="SaveSelectedContents();">
	</td>
	<td align="right">
		<input type="button" value="�űԵ��" class="button" onclick="AddNewMainContents('0');">
	</td>
</tr>
</table>
<% End If %>
<!-- �׼� �� -->

<form name="frmList" id="frmList" method="post" action="image_proc.asp" style="margin:0px;">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7">�˻���� : <b><%=lbrand.FTotalCount%></b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% if mode<>"img" then %>
	    <td></td>
		<% End If %>
		<td align="center">No.</td>
		<td align="center" width="200">Image</td>
	    <td align="center">�귣��ID</td>
	    <td align="center">�����</td>
	    <td align="center">������</td>
		<td align="center">��뿩��</td>
    </tr>
	<% If lbrand.FResultCount > 0 Then %> 
   	<% For i = 0 to lbrand.FResultCount - 1 %>
    <tr align="center" <%= chkiif(lbrand.FItemList(i).FIsusing,"bgcolor='#FFFFFF'","bgcolor='#DDDDDD'") %> >
		<% if mode<>"img" then %>
	    <td align="center"><input type="checkbox" name="idx" value="" data-idx="<%= lbrand.FItemlist(i).Fidx %>"></td>
		<% End If %>
		<% if mode="img" then %>
		<td align="center"><%= lbrand.FItemlist(i).Fidx %></td>
		<td align="center">
	    	<a href="javascript:fnSelectIMG('<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>');">
	    	<img src="<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>" style="width:100px; border:1px #FDFDFD; border-radius:3px;" />
	    	</a>
	    </td>
		<% else %>
		<td align="center"><a href="javascript:AddNewMainContents('<%= lbrand.FItemlist(i).Fidx %>');"><%= lbrand.FItemlist(i).Fidx %></a></td>
		<td align="center">
			<% if lbrand.FItemlist(i).Fbrandimage<>"" then %>
	    	<a href="javascript:AddNewMainContents('<%= lbrand.FItemlist(i).Fidx %>');">
	    	<img src="<%=uploadUrl%>/brandstreet/main/<%= lbrand.FItemlist(i).Fbrandimage %>" style="width:100px; border:1px #FDFDFD; border-radius:3px;" />
	    	</a>
			<% End If %>
	    </td>
		<% End If %>
	    <td align="center"><%= lbrand.FItemlist(i).Fmakerid %></td>
	    <td align="center"><%= lbrand.FItemlist(i).FRegdate %><br/><%= lbrand.FItemlist(i).Fadminid %></td>
	    <td align="center"><%= lbrand.FItemlist(i).Flastupdate %><br/><%= lbrand.FItemlist(i).Flastadminid %></td>
		<td align="center">
			<label><input type="radio" name="isus<%= lbrand.FItemlist(i).Fidx %>" value="1" <%=chkIIF(lbrand.FItemList(i).FIsusing,"checked","")%> />���</label>
			<label><input type="radio" name="isus<%= lbrand.FItemlist(i).Fidx %>" value="0" <%=chkIIF(lbrand.FItemList(i).FIsusing,"","checked")%> />������</label>
		</td>
	</tr>
	<% Next %>
	<% Else %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="7" align="center" class="page_link">[��ϵ� �����Ͱ� �����ϴ�.]</td>
	</tr>
	<% End If %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="7" align="center">
	       	<% If lbrand.HasPreScroll Then %>
				<span class="list_link"><a href="?page=<%= lbrand.StartScrollPage-1 %>">[pre]</a></span>
			<% Else %>
			[pre]
			<% End If %>
			<% for i = 0 + lbrand.StartScrollPage to lbrand.StartScrollPage + lbrand.FScrollCount - 1 %>
				<% if (i > lbrand.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(lbrand.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if lbrand.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->