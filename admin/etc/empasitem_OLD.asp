<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/yahooitemcls.asp"-->
<%
dim oempas
dim page
page = request("page")
if page="" then page=1

dim ix

set oempas = new CYahooItemList
oempas.FPageSize = 2000
oempas.FCurrPage = page
oempas.GetEmpasItem

%>
�� �Ǽ� : <%= oempas.FtotalCount %> <br>
������ : <%= page %>/<%= oempas.FtotalPage %><br>

<%
dim bufstr
dim fso, FileName,tFile,appPath
appPath = server.mappath("/admin/etc/empasitem/") + "\"
FileName = "empasitem" + CStr(page) + ".txt"

Set fso = CreateObject("Scripting.FileSystemObject")

Set tFile = fso.CreateTextFile(appPath & FileName )

for ix = 0 to oempas.FresultCount - 1
	tFile.WriteLine "``begin``"
	tFile.WriteLine "``���θ���``�ٹ�����"
	tFile.WriteLine "``��ǰID``" + CStr(oempas.FItemList(ix).FItemId)
	tFile.WriteLine "``��з�``" + oempas.FItemList(ix).GetEmpasLargeCode
	tFile.WriteLine "``�ߺз�``" + oempas.FItemList(ix).GetEmpasMidCode
	tFile.WriteLine "``�Һз�``" + oempas.FItemList(ix).GetEmpasSmallCode
	tFile.WriteLine "``���з�``" + oempas.FItemList(ix).GetEmpasSeCode
	tFile.WriteLine "``�ڻ�з�``" + oempas.FItemList(ix).GetTenbytenCategoryName
	tFile.WriteLine "``�̹���``" + oempas.FItemList(ix).getImage
	tFile.WriteLine "``��ǰ��``" + oempas.FItemList(ix).GetEmpasItemName
	tFile.WriteLine "``�𵨸�``" + oempas.FItemList(ix).GetModelName
	tFile.WriteLine "``������``" + oempas.FItemList(ix).GetSourceArea
	tFile.WriteLine "``��ǰURL``" + oempas.FItemList(ix).GetEmpasUrl
	tFile.WriteLine "``��ǰ����``" + oempas.FItemList(ix).GetEmpasDesc
	tFile.WriteLine "``������``" + oempas.FItemList(ix).GetJejosa
	tFile.WriteLine "``�귣��``" + oempas.FItemList(ix).GetBrandName
	tFile.WriteLine "``�Һ��ڰ�``" + CStr(oempas.FItemList(ix).GetOrgSellcash)
	tFile.WriteLine "``�ǸŰ�``" + CStr(oempas.FItemList(ix).GetRealSellcash)
	tFile.WriteLine "``�����``" + oempas.FItemList(ix).GetLastEditDate
	tFile.WriteLine "``end``"

next

tFile.Close
Set tFile = Nothing
 Set fso = Nothing

%>
<table width="600">
<tr>
	<td colspan="12" align="center"><a href="empasitem/<%= FileName %>">�ٿ�ε�</a></td>
</tr>
<tr>
	<td colspan="12" align="center">
	<% if oempas.HasPreScroll then %>
		<a href="?page=<%= oempas.StarScrollPage-1 %>">[pre]</a>
	<% else %>
	<% end if %>

	<% for ix=0 + oempas.StarScrollPage to oempas.FScrollCount + oempas.StarScrollPage - 1 %>
		<% if ix > oempas.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(ix) then %>
		<font color="red">[<%= ix %>]</font>
		<% else %>
		<a href="?page=<%= ix %>">[<%= ix %>]</a>
		<% end if %>
	<% next %>

	<% if oempas.HasNextScroll then %>
		<a href="?page=<%= ix %>">[next]</a>
	<% else %>
	<% end if %>
	</td>
</tr>
</table>
<%
set oempas = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->