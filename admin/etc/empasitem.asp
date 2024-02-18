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
총 건수 : <%= oempas.FtotalCount %> <br>
페이지 : <%= page %>/<%= oempas.FtotalPage %><br>

<%
dim bufstr
dim fso, FileName,tFile,appPath
appPath = server.mappath("/admin/etc/empasitem/") + "\"
FileName = "empasitem" + CStr(page) + ".txt"

Set fso = CreateObject("Scripting.FileSystemObject")

Set tFile = fso.CreateTextFile(appPath & FileName )

for ix = 0 to oempas.FresultCount - 1
	tFile.WriteLine "``begin``"
	tFile.WriteLine "``쇼핑몰명``텐바이텐"
	tFile.WriteLine "``상품ID``" + CStr(oempas.FItemList(ix).FItemId)
	tFile.WriteLine "``대분류``" + oempas.FItemList(ix).GetEmpasLargeCode
	tFile.WriteLine "``중분류``" + oempas.FItemList(ix).GetEmpasMidCode
	tFile.WriteLine "``소분류``" + oempas.FItemList(ix).GetEmpasSmallCode
	tFile.WriteLine "``세분류``" + oempas.FItemList(ix).GetEmpasSeCode
	tFile.WriteLine "``자사분류``" + oempas.FItemList(ix).GetTenbytenCategoryName
	tFile.WriteLine "``이미지``" + oempas.FItemList(ix).getImage
	tFile.WriteLine "``상품명``" + oempas.FItemList(ix).GetEmpasItemName
	tFile.WriteLine "``모델명``" + oempas.FItemList(ix).GetModelName
	tFile.WriteLine "``원산지``" + oempas.FItemList(ix).GetSourceArea
	tFile.WriteLine "``상품URL``" + oempas.FItemList(ix).GetEmpasUrl
	tFile.WriteLine "``상품설명``" + oempas.FItemList(ix).GetEmpasDesc
	tFile.WriteLine "``제조사``" + oempas.FItemList(ix).GetJejosa
	tFile.WriteLine "``브랜드``" + oempas.FItemList(ix).GetBrandName
	tFile.WriteLine "``소비자가``" + CStr(oempas.FItemList(ix).GetOrgSellcash)
	tFile.WriteLine "``판매가``" + CStr(oempas.FItemList(ix).GetRealSellcash)
	tFile.WriteLine "``등록일``" + oempas.FItemList(ix).GetLastEditDate
	tFile.WriteLine "``end``"

next

tFile.Close
Set tFile = Nothing
 Set fso = Nothing

%>
<table width="600">
<tr>
	<td colspan="12" align="center"><a href="empasitem/<%= FileName %>">다운로드</a></td>
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