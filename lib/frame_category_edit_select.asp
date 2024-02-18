<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual ="/lib/classes/items/category_selectcls.asp" -->
<%
dim ocate,ix
dim cd1,cd2,cd3
dim search_code
dim form_name
dim element_name
dim cd1select

search_code = request("search_code")
form_name = request("form_name")
element_name = request("element_name")


search_code = split(search_code,",")

cd1 = search_code(0)
cd2 = search_code(1)
cd3 = search_code(2)

dim ocd1,ocd2,ocd3

set ocd1 = New CCategory
ocd1.CategoryCodeLarge


set ocd2 = New CCategory
ocd2.FRectCD1 = cd1
ocd2.FRectCD2 = cd2
ocd2.CategoryCodeMid2

set ocd3 = New CCategory
ocd3.FRectCD1 = cd1
ocd3.FRectCD2 = cd2
ocd3.FRectCD3 = cd3
ocd3.CategoryCodeSmall2


if cd1 = "10" then
cd1select = 0
elseif cd1 = "15" then
cd1select = 1
elseif cd1 = "20" then
cd1select = 2
elseif cd1 = "25" then
cd1select = 3
elseif cd1 = "30" then
cd1select = 4
elseif cd1 = "35" then
cd1select = 5
elseif cd1 = "40" then
cd1select = 6
elseif cd1 = "45" then
cd1select = 7
elseif cd1 = "50" then
cd1select = 8
end if
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html">
<script>
var selectBox = parent.<% = form_name %> ;

selectBox.cd1.length = <% = ocd1.FResultCount + 1 %>;
<% for ix=0 to ocd1.FResultCount - 1 %>
selectBox.cd1.options[<% = ix + 1 %>].value= '<% = ocd1.FItemList(ix).FCD1 %>' ;
selectBox.cd1.options[<% = ix + 1 %>].text = '<% = ocd1.FItemList(ix).FCDName %>';
<% if cd1select = ix then %>
selectBox.cd1.selectedIndex = <%= ix+1 %>;
<% end if %>

<% next %>

selectBox.cd2.length = <% = ocd2.FResultCount + 1%>;
<% for ix=0 to ocd2.FResultCount - 1 %>
selectBox.cd2.options[<% = ix + 1 %>].value= '<% = ocd2.FItemList(ix).FCD1 %>,<% = ocd2.FItemList(ix).FCD2 %>' ;
selectBox.cd2.options[<% = ix + 1 %>].text = '<% = ocd2.FItemList(ix).FCDName %>';
selectBox.cd2.selectedIndex = 1 ;

<% next %>

selectBox.cd3.length = <% = ocd3.FResultCount + 1 %>;
<% for ix=0 to ocd3.FResultCount - 1 %>
selectBox.cd3.options[<% = ix + 1 %>].value= '<% = ocd3.FItemList(ix).FCD1 %>,<% = ocd3.FItemList(ix).FCD2 %>,<% = ocd3.FItemList(ix).FCD3 %>' ;
selectBox.cd3.options[<% = ix + 1 %>].text = '<% = ocd3.FItemList(ix).FCDName %>';
selectBox.cd3.selectedIndex = 1 ;

<% next %>


</script>
</head>
<body>
</body>
</html>
<%
set ocd1 = Nothing
set ocd2 = Nothing
set ocd3 = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
