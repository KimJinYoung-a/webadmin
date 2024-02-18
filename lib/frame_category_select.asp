<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual ="/lib/classes/items/category_selectcls.asp" -->
<%
dim ocate,ix
dim cd1,cd2
dim search_code
dim form_name
dim element_name

search_code = request("search_code")
form_name = request("form_name")
element_name = request("element_name")


search_code = split(search_code,",")

if element_name = "cd2" then
cd1 = search_code(0)
elseif element_name = "cd3" then
cd1 = search_code(0)
cd2 = search_code(1)
end if



set ocate = New CCategory
ocate.FRectCD1 = cd1
ocate.FRectCD2 = cd2
if element_name = "cd1" then
ocate.CategoryCodeLarge
elseif element_name = "cd2" then
ocate.CategoryCodeMid
elseif element_name = "cd3" then
ocate.CategoryCodeSmall
end if

%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html">
<script>
var selectBox = parent.<% = form_name %>.<% = element_name %> ;

selectBox.length = <% = ocate.FResultCount %> + 1;

<% for ix=0 to ocate.FResultCount - 1 %>
<% if element_name = "cd1" then %>
selectBox.options[<% = ix + 1 %>].value= '<% = ocate.FItemList(ix).FCD1 %>' ;
<% elseif element_name = "cd2" then %>
selectBox.options[<% = ix + 1 %>].value= '<% = ocate.FItemList(ix).FCD1 %>,<% = ocate.FItemList(ix).FCD2 %>' ;
<% elseif element_name = "cd3" then %>
selectBox.options[<% = ix + 1 %>].value= '<% = ocate.FItemList(ix).FCD1 %>,<% = ocate.FItemList(ix).FCD2 %>,<% = ocate.FItemList(ix).FCD3 %>' ;
<% end if %>
selectBox.options[<% = ix + 1 %>].text = '<% = ocate.FItemList(ix).FCDName %>';

<% next %>
</script>
</head>
<body>
</body>
</html>
<%
set ocate = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
