<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/optionmanagecls.asp"-->

<%
dim ix
dim search_code,form_name,element_name

search_code = request("search_code")
form_name = request("form_name")
element_name = request("element_name")

dim ooption
set ooption = new COptionManager
ooption.FRectOnlyUsing = "on"
ooption.FRectOrderType= "d"

if element_name = "opt1" then
ooption.GetOption01Select
end if

if element_name = "opt2" then
ooption.GetOption02Select search_code
end if

%>
<script language="JavaScript">
<!--

var selectBox = parent.<% = form_name %>.<% = element_name %> ;

selectBox.length = <% = ooption.FResultCount %> + 1;

<% for ix=0 to ooption.FResultCount - 1 %>
<% if element_name = "opt1" then %>
selectBox.options[<% = ix + 1 %>].value= '<% = ooption.FItemList(ix).FCode01 %>' ;
<% elseif element_name = "opt2" then %>
selectBox.options[<% = ix + 1 %>].value= '<% = ooption.FItemList(ix).FCode01 %><% = ooption.FItemList(ix).FCode02 %>' ;
<% end if %>
selectBox.options[<% = ix + 1 %>].text = '<% = ooption.FItemList(ix).FCodeName %>';

<% next %>

//-->
</script>

<%
set ooption = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->