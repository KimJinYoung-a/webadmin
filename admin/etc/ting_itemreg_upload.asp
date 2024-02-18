<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<%

Dim iting
Dim itemid,propcost,propmon

itemid = request("itemid")
propcost = request("propcost")
propmon = request("yyyy1") & "-" & request("mm1")

set iting = new CWaitItemUpload
iting.WaitProductUpload itemid,propcost,propmon

%>
<%
set iting = Nothing
%>
<script language="JavaScript">
<!--
alert("데이터를 저장하였습니다.");
location.replace("ting_wait_itemreg.asp");
//-->
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->