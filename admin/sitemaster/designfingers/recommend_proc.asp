<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->

<%
	Dim vIdx, vQuery, vGubun, vUseYN
	vGubun	= Request("gubun")
	vIdx	= Request("idx")
	vUseYN 	= Request("useyn")
	
	IF vGubun = "d" Then
		vQuery = "UPDATE [db_sitemaster].[dbo].[tbl_designfingers_recommend] SET USEYN = '" & vUseYN & "' WHERE IDX = '" & vIdx & "' "
		dbget.execute vQuery
	End IF
%>

<script language="javascript">
alert("처리되었습니다.");
location.href = "/admin/sitemaster/designfingers/recommend_list.asp";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->