<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylepick_cls.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
	Dim i, j, vGubun, vQuery, vTemp, vItemID, vReturnURL
	vGubun 		= requestCheckVar(Request("gubun"),15)
	vItemID		= Request("itemid")
	vReturnURL	= Request("returnUrl")
	
	Dim cdl, cdm, cds, vCD1, vCD2, vCD3, vCate3
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	vCD1 = request("cd1")
	vCD2 = request("cd2")
	vCD3 = request("cd3")
	vCate3 = Request("cate3")
	
	
	If (vGubun = "" AND vItemID = "") OR vCate3 = "" Then
		Response.Write "<script>alert('잘못된 접근입니다.');</script>"
		dbget.close()
		Response.End
	Else
		vQuery = "EXECUTE [db_giftplus].[dbo].[ten_stylelife_stylecate_mid_update] '" & vGubun & "','" & vItemID & "','" & cdl & "','" & cdm & "','" & cds & "','" & vCD1 & "','" & vCD2 & "','" & vCD3 & "','" & vCate3 & "'"
		'Response.Write vQuery
		'dbget.close()
		'Response.End
		dbget.execute vQuery
	End IF
%>

<script language="javascript">
document.location.href = "<%=vReturnURL%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->