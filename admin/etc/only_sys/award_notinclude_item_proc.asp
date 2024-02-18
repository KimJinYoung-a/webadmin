<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- # include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->
<%
	Dim vQuery, vGubun, vItemID, vTemp, i
	vGubun = Request("gubun")
	vTemp = Trim(Request("itemid"))
	vTemp = Replace(vTemp," ","")
	
	If vGubun = "insert" Then
		vQuery = "SELECT count(itemid) FROM [db_const].[dbo].[tbl_const_award_NotInclude_Item] WHERE itemid in(" & vTemp & ")"
		rsget.Open vQuery,dbget
		if rsget(0) > 0 Then
			Response.Write "<script>alert('이미 제외된 상품이 있습니다.\n상품코드를 다시 확인해주세요.');history.back();</script>"
			rsget.Close
			dbget.close()
			Response.End
		else
			rsget.Close
		end if
		
		For i = LBound(Split(vTemp,",")) To UBound(Split(vTemp,","))
			vItemID = Split(vTemp,",")(i)
			vQuery = "INSERT INTO [db_const].[dbo].[tbl_const_award_NotInclude_Item](itemid) VALUES('" & vItemID & "')"
			dbget.Execute vQuery
		Next
	ElseIf vGubun = "delete" Then
		vItemID = Replace(vTemp,",","")
		vQuery = "DELETE [db_const].[dbo].[tbl_const_award_NotInclude_Item] WHERE itemid = '" & vItemID & "'"
		dbget.Execute vQuery
	End If
%>

<script>
location.href = "award_notinclude_item.asp";
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->