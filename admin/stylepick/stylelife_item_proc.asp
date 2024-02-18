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
	Dim i, j, vGubun, vQuery, vTemp, vItemID, vStyleCate, vItemCate, vStyleCate2, vReturnURL
	vGubun 		= requestCheckVar(Request("gubun"),15)
	vItemID		= Request("itemid")
	vStyleCate 	= Request("stylecate")
	vItemCate	= Request("itemcate")
	vReturnURL	= Request("returnUrl")
	
	If vGubun = "" OR vItemID = "" Then
		Response.Write "<script>alert('잘못된 접근입니다.');</script>"
		dbget.close()
		Response.End
	Else
		
		If vGubun = "oneitemchange" Then
			vTemp		= vItemID
			vItemID 	= Split(vTemp,"|")(0)
			vItemCate	= Split(vTemp,"|")(1)
			vStyleCate 	= Split(vTemp,"|")(2)
			vStyleCate2	= fnStyleCate2(vItemCate)
			
			vQuery = "INSERT INTO [db_giftplus].[dbo].[tbl_stylepick_item](itemid,cd1,cd2,cd3) VALUES('" & vItemID & "','" & vStyleCate & "','" & vStyleCate2 & "','')"
			dbget.execute vQuery
			
		ElseIf vGubun = "default" Then
			vQuery = "DELETE [db_giftplus].[dbo].[tbl_stylepick_item] WHERE itemid IN(" & Trim(vItemID) & ")"
			dbget.execute vQuery
			
		ElseIf vGubun = "setstyle" Then
 			For j = LBound(Split(vItemID,",")) To UBound(Split(vItemID,","))
				For i = LBound(Split(vStyleCate,",")) To UBound(Split(vStyleCate,","))
					vQuery = vQuery & "EXECUTE [db_giftplus].[dbo].[ten_stylelife_stylecate_insert] '" & Trim(Split(vItemID,",")(j)) & "', '" & Trim(Split(vStyleCate,",")(i)) & "', '" & fnStyleCate2(Split(vItemCate,",")(j)) & "'" & vbCrLf
				Next
			Next
			dbget.execute vQuery
			
		ElseIf vGubun = "notuseitem" Then
 			For j = LBound(Split(vItemID,",")) To UBound(Split(vItemID,","))
				vQuery = vQuery & " INSERT INTO [db_giftplus].[dbo].[tbl_stylelife_notuse_item](itemid) VALUES('" & Trim(Split(vItemID,",")(j)) & "') " & vbCrLf
			Next
			dbget.execute vQuery
			
		End IF
		
	End IF
%>

<script language="javascript">
document.location.href = "<%=vReturnURL%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->