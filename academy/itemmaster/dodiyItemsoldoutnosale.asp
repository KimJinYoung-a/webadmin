<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim i, vQuery, vAction, vAllItemID, vItemID
	vAction		= RequestCheckvar(Request("action"),16)
	vAllItemID	= Request("allitemid")
	
	If vAllItemID = "" Then
		dbACADEMYget.close()
		Response.End
	End IF
  	if vAllItemID <> "" then
		if checkNotValidHTML(vAllItemID) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
	vAllItemID = Replace(Trim(vAllItemID)," ","")
	If Right(vAllItemID,1) = "," Then
		vAllItemID = Left(vAllItemID,(Len(vAllItemID)-1))
	End IF

	For i = 0 To UBound(Split(vAllItemID,","))
		vItemID = Split(vAllItemID,",")(i)
		If vAction = "soldout" Then
			vQuery = "update [db_academy].[dbo].[tbl_diy_item] set sellyn='S' WHERE itemid = '" & vItemID & "'"  & vbCrLf
			dbACADEMYget.execute vQuery
		Else
			vQuery = "update [db_academy].[dbo].[tbl_diy_item] set sellyn='N' WHERE itemid = '" & vItemID & "'"  & vbCrLf
			dbACADEMYget.execute vQuery
		End If
	Next
%>
<script>parent.fnSellYNIsusingEditEnd();</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->