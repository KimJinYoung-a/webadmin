<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
	Dim vIdx, vQuery, vAction, vDisp1, vType, vSubject, vItemID, vLinkURL, vSDate, vEDate, vUseYN, vSortNo, vRegdate, vImgURL, vOrderText
	vIdx = Request("idx")
	vAction = Request("action")
	
	If vAction = "del" Then
		vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate_menu_top] SET useyn = 'n' WHERE idx in(" & vIdx & ")"
		dbget.execute vQuery
		
		Response.Write "<script>alert('처리 완료!');parent.location.reload();</script>"
		dbget.close()
		Response.End
	End If
	
	If vIdx <> "" AND isNumeric(vIdx) = False Then
		Response.Write "<script>alert('잘못된 접근임!');window.close();</script>"
		dbget.close()
		Response.End
	End IF
	vDisp1 = Request("disp1")
	vType = Request("type")
	vSubject = html2db(Request("subject"))
	vItemID = Trim(Request("itemid"))
	vLinkURL = Request("linkurl")
	vSDate = Request("sdate")
	vEDate = Request("edate")
	vUseYN = Request("useyn")
	vSortNo = Trim(Request("sortno"))
	vImgURL = Trim(Request("imgurl"))
	vOrderText = html2db(Request("ordertext"))
	If isNumeric(vSortNo) = False Then
		Response.Write "<script>alert('정렬번호를 숫자로 입력하세요.');history.back();</script>"
		dbget.close()
		Response.End
	End IF
	
	If vIdx <> "" Then
		vQuery = "UPDATE [db_item].[dbo].[tbl_display_cate_menu_top] SET "
		vQuery = vQuery & " 	disp1 = '" & vDisp1 & "', "
		vQuery = vQuery & " 	type = '" & vType & "', "
		vQuery = vQuery & " 	subject = '" & vSubject & "', "
		vQuery = vQuery & " 	linkurl = '" & vLinkURL & "', "
		vQuery = vQuery & " 	itemid = '" & vItemID & "', "
		vQuery = vQuery & " 	imgurl = '" & vImgURL & "', "
		vQuery = vQuery & " 	sortno = '" & vSortNo & "', "
		vQuery = vQuery & " 	sdate = '" & vSDate & "', "
		vQuery = vQuery & " 	edate = '" & vEDate & "', "
		vQuery = vQuery & " 	useyn = '" & vUseYN & "', "
		vQuery = vQuery & " 	ordertext = '" & vOrderText & "', "
		vQuery = vQuery & " 	reguserid = '" & session("ssBctId") & "' "
		vQuery = vQuery & " WHERE idx = '" & vIdx & "'"
	Else
		vQuery = "INSERT INTO [db_item].[dbo].[tbl_display_cate_menu_top](disp1,type,subject,linkurl,itemid,imgurl,sortno,sdate,edate,useyn,ordertext,reguserid) "
		vQuery = vQuery & "VALUES('" & vDisp1 & "','" & vType & "','" & vSubject & "','" & vLinkURL & "','" & vItemID & "','" & vImgURL & "','" & vSortNo & "','" & vSDate & "','" & vEDate & "','" & vUseYN & "','" & vOrderText & "','" & session("ssBctId") & "')"
	End If
	dbget.execute vQuery
%>
<script>
alert("저장완료!");
opener.location.reload();
window.close();
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->