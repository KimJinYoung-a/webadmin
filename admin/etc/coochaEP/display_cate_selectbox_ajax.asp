<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	Response.CharSet = "euc-kr"

	Dim cDisp, i, vDepth, vCateCode, vGubun, vTempDepth, vIsThisLine
	vDepth			= Request("depth")
	vCateCode 		= Request("cate")
	vGubun			= Request("gubun")
	
	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = 3
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateList()

	If cDisp.FResultCount > 0 Then
		For i=0 To cDisp.FResultCount-1
			vIsThisLine = fnIsThisLine(cDisp.FItemList(i).FDepth,cDisp.FItemList(i).FCateCode,vCateCode)
			If i=0 Then
				vTempDepth = cDisp.FItemList(i).FDepth
			End If
			
			If i=0 OR vTempDepth <> cDisp.FItemList(i).FDepth Then
				If i <> 0 Then
					Response.Write "</select>&nbsp;" & vbCrLf
				End If
				Response.Write "<select name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value," & cDisp.FItemList(i).FDepth+1 & ",'" & vGubun & "');"">" & vbCrLf
				Response.Write "<option value="""&CHKIIF(i=0,"",Left(cDisp.FItemList(i).FCateCode,((cDisp.FItemList(i).FDepth-1)*3)))&""">-전체-</option>" & vbCrLf
			End If
			
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(vIsThisLine="o","selected","") & ">" & cDisp.FItemList(i).FCateName & chkIIF(cDisp.FItemList(i).FUseYN="N"," (사용안함)","") & "</option>"
			
			If i = cDisp.FResultCount-1 Then
				Response.Write "</select>"
			End If
			
			vTempDepth = cDisp.FItemList(i).FDepth
		Next
		Response.Write "<input type=""hidden"" name=""catecode_"&vGubun&""" value=""" & vCateCode & """>"
	End If
	
	SET cDisp = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->