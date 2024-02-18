<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycatePartnerCls.asp"-->

<%
	Response.CharSet = "euc-kr"

	Dim cDisp, i, vDepth, vCateCode, vGubun, vTempDepth, vIsThisLine
	vDepth			= requestCheckVar(Request("depth"),10)
	vCateCode 		= requestCheckVar(Request("cate"),128)
	vGubun			= requestCheckVar(Request("gubun"),32)

	SET cDisp = New cDispCate
	cDisp.FCurrPage = 1
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = vDepth
	cDisp.FRectCateCode = vCateCode
	cDisp.FRectUseYN = "Y"
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
				Response.Write "<option value="""&CHKIIF(i=0,"",Left(cDisp.FItemList(i).FCateCode,((cDisp.FItemList(i).FDepth-1)*3)))&""">-ÀüÃ¼-</option>" & vbCrLf
			End If
			
				Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(vIsThisLine="o","selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
			
			If i = cDisp.FResultCount-1 Then
				Response.Write "</select>"
			End If
			
			vTempDepth = cDisp.FItemList(i).FDepth
		Next
		Response.Write "<input type=""hidden"" name=""catecode_"&vGubun&""" value=""" & vCateCode & """>"
		response.write "<input type=""hidden"" name=""catecode_depth"" value="""&vTempDepth&""">"
	End If
	
	SET cDisp = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->