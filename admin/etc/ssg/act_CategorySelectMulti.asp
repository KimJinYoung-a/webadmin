<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/etc/ssg/ssgItemcls.asp"-->
<%
Response.CharSet = "euc-kr"

Dim vDepth, vCateCode
vDepth = requestCheckVar(Request("depth"),1)
vCateCode = requestCheckVar(Request("cate"),3)

Dim oSsg, i
Set oSsg = new Cssg
	oSsg.FCurrPage	= 1
	oSsg.FPageSize	= 50
	oSsg.getCateLargeList

	If oSsg.FResultCount > 0 Then
		Response.Write "<select id=""cate"" name=""cate"" class=""select"" onChange=""jsCateCodeSelectBox(this.value,2);"" >" & vbCrLf
		Response.Write "<option value="""">1 Depth</option>" & vbCrLf
		For i=0 To oSsg.FResultCount-1
			If CStr(vCateCode) = CStr(oSsg.FItemList(i).FCode_large) Then
				Response.Write "<option selected value=""" & oSsg.FItemList(i).FCode_large & """>" & oSsg.FItemList(i).FCode_nm &"</option>"
			Else
				Response.Write "<option value=""" & oSsg.FItemList(i).FCode_large & """>" & oSsg.FItemList(i).FCode_nm &"</option>"
			End If
		Next
		Response.Write "</select>"
	End If
set oSsg = Nothing

Set oSsg = new Cssg
	oSsg.FCurrPage	= 1
	oSsg.FPageSize	= 50
	oSsg.FRectDepth = vDepth
	oSsg.FRectCateCode = vCateCode
	oSsg.GetMngCateList()

	If oSsg.FResultCount > 0 Then
		Response.Write "<select id=""cate"&vDepth&""" name=""cate"&vDepth&""" class=""select"" >" & vbCrLf
		Response.Write "<option value="""">2 Depth</option>" & vbCrLf
		For i=0 To oSsg.FResultCount-1
			Response.Write "<option value=""" & oSsg.FItemList(i).FCode_mid & """>" & oSsg.FItemList(i).FCode_nm &"</option>"
		Next
		Response.Write "</select>"
	End If
set oSsg = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->