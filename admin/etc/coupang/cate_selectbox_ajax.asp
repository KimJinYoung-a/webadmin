<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/etc/coupang/coupangcls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim cdl
cdl = request("cdl")

If cdl <> "" Then
	Dim oCoupang, i
	Set oCoupang = new CCoupang
		oCoupang.FCurrPage	= 1
		oCoupang.FPageSize	= 50
		oCoupang.FRectCdl	= cdl
		oCoupang.getCateMiddleList

		If oCoupang.FResultCount > 0 Then
			Response.Write "<select id=""cdm"" name=""cdm"" class=""select"">" & vbCrLf
			Response.Write "<option value="""">2 Depth</option>" & vbCrLf
			For i=0 To oCoupang.FResultCount-1
				Response.Write "<option value=""" & oCoupang.FItemList(i).FCode_mid & """>" & oCoupang.FItemList(i).FCode_nm &"</option>"
			Next
			Response.Write "</select>"
		End If
	set oCoupang = Nothing
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->