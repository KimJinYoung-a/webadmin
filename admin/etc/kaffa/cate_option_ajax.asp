<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/kaffa/kaffaCls.asp"-->
<%
	Response.CharSet = "euc-kr"

	Dim vCateGubun, vTenCode, cKaffa, vQuery, vBody, vCate1, vCate2, vCate3, arrList, intLoop, vTotCount
	vCateGubun = Request("categubun")
	vTenCode = Split(Request("area"),"-")(1)
	vCate1 = Request("cate1")
	vCate2 = Request("cate2")
	vCate3 = Request("cate3")

	vBody = ""
	set cKaffa = new cKaffaItem
	If vCateGubun = "1" Then
	ElseIf vCateGubun = "2" Then
		cKaffa.FRectCate1 = vCate1
		arrList = cKaffa.GetKaffaCate2List
		vTotCount = cKaffa.FTotalCount
		If vTotCount = 0 then
			vBody = "x"
		else
			vBody = vBody & "<select name=""kaffacate2-"&vTenCode&""" onChange=""goCate3('cate3-"&vTenCode&"',$('select[name=kaffacate1-"&vTenCode&"]').val(),this.value);"">" & vbCrLf
			vBody = vBody & "<option value=""x"">-</option>" & vbCrLf
			For intLoop =0 To UBound(arrList,2)
				vBody = vBody & "<option value=""" & arrList(1,intLoop) & """>" & arrList(2,intLoop) & "</option>" & vbCrLf
			Next
			vBody = vBody & "</select>" & vbCrLf
		end if
	ElseIf vCateGubun = "3" Then
		cKaffa.FRectCate1 = vCate1
		cKaffa.FRectCate2 = vCate2
		arrList = cKaffa.GetKaffaCate3List
		vTotCount = cKaffa.FTotalCount
		If vTotCount = 0 then
			vBody = "x"
		else
			vBody = vBody & "<select name=""kaffacate3-"&vTenCode&""">" & vbCrLf
			vBody = vBody & "<option value=""x"">-</option>" & vbCrLf
			For intLoop =0 To UBound(arrList,2)
				vBody = vBody & "<option value=""" & arrList(2,intLoop) & """>" & arrList(3,intLoop) & "</option>" & vbCrLf
			Next
			vBody = vBody & "</select>" & vbCrLf
		end if
	End If
	set cKaffa = nothing

	Response.Write vBody
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->