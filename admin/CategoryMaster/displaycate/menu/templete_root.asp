<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->
<%
	Dim cDisp, vCateCode, vArrList, vWWW, vServerGubun
	vCateCode = Request("catecode")
	SET cDisp = New cDispCateMenu
	cDisp.FRectCateCode = vCateCode
	vArrList = cDisp.GetDispCateMenuList
	SET cDisp = Nothing
	
	
	Dim intLoop
	Dim cate(58), catecode(58), vItemID, vImgLink, vIsUseItemImg
	
	IF isArray(vArrList) THEN
	
		For intLoop = 0 To UBound(vArrList,2)
			If vArrList(0,intLoop) = "bookitemid" Then
				vItemID = vArrList(2,intLoop)
			ElseIf vArrList(0,intLoop) = "bookimg" Then
				vImgLink = vArrList(2,intLoop)
			Else
				catecode(intLoop) = vArrList(2,intLoop)
				cate(intLoop) = vArrList(3,intLoop)
			End If
		Next
	End If
%>

<% If vCateCode = "101" Then		'### 디자인문구 	%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete101.asp"-->
<% ElseIf vCateCode = "102" Then	'### 핸드폰/디지털 	%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete102.asp"-->
<% ElseIf vCateCode = "103" Then	'### 캠핑/트래블 	%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete103.asp"-->
<% ElseIf vCateCode = "104" Then	'### 토이 			%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete104.asp"-->
<% ElseIf vCateCode = "105" Then	'### 그래픽 		%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete105.asp"-->
<% ElseIf vCateCode = "106" Then	'### 홈인테리어 	%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete106.asp"-->
<% ElseIf vCateCode = "107" Then	'### 키친/푸드 		%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete107.asp"-->
<% ElseIf vCateCode = "108" Then	'### 패션/뷰티 		%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete108.asp"-->
<% ElseIf vCateCode = "109" Then	'### 베이비 		%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete109.asp"-->
<% ElseIf vCateCode = "110" Then	'### CAT&DOG 		%>
	<!-- #include virtual="/admin/CategoryMaster/displaycate/menu/templete110.asp"-->
<% End If %>

<!-- #include virtual="/lib/db/dbclose.asp" -->