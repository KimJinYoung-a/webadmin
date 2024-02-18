<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->

<%
	Dim vQuery, vIdx, vMakerId, vCateCode, vListtitleimgName, vListmainimgName, vListspareimgName, vIsUsing, vListText, vContentHtml, vSorting, vContentTitleName
	vIdx 				= request("idx")
	vMakerId			= request("makerid")
	vCateCode			= request("cate")
	vListtitleimgName	= request("listtitleimgName")
	If vListtitleimgName <> "" Then
		If Left(vListtitleimgName,7) <> "http://" Then
			vListtitleimgName = "http://webimage.10x10.co.kr/diary_collection/2012/listtitleimg/" & vListtitleimgName
		End If
	End If
	vListmainimgName	= request("listmainimgName")
	If vListmainimgName <> "" Then
		If Left(vListmainimgName,7) <> "http://" Then
			vListmainimgName = "http://webimage.10x10.co.kr/diary_collection/2012/listmainimg/" & vListmainimgName
		End If
	End If
	vListspareimgName	= request("listspareimgName")
	If vListspareimgName <> "" Then
		If Left(vListspareimgName,7) <> "http://" Then
			vListspareimgName = "http://webimage.10x10.co.kr/diary_collection/2012/listspareimg/" & vListspareimgName
		End If
	End If
	vContentTitleName	= request("contenttitleName")
	If vContentTitleName <> "" Then
		If Left(vContentTitleName,7) <> "http://" Then
			vContentTitleName = "http://webimage.10x10.co.kr/diary_collection/2012/contenttitle/" & vContentTitleName
		End If
	End If
	vIsUsing			= request("isusing")
	vListText			= html2db(request("list_text"))
	vContentHtml		= html2db(request("content_html"))
	vSorting			= request("sorting")
	
	If vIdx <> "" Then
		vQuery = "UPDATE [db_diary2010].[dbo].[tbl_diary_brandstory_2012] set "
		vQuery = vQuery & " makerid = '" & vMakerId & "', "
		vQuery = vQuery & " cate = '" & vCateCode & "', "
		vQuery = vQuery & " list_titleimg = '" & vListtitleimgName & "', "
		vQuery = vQuery & " list_mainimg = '" & vListmainimgName & "', "
		vQuery = vQuery & " list_spareimg = '" & vListspareimgName & "', "
		vQuery = vQuery & " isusing = '" & vIsUsing & "', "
		vQuery = vQuery & " list_text = '" & vListText & "', "
		vQuery = vQuery & " content_title = '" & vContentTitleName & "', "
		vQuery = vQuery & " content_html = '" & vContentHtml & "', "
		vQuery = vQuery & " sorting = '" & vSorting & "' "
		vQuery = vQuery & "WHERE idx = '" & vIdx & "' "
		
		dbget.execute vQuery
	Else
		vQuery = "INSERT INTO [db_diary2010].[dbo].[tbl_diary_brandstory_2012](makerid, cate, list_titleimg, list_mainimg, list_spareimg, isusing, list_text, content_title, content_html, sorting) VALUES"
		vQuery = vQuery & "('" & vMakerId & "', '" & vCateCode & "', '" & vListtitleimgName & "', '" & vListmainimgName & "', '" & vListspareimgName & "', "
		vQuery = vQuery & "'" & vIsUsing & "', '" & vListText & "', '" & vContentTitleName & "', '" & vContentHtml & "', '" & vSorting & "') "
		
		dbget.execute vQuery
	End If
	'rw vQuery
%>

<script type="text/javascript">
alert("저장되었습니다.");
opener.location.reload();
window.close();

</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->