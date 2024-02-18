<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station 에디터 등록  
' Hieditor : 2008.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim editor_name,isusing,listimgName,list2imgName,regdate,editor_no,ticket_isusing
dim main1imgName,main2imgName,barnerimgName,barner2imgName,image_main_link , comment_isusing , mode
dim main3imgName, main4imgName, main5imgName, list2015imgName
	editor_name = html2db(request("editor_name"))
	isusing = requestCheckVar(request("isusing"),1)
	listimgName = request("listimgName")
	list2imgName = request("list2imgName")
	list2015imgName = request("list2015imgName")
	main1imgName = html2db(request("main1imgName"))
	main2imgName = html2db(request("main2imgName"))
	main3imgName = html2db(request("main3imgName"))
	main4imgName = html2db(request("main4imgName"))
	main5imgName = html2db(request("main5imgName"))	
	barnerimgName = request("barnerimgName")
	barner2imgName = request("barner2imgName")
	regdate = request("regdate") 					
	image_main_link = html2db(request("image_main_link"))
	editor_no = requestCheckVar(getNumeric(request("editor_no")),10)

	comment_isusing = requestCheckVar(request("comment_isusing"),10)
	mode = request("mode")
	if comment_isusing = "" then comment_isusing="OFF"
	
dim sql

'//신규저장
if mode = "add" then
	if editor_name <> "" and not(isnull(editor_name)) then
		editor_name = ReplaceBracket(editor_name)

		if checkNotValidHTML(editor_name) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('에디터명에는 HTML을 사용하실 수 없습니다.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	sql = "insert into db_culture_station.dbo.tbl_culturestation_editor" + vbcrlf
	sql = sql & " (comment_isusing, editor_name, isusing , image_list,image_list2,image_main,image_main2,image_barner,image_barner2,image_main_link,image_main3,image_main4,image_main5, image_list2015)" + vbcrlf
	sql = sql & " values ("  + vbcrlf
	sql = sql & " '"&comment_isusing&"'" + vbcrlf
	sql = sql & " ,'"& html2db(editor_name) &"'" + vbcrlf
	sql = sql & " ,'"&isusing&"'" + vbcrlf
	sql = sql & " ,'"&listimgName&"'" + vbcrlf
	sql = sql & " ,'"&list2imgName&"'" + vbcrlf
	sql = sql & " ,'"&main1imgName&"'" + vbcrlf
	sql = sql & " ,'"&main2imgName&"'" + vbcrlf
	sql = sql & " ,'"&barnerimgName&"'"	 + vbcrlf
	sql = sql & " ,'"&barner2imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&image_main_link&"'" + vbcrlf
	sql = sql & " ,'"&main3imgName&"'" + vbcrlf
	sql = sql & " ,'"&main4imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&main5imgName&"'"	 + vbcrlf
	sql = sql & " ,'"&list2015imgName&"'"	 + vbcrlf
	sql = sql & ")"
	'response.write sql
	dbget.execute sql

'//수정	
elseif mode = "edit" then
	if editor_name <> "" and not(isnull(editor_name)) then
		editor_name = ReplaceBracket(editor_name)
		if checkNotValidHTML(editor_name) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('에디터명에는 HTML을 사용하실 수 없습니다.');history.back();"
			response.write "</script>"
			response.End
		end if
	end If

	sql = "update db_culture_station.dbo.tbl_culturestation_editor set" + vbcrlf
	sql = sql & " comment_isusing='"&comment_isusing&"'" + vbcrlf
	sql = sql & " ,editor_name='"& html2db(editor_name) &"'" + vbcrlf
	sql = sql & " ,isusing='"&isusing&"'" + vbcrlf
	sql = sql & " ,image_list='"&listimgName&"'" + vbcrlf
	sql = sql & " ,image_list2='"&list2imgName&"'" + vbcrlf
	sql = sql & " ,image_main='"&main1imgName&"'" + vbcrlf
	sql = sql & " ,image_main2='"&main2imgName&"'" + vbcrlf
	sql = sql & " ,image_barner='"&barnerimgName&"'" + vbcrlf	
	sql = sql & " ,image_barner2='"&barner2imgName&"'" + vbcrlf	
	sql = sql & " ,image_main_link='"&image_main_link&"'" + vbcrlf
	sql = sql & " ,image_main3='"&main3imgName&"'" + vbcrlf
	sql = sql & " ,image_main4='"&main4imgName&"'" + vbcrlf
	sql = sql & " ,image_main5='"&main5imgName&"'" + vbcrlf
	sql = sql & " ,image_list2015='"&list2015imgName&"'" + vbcrlf
	sql = sql & " where editor_no = "&editor_no&"" + vbcrlf
	
	'response.write sql
	dbget.execute sql
end if	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<script type='text/javascript'>
	opener.location.reload();
	self.close();
</script>
