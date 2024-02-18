<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/interviewCls.asp"-->
<%
dim idx, makerid, startdate, title, mainimg, isusing, sortNo, sqlStr, mode, adminid, comment, detailitemcnt, detailimg
dim detailimglink, existsbrandcnt
	idx 	= request("idx")
	makerid	= request("makerid")
	title 	= request("title")
	startdate 	= request("startdate")
	mainimg	= request("mainimg")
	detailimg 	= request("detailimg")
	detailimglink 	= request("detailimglink")	
	isusing	= request("isusing")
	sortNo 	= request("sortNo")
	comment 	= request("comment")
	mode 	= request("mode")
	menupos 	= request("menupos")
	
adminid = session("ssBctId")
detailitemcnt = 0
existsbrandcnt = 0

If mode = "I" Then
	sqlStr = "SELECT count(*) as cnt"
	sqlStr = sqlStr & " from db_user.dbo.tbl_user_c"
	sqlStr = sqlStr & " WHERE userid='"&makerid&"'"
	
	'response.write sqlStr & "<BR>"
	rsget.Open sqlStr, dbget, 1
    If Not rsget.Eof then
    	existsbrandcnt = rsget("cnt")
	End If
    rsget.Close

	If existsbrandcnt = 0 Then
		Response.Write  "<script language='javascript'>"
		Response.Write  "	alert('해당되는 브랜드가 없습니다.');"
		Response.Write  "	location.replace('/admin/brand/interview/interviewModify.asp?menupos="&menupos&"');"
		Response.Write  "</script>"
		dbget.close()	:	response.End
	End If

	if checkNotValidHTML(comment) or checkNotValidHTML(title) then
	%>
		<script>
			alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if
	
	sqlStr = "INSERT INTO db_brand.dbo.tbl_street_interview_main (" + vbcrlf
	sqlStr = sqlStr & " makerid, title, startdate, mainimg, detailimg, detailimglink, isusing, regadminid, lastadminid, comment)" + vbcrlf
	sqlStr = sqlStr & " 	select" + vbcrlf
	sqlStr = sqlStr & " 	c.userid as makerid, '"&html2db(title)&"', '"& left(startdate,10) &"', '"&mainimg&"', '"&detailimg&"'" + vbcrlf
	sqlStr = sqlStr & " 	, '"&html2db(detailimglink)&"', '"&isusing&"','"&adminid&"','"&adminid&"', '"&html2db(comment)&"'" + vbcrlf
	sqlStr = sqlStr & " 	from db_user.dbo.tbl_user_c c" + vbcrlf
	sqlStr = sqlStr & " 	where userid='"& makerid &"'"
		
	'response.write sqlStr & "<BR>"
	dbget.execute sqlStr

	sqlStr = "select IDENT_CURRENT('db_brand.dbo.tbl_street_interview_main') as idx"
	rsget.Open sqlStr, dbget, 1

	If Not rsget.Eof then
		idx = rsget("idx")
	End If
	rsget.close

	Response.Write "<script language='javascript'>alert('저장되었습니다');location.replace('/admin/brand/interview/interviewModify.asp?idx="&idx&"&menupos="&menupos&"');</script>"

ElseIf mode = "U" Then
	if idx="" then
	%>
		<script>
			alert('IDX가 없습니다.');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if
	

	if checkNotValidHTML(comment) or checkNotValidHTML(title) then
	%>
		<script>
			alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
			history.go(-1);
		</script>		
	<%		
		dbget.close()	:	response.End
	end if

	sqlStr = "UPDATE db_brand.dbo.tbl_street_interview_main SET" + VBCRLF
	sqlStr = sqlStr & " makerid = '"& makerid &"'" + VBCRLF
	sqlStr = sqlStr & " ,startdate = '"& left(startdate,10) &"'" + VBCRLF	
	sqlStr = sqlStr & " , title = '"& html2db(title) &"'" + VBCRLF
	sqlStr = sqlStr & " ,mainimg = '"&mainimg&"'" + VBCRLF
	sqlStr = sqlStr & " ,detailimg = '"&detailimg&"'" + VBCRLF
	sqlStr = sqlStr & " ,detailimglink = '"&html2db(detailimglink)&"'" + VBCRLF		
	sqlStr = sqlStr & " , isusing = '"&isusing&"'" + VBCRLF
	sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
	sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf
	sqlStr = sqlStr & " , comment = '"& html2db(comment) &"'" + VBCRLF
	sqlStr = sqlStr & " where mainidx ='" & Cstr(idx) & "'"

	'response.write sqlStr & "<BR>"	
	dbget.execute sqlStr

	response.write "<script language='javascript'>"
	response.write "	alert('저장되었습니다');"
	response.write "	location.replace('/admin/brand/interview/interviewModify.asp?idx="&idx&"&menupos="&menupos&"');"
	response.write "</script>"	

else
	Response.Write "<script language='javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
End If
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->