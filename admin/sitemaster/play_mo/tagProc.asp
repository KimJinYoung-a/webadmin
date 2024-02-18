<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description : 태그처리 페이지 play(공통)
' Hieditor : 2013-09-03 이종화 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, playcate
dim sqlStr
Dim tagname , tagurl , mode , subidx
Dim arrtagname , arrtagurl , arrtagnamecnt , arrtagurlcnt , i 

idx		= RequestCheckVar(request("idx"),10)
subidx		= RequestCheckVar(request("subidx"),10)
playcate = RequestCheckVar(request("playcate"),1)

tagname	= RequestCheckVar(request("tagname"),500)
tagurl	= RequestCheckVar(request("tagurl"),500)

mode = RequestCheckVar(request("mode"),10)

arrtagname = Split(tagname,",")
arrtagnamecnt = UBound(arrtagname)

arrtagurl = Split(tagurl,",")
arrtagurlcnt = UBound(arrtagurl)

if (mode = "tag") then

	If subidx = "" then
	    sqlStr = " delete from db_sitemaster.dbo.tbl_play_mo_tag where playcate = '"& playcate &"' and playidx = '"& idx &"'" & vbCrLf
	Else
	    sqlStr = " delete from db_sitemaster.dbo.tbl_play_mo_tag where playcate = '"& playcate &"' and playidx = '"& idx &"' and playidxsub = '"& subidx &"'" & vbCrLf
	End If 
   	'response.write sqlStr
	If arrtagnamecnt > 0 Then
		For i = 0 To arrtagnamecnt
			If Trim(arrtagname(i)) <> "" then
			sqlStr = sqlStr & " insert into db_sitemaster.dbo.tbl_play_mo_tag (playcate ,playidx , tagname , tagurl) values ( '"& playcate &"',"& idx &",'"&Trim(arrtagname(i))&"','"&Trim(arrtagurl(i))&"') " & vbCrLf
			End If 
		Next 
	Else
			sqlStr = sqlStr & " insert into db_sitemaster.dbo.tbl_play_mo_tag (playcate ,playidx , tagname , tagurl) values ( '"& playcate &"',"& idx &",'"&tagname&"','"&tagurl&"') " & vbCrLf
	End If 
    dbget.Execute sqlStr

End If 

dim referer
referer = request.ServerVariables("HTTP_REFERER")
If subidx = "" then
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play_mo/pop_tagReg.asp?idx=" + Cstr(idx) + "&playcate="+ CStr(playcate)+"&reload=on'</script>"
Else 
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.href='" & manageUrl & "/admin/sitemaster/play_mo/pop_tagReg.asp?idx=" + Cstr(idx) + "&playcate="+ CStr(playcate)+"&subidx="+ CStr(subidx) +"&reload=on'</script>"
End If 
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
