<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트
' History : 2014.03.19 한용민 생성
' History : 2014.10.31 유태욱 mtitle 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftday_cls.asp"-->

<%
dim masteridx, title, mtitle, startdate, enddate, listtopimg_w, listtopimg_m, regtopimg_w, regtopimg_m, mainimg_W, regdate, isusing, adminid, mode, sqlStr, menupos
	masteridx 	= requestcheckvar(request("masteridx"),10)
	title	= requestcheckvar(request("title"),128)
	mtitle	= requestcheckvar(request("mtitle"),128)
	startdate 	= requestcheckvar(request("startdate"),10)
	enddate 	= requestcheckvar(request("enddate"),10)
	listtopimg_w	= request("listtopimg_w")
	listtopimg_m	= request("listtopimg_m")	
	regtopimg_w	= request("regtopimg_w")
	regtopimg_m	= request("regtopimg_m")
	mainimg_W	= request("mainimg_W")
	isusing 	= requestcheckvar(request("isusing"),1)
	mode 	= requestcheckvar(request("mode"),32)
	menupos 	= requestcheckvar(request("menupos"),10)

adminid = session("ssBctId")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
if InStr(refer,"10x10.co.kr")<1 then
	Response.Write "잘못된 접속입니다."
	dbget.close()	:	response.End
end if

If mode = "dayedit" Then
	if masteridx<>"" then
		if masteridx="" or title="" or mtitle="" or startdate="" or enddate="" or isusing="" then
			Response.Write "<script type='text/javascript'>alert('미 입력된 내용이 있습니다.'); history.back(-1);</script>"
			dbget.close()	:	response.End			
		end if

		if title <> "" and not(isnull(title)) then
			title = ReplaceBracket(title)
		end If
		if mtitle <> "" and not(isnull(mtitle)) then
			mtitle = ReplaceBracket(mtitle)
		end If

		if checkNotValidHTML(title) then
		%>
			<script type='text/javascript'>
				alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
				history.go(-1);
			</script>		
		<%
			dbget.close()	:	response.End
		end if

		if checkNotValidHTML(mtitle) then
		%>
			<script type='text/javascript'>
				alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
				history.go(-1);
			</script>		
		<%
			dbget.close()	:	response.End
		end if

		sqlStr = "UPDATE db_board.dbo.tbl_giftday_master" + VBCRLF
		sqlStr = sqlStr & " SET title = '"&trim(html2db(title))&"'" + VBCRLF
		sqlStr = sqlStr & " ,mtitle = '"&trim(html2db(mtitle))&"'" + VBCRLF
		sqlStr = sqlStr & " ,startdate = '"&trim(startdate)&" 00:00:00'" + VBCRLF
		sqlStr = sqlStr & " ,enddate = '"&trim(enddate)&" 23:59:59'" + VBCRLF
		sqlStr = sqlStr & " ,listtopimg_w = '"&trim(listtopimg_w)&"'" + VBCRLF
		sqlStr = sqlStr & " ,listtopimg_m = '"&trim(listtopimg_m)&"'" + VBCRLF
		sqlStr = sqlStr & " ,regtopimg_w = '"&trim(regtopimg_w)&"'" + VBCRLF
		sqlStr = sqlStr & " ,regtopimg_m = '"&trim(regtopimg_m)&"'" + VBCRLF
		sqlStr = sqlStr & " ,mainimg_W = '"&trim(mainimg_W)&"'" + VBCRLF
		sqlStr = sqlStr & " ,isusing = '"&trim(isusing)&"'" + VBCRLF
		sqlStr = sqlStr & " where masteridx ='" & Cstr(masteridx) & "'"
	
		'response.write sqlStr & "<BR>"	
		dbget.execute sqlStr

		Response.Write "<script type='text/javascript'>alert('저장되었습니다.'); opener.location.reload(); self.close();</script>"

	else
		if title="" or mtitle="" or startdate="" or enddate="" or isusing="" then
			Response.Write "<script type='text/javascript'>alert('미 입력된 내용이 있습니다.'); history.back(-1);</script>"
			dbget.close()	:	response.End			
		end if

		if title <> "" and not(isnull(title)) then
			title = ReplaceBracket(title)
		end If
		if mtitle <> "" and not(isnull(mtitle)) then
			mtitle = ReplaceBracket(mtitle)
		end If

		if checkNotValidHTML(title) then
		%>
			<script type='text/javascript'>
				alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
				history.go(-1);
			</script>		
		<%
			dbget.close()	:	response.End
		end if
		
		if checkNotValidHTML(mtitle) then
		%>
			<script type='text/javascript'>
				alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');
				history.go(-1);
			</script>		
		<%
			dbget.close()	:	response.End
		end if

		sqlStr = "INSERT INTO db_board.dbo.tbl_giftday_master (" + vbcrlf
		sqlStr = sqlStr & " title, mtitle, startdate, enddate, listtopimg_w, listtopimg_m, regtopimg_w, regtopimg_m, mainimg_W, regdate, isusing) values (" + vbcrlf
		sqlStr = sqlStr & " '"&trim(html2db(title))&"', '"&trim(html2db(mtitle))&"', '"&trim(startdate)&" 00:00:00', '"&trim(enddate)&" 23:59:59', '"&trim(listtopimg_w)&"', '"&trim(listtopimg_m)&"'" + vbcrlf
		sqlStr = sqlStr & " , '"&trim(regtopimg_w)&"', '"&trim(regtopimg_m)&"', '"&trim(mainimg_W)&"',getdate(), '"&trim(isusing)&"'" + vbcrlf
		sqlStr = sqlStr & " )"
	
		'response.write sqlStr & "<BR>"
		dbget.execute sqlStr
	
		Response.Write "<script type='text/javascript'>alert('저장되었습니다.'); opener.location.reload(); self.close();</script>"
	end if

else
	Response.Write "<script type='text/javascript'>alert('구분자가 없습니다.'); history.back(-1);</script>"
	dbget.close()	:	response.End
End If
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->