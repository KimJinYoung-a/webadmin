<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 이상구 생성
'			2012.05.23 한용민 이동/수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
dim i, userid, orderserial, divcd, contents_jupsu, id,contents_div , backwindow, mmGubun, qadiv
dim mode, sqlStr ,phoneNumber
userid          = RequestCheckVar(request("userid"),32)
orderserial     = RequestCheckVar(request("orderserial"),11)
mode            = RequestCheckVar(request("mode"),32)
contents_jupsu  = request("contents_jupsu")
backwindow      = RequestCheckVar(request("backwindow"),32)
id              = RequestCheckVar(request("id"),9)
contents_div    = RequestCheckVar(request("contents_div"),9)
divcd           = RequestCheckVar(request("divcd"),32)
mmGubun         = RequestCheckVar(request("mmGubun"),32)
phoneNumber     = RequestCheckVar(request("phoneNumber"),16)
qadiv           = RequestCheckVar(request("qadiv"),16)
    
dim referer
	referer = Request.Servervariables("HTTP_REFERER") 

if (divcd="") then divcd="1"

dim PreDivcd
'==============================================================================
if (mode = "write") then	'신규저장모드
	if contents_jupsu <> "" then
		if checkNotValidHTML(contents_jupsu) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
	end if

    if (divcd = "2") then			'요청메모
        sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, contents_jupsu, finishyn,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N',getdate()) "
        dbget.Execute sqlStr
        
        sqlStr = " select top 1 @@identity as id"
        rsget.Open sqlStr, dbget, 1
            IF Not rsget.EOF then
            id = rsget("id")
            end if
        rsget.close
   else 		'단순메모
        sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
        sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y',getdate(),getdate()) "
        dbget.Execute sqlStr
        
        sqlStr = " select top 1 @@identity as id"
        'response.write sqlStr &"<Br>"
        rsget.Open sqlStr, dbget, 1
            IF Not rsget.EOF then
            id = rsget("id")
            end if
        rsget.close
    end if

    response.write "<script type='text/javascript'>alert('등록되었습니다.'); </script>"

	if InStr(referer,"cscenter_memo.asp")>0 then
		response.write "<script type='text/javascript'>"
		'response.write "	location.replace('" + replace(referer,"&id=","") + "&id=" & id & "');"
		response.write "	opener.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"';"
		response.write "	self.close();"
		response.write "</script>"
    end if

    'dbget.close()	:	response.End
elseif (mode = "modify") then		'수정모드
	if contents_jupsu <> "" then
		if checkNotValidHTML(contents_jupsu) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

    ''기존 요청구분이 처리 요청으로 변경된경우 완료 처리를 NULL로
    sqlStr = " select top 1 * from [db_cs].[dbo].tbl_cs_memo"
    sqlStr = sqlStr + " where id = " + CStr(id) + " "
    'response.write sqlStr &"<Br>"
    rsget.Open sqlStr, dbget, 1
    IF Not rsget.Eof then
        PreDivcd = rsget("divcd")
    End IF
    rsget.Close
    
    if (PreDivcd="1") and (divcd="2") then
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " set divcd = '" + CStr(divcd) + "'"
        sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , phoneNumber = '" + CStr(phoneNumber) + "'"
        sqlStr = sqlStr + " , userid = '" + CStr(userid) + "'"
        sqlStr = sqlStr + " , orderserial = '" + CStr(orderserial) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " ,finishyn ='N'"
        sqlStr = sqlStr + " ,finishdate =NULL"
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        
        dbget.Execute sqlStr
    elseif (PreDivcd="2") and (divcd="1") then
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " set divcd = '" + CStr(divcd) + "'"
        sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , phoneNumber = '" + CStr(phoneNumber) + "'"
        sqlStr = sqlStr + " , userid = '" + CStr(userid) + "'"
        sqlStr = sqlStr + " , orderserial = '" + CStr(orderserial) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " ,finishyn ='Y'"
        sqlStr = sqlStr + " ,finishdate =getdate()"
        sqlStr = sqlStr + " ,finishuser='" + session("ssBctId") + "'"
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        
        dbget.Execute sqlStr
    else
    
        sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " set mmgubun = '" + CStr(mmgubun) + "'"
        sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
        sqlStr = sqlStr + " , phoneNumber = '" + CStr(phoneNumber) + "'"
        sqlStr = sqlStr + " , userid = '" + CStr(userid) + "'"
        sqlStr = sqlStr + " , orderserial = '" + CStr(orderserial) + "'"
        sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr
    end if
	'response.write sqlStr&"<br>"
    response.write "<script type='text/javascript'>alert('수정되었습니다.'); </script>"

	if InStr(referer,"cscenter_memo.asp")>0 then
		response.write "<script>"
		'response.write "	location.replace('" + referer + "');"
		response.write "	opener.parent.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"';"
		response.write "	self.close();"
		response.write "</script>"
    end if

    'dbget.close()	:	response.End
elseif (mode = "finish") then
	if contents_jupsu <> "" then
		if checkNotValidHTML(contents_jupsu) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

    sqlStr = " update [db_cs].[dbo].tbl_cs_memo "
    sqlStr = sqlStr + " set finishyn = 'Y'"
    sqlStr = sqlStr + " , finishuser = '" + session("ssBctId") + "'"
    sqlStr = sqlStr + " , finishdate = getdate() "
    sqlStr = sqlStr + " , mmgubun = '" + CStr(mmgubun) + "'"
    sqlStr = sqlStr + " , qadiv = '" + CStr(qadiv) + "'"
    sqlStr = sqlStr + " , contents_jupsu = '" + CStr(html2db(contents_jupsu)) + "' "
    sqlStr = sqlStr + " where id = '" &id&"'"
    'response.write sqlstr
    dbget.Execute sqlStr
    
    response.write "<script type='text/javascript'>alert('완료되었습니다.'); </script>"

	if InStr(referer,"cscenter_memo.asp")>0 then
		response.write "<script type='text/javascript'>"
		response.write "	location.replace('" + referer + "');"
		response.write "	opener.parent.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"';"	
		response.write "	self.close();"
		response.write "</script>"
    end if

    'dbget.close()	:	response.End
elseif (mode = "delete") then
    sqlStr = " delete from [db_cs].[dbo].tbl_cs_memo "
    sqlStr = sqlStr + " where id = " + CStr(id) + " "
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('삭제되었습니다.'); </script>"

	if InStr(referer,"cscenter_memo.asp")>0 then
		response.write "<script type='text/javascript'>"
		response.write "	self.close();"
		response.write "	opener.parent.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"';"		
		response.write "</script>"
    end if

    'dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->