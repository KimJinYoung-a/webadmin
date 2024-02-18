<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
dim referer, retUrl
referer = Request.Servervariables("HTTP_REFERER")
if InStr(referer,"popCallRing.asp")>0 then
   retUrl = "popCallRing.asp"
else
   retUrl = "CallRingWithOrderFrame.asp"
end if

dim i, userid, orderserial, divcd, contents_jupsu, id,contents_div , backwindow, mmGubun, qadiv
dim mode, sqlStr
dim phoneNumber
dim yyyy1, mm1, dd1, hh1, retrydate
dim specialmemo

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

yyyy1           = RequestCheckVar(request("yyyy1"),16)
mm1           	= RequestCheckVar(request("mm1"),16)
dd1           	= RequestCheckVar(request("dd1"),16)
hh1           	= RequestCheckVar(request("hh1"),16)

specialmemo    	= RequestCheckVar(request("specialmemo"),8)

if (yyyy1 <> "") and (mm1 <> "") and (dd1 <> "") and (hh1 <> "") then
	retrydate = "'" & yyyy1 & "-" & mm1 & "-" & dd1 & " " & hh1 & ":00:00" & "'"
else
	retrydate = "NULL"
end if

''response.write "retrydate=" & retrydate
''response.end
''response.write "divcd=" & divcd

if (divcd="") then divcd="1"

dim PreDivcd
'==============================================================================
if (mode = "write") then	'신규저장모드

        if (divcd = "2") then			'요청메모
            sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, contents_jupsu, finishyn,regdate, retrydate, specialmemo) "
            sqlStr = sqlStr + " values('" + CStr(orderserial) + "','2','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','N',getdate(), " + CStr(retrydate) + ", '" + CStr(specialmemo) + "') "
            dbget.Execute sqlStr

            sqlStr = " select top 1 @@identity as id"
            rsget.Open sqlStr, dbget, 1
                IF Not rsget.EOF then
                id = rsget("id")
                end if
            rsget.close
       else 		'단순메모
            sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate, specialmemo) "
            sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','" + mmGubun + "','" + qadiv + "','" + phoneNumber + "','" + session("ssBctId") + "','" + session("ssBctId") + "','" + html2db(contents_jupsu) + "','Y',getdate(),getdate(), '" + CStr(specialmemo) + "') "
            dbget.Execute sqlStr

            sqlStr = " select top 1 @@identity as id"
            rsget.Open sqlStr, dbget, 1
                IF Not rsget.EOF then
                id = rsget("id")
                end if
            rsget.close
        end if

        response.write "<script>alert('등록되었습니다.'); </script>"
        response.write "<script>location.replace('" + retUrl + "?id=" & id & "'); </script>"
        'dbget.close()	:	response.End
elseif (mode = "modify") then		'수정모드
        ''기존 요청구분이 처리 요청으로 변경된경우 완료 처리를 NULL로
        sqlStr = " select top 1 * from [db_cs].[dbo].tbl_cs_memo"
        sqlStr = sqlStr + " where id = " + CStr(id) + " "

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
            sqlStr = sqlStr + " , retrydate = " + CStr(retrydate) + " "
			sqlStr = sqlStr + " , specialmemo = '" + CStr(html2db(specialmemo)) + "' "
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
			sqlStr = sqlStr + " , specialmemo = '" + CStr(html2db(specialmemo)) + "' "
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
            sqlStr = sqlStr + " , retrydate = " + CStr(retrydate) + " "
			sqlStr = sqlStr + " , specialmemo = '" + CStr(html2db(specialmemo)) + "' "
            sqlStr = sqlStr + " where id = " + CStr(id) + " "
            dbget.Execute sqlStr
        end if
		'response.write sqlStr&"<br>"
        response.write "<script>alert('수정되었습니다.'); </script>"
        response.write "<script>location.replace('" + retUrl + "?id=" & id & "'); </script>"
        'dbget.close()	:	response.End
elseif (mode = "finish") then
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

        response.write "<script>alert('완료되었습니다.'); </script>"
        response.write "<script>location.replace('" + retUrl + "?id=" & id & "'); </script>"
        'dbget.close()	:	response.End
elseif (mode = "delete") then
        sqlStr = " delete from [db_cs].[dbo].tbl_cs_memo "
        sqlStr = sqlStr + " where id = " + CStr(id) + " "
        dbget.Execute sqlStr

        response.write "<script>alert('삭제되었습니다.'); </script>"
        response.write "<script>location.replace('" + retUrl + "'); </script>"
        'dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
