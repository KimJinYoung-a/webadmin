<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim userid, userpass, backurl
userid  = Left(trim(request.Form("uid")),32)
userpass = Left(trim(request.Form("upwd")),32)

dim dbpassword
dim sql

dim reFAddr
reFAddr = request.ServerVariables("REMOTE_ADDR")

dim errMsg

if ( userid = "" or userpass = "") then
    response.write("<script>window.alert('아이디 또는 비밀번호가 입력되지 않았습니다.');</script>")
    response.write("<script>history.go(-1);</script>")
    dbget.close()	:	response.End
else
    sql = "select top 1 id,company_name,tel,fax,url,email,bigo,userdiv,password,groupid " + vbCrlf
    sql = sql + "	, part_sn, level_sn " + vbCrlf
    sql = sql + " from [db_partner].[dbo].tbl_partner where id = '" + userid + "'" + vbCrlf
    sql = sql + " and password = '" + userpass + "' " + vbCrlf
    sql = sql + " and isusing='Y'"
    rsget.Open sql,dbget,1

    if  not rsget.EOF  then
    	dbpassword  = rsget("password")
        session("ssBctId") = rsget("id")
        session("ssBctDiv") = rsget("userdiv")
        session("ssBctBigo") = rsget("bigo")
        session("ssBctCname") = db2html(rsget("company_name"))
		session("ssBctEmail") = db2html(rsget("email"))
		session("ssGroupid") = rsget("groupid")
		session("ssAdminPsn") = rsget("part_sn")		'부서 번호
		session("ssAdminLsn") = rsget("level_sn")		'등급 번호
	end if
    rsget.close


	''강사임시
	dim cuseridv
	sql = "select top 1 * "
	sql = sql + " from [db_user].[dbo].tbl_user_c"
	sql = sql + " where userid = '" + userid + "'" + vbCrlf

	rsget.Open sql,dbget,1
	if  not rsget.EOF  then
		cuseridv = rsget("userdiv")
	end if
	rsget.close
    
	if trim(LCase(dbpassword))=trim(LCase(userpass)) then
		response.Cookies("partner").domain = "10x10.co.kr"
        response.Cookies("partner")("userid") = session("ssBctId")
        response.Cookies("partner")("userdiv") = session("ssBctDiv")


		if LCase(userpass)=LCase(userid) then
			response.write "<script language='javascript'>alert('패스워드를 아이디와 동일 하게 사용하고 있습니다. \n\n패스워드를 변경 해 주세요. \n업체정보수정 제일 하단 브랜드비밀번호 변경에서 가능합니다.');</script>"
		end if
        
        ''backpath Redirect
        if (request("backpath")<>"") then
            response.redirect request("backpath")    
            dbget.close()	:	response.End
        end if

		''강사 임시
		if (cuseridv="14") then				
				response.write "<script language='javascript'>location.replace('/lectureadmin/index.asp')</script>"
            	dbget.close()	:	response.End
		end if


        ''직원인경우. 프런트Pw와 어드민Pw가 같을경우 errMsg
        if (session("ssBctDiv")<=9) then
            sql = "select * from [db_partner].[dbo].tbl_partner p,"
            sql = sql + " [db_user].[dbo].tbl_logindata u"
            sql = sql + " where p.id='" & session("ssBctId") & "'"
            sql = sql + " and p.id=u.userid"
            sql = sql + " and p.password=u.userpass"
            
            rsget.Open sql,dbget,1
        	if  not rsget.EOF  then
        		errMsg = "프런트 와 어드민 비밀번호를 동일하게 사용하고 있습니다. \n\nMyInfo에서 어드민 비밀번호를 변경하여 사용하세요."
        	end if
        	rsget.close
        end if
        
    
        if (session("ssBctId")="10x10") then
            ''사용안함.
            sql = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
        	sql = sql + " (userid,refip)" + VbCrlf
        	sql = sql + " values(" + VbCrlf
        	sql = sql + " '" + session("ssBctId") + "'," + VbCrlf
        	sql = sql + " '" + Left(reFAddr,32) + "'"
        	sql = sql + " )" + VbCrlf

        	rsget.Open sql,dbget,1
        	
            response.write "<script language='javascript'>location.replace('/admin/index.asp')</script>"
            dbget.close()	:	response.End
        
        ''직원Level
        elseif (session("ssBctDiv")<=9) then
        	sql = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
        	sql = sql + " (userid,refip)" + VbCrlf
        	sql = sql + " values(" + VbCrlf
        	sql = sql + " '" + session("ssBctId") + "'," + VbCrlf
        	sql = sql + " '" + Left(reFAddr,32) + "'"
        	sql = sql + " )" + VbCrlf

        	rsget.Open sql,dbget,1
        	
        	if (errMsg<>"") then
                response.write "<script language='javascript'>alert('" & errMsg & "');</script>"
            end if
            
            if inStr(Request.ServerVariables("HTTP_USER_AGENT"),"Windows CE")>0 then
				response.redirect "/PDAadmin/index.asp"
				dbget.close()	:	response.End
			else
				response.write "<script language='javascript'>location.replace('/admin/index.asp')</script>"
            dbget.close()	:	response.End
			end if
        
        	
        elseif (session("ssBctDiv")=999) then
        	''제휴 업체 (yahoo, empas..)
            response.write "<script language='javascript'>location.replace('/company/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=9999) then
        	''브랜드 업체
        	response.write "<script language='javascript'>location.replace('/designer/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=9000) then
        	''강사 업체
        	response.write "<script language='javascript'>location.replace('/lectureradmin/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=501) or (session("ssBctDiv")=502) or (session("ssBctDiv")=503) then
        	sql = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
        	sql = sql + " (userid,refip)" + VbCrlf
        	sql = sql + " values(" + VbCrlf
        	sql = sql + " '" + session("ssBctId") + "'," + VbCrlf
        	sql = sql + " '" + Left(reFAddr,32) + "'"
        	sql = sql + " )" + VbCrlf

        	rsget.Open sql,dbget,1

        	response.write "<script language='javascript'>location.replace('/offshop/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=101) or (session("ssBctDiv")=111) or (session("ssBctDiv")=112) or (session("ssBctDiv")=201) or (session("ssBctDiv")=301) then
        	sql = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
        	sql = sql + " (userid,refip)" + VbCrlf
        	sql = sql + " values(" + VbCrlf
        	sql = sql + " '" + session("ssBctId") + "'," + VbCrlf
        	sql = sql + " '" + Left(reFAddr,32) + "'"
        	sql = sql + " )" + VbCrlf

        	rsget.Open sql,dbget,1

        	response.write "<script language='javascript'>location.replace('/admin/index.asp')</script>"
            dbget.close()	:	response.End
        end if



    else
        response.write("<script>window.alert('아이디 또는 비밀번호가 틀렸습니다.');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
    end if
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
