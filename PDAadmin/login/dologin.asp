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
    response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� �Էµ��� �ʾҽ��ϴ�.');</script>")
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
		session("ssAdminPsn") = rsget("part_sn")		'�μ� ��ȣ
		session("ssAdminLsn") = rsget("level_sn")		'��� ��ȣ
	end if
    rsget.close


	''�����ӽ�
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
			response.write "<script language='javascript'>alert('�н����带 ���̵�� ���� �ϰ� ����ϰ� �ֽ��ϴ�. \n\n�н����带 ���� �� �ּ���. \n��ü�������� ���� �ϴ� �귣���й�ȣ ���濡�� �����մϴ�.');</script>"
		end if
        
        ''backpath Redirect
        if (request("backpath")<>"") then
            response.redirect request("backpath")    
            dbget.close()	:	response.End
        end if

		''���� �ӽ�
		if (cuseridv="14") then				
				response.write "<script language='javascript'>location.replace('/lectureadmin/index.asp')</script>"
            	dbget.close()	:	response.End
		end if


        ''�����ΰ��. ����ƮPw�� ����Pw�� ������� errMsg
        if (session("ssBctDiv")<=9) then
            sql = "select * from [db_partner].[dbo].tbl_partner p,"
            sql = sql + " [db_user].[dbo].tbl_logindata u"
            sql = sql + " where p.id='" & session("ssBctId") & "'"
            sql = sql + " and p.id=u.userid"
            sql = sql + " and p.password=u.userpass"
            
            rsget.Open sql,dbget,1
        	if  not rsget.EOF  then
        		errMsg = "����Ʈ �� ���� ��й�ȣ�� �����ϰ� ����ϰ� �ֽ��ϴ�. \n\nMyInfo���� ���� ��й�ȣ�� �����Ͽ� ����ϼ���."
        	end if
        	rsget.close
        end if
        
    
        if (session("ssBctId")="10x10") then
            ''������.
            sql = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
        	sql = sql + " (userid,refip)" + VbCrlf
        	sql = sql + " values(" + VbCrlf
        	sql = sql + " '" + session("ssBctId") + "'," + VbCrlf
        	sql = sql + " '" + Left(reFAddr,32) + "'"
        	sql = sql + " )" + VbCrlf

        	rsget.Open sql,dbget,1
        	
            response.write "<script language='javascript'>location.replace('/admin/index.asp')</script>"
            dbget.close()	:	response.End
        
        ''����Level
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
        	''���� ��ü (yahoo, empas..)
            response.write "<script language='javascript'>location.replace('/company/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=9999) then
        	''�귣�� ��ü
        	response.write "<script language='javascript'>location.replace('/designer/index.asp')</script>"
            dbget.close()	:	response.End
        elseif (session("ssBctDiv")=9000) then
        	''���� ��ü
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
        response.write("<script>window.alert('���̵� �Ǵ� ��й�ȣ�� Ʋ�Ƚ��ϴ�.');</script>")
        response.write("<script>history.go(-1);</script>")
        dbget.close()	:	response.End
    end if
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
