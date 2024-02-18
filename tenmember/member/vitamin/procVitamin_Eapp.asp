<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sReturnUrl
dim empno,userid, reportidx,reportprice,midx
dim totvm, reqvm, usevm,payvm
dim sMode
dim didx 
dim strSql
dim authstate,vmstatus
  
 
empno =  session("ssBctSn")  
 didx			= requestCheckvar(request("iSL"),10)		 
 sReturnUrl 			= requestCheckvar(request("hidRU"),100)
 authstate = requestCheckvar(request("ias"),10)	
 
		 if didx ="" or authstate ="" THEN 
		 		response.write "<script>alert('파라미터에 문제가 발생했습니다.');</script>"
		 		response.write "<script>history.back();</script>"
				response.end
			end if	
	
	if authstate = 7 THEN
		vmstatus =1 
	else
		vmstatus = authstate
	end if
		 	
	if authstate = "5"	   then 
		 strSql = "update m set usevm =  m.usevm - d.vmmoney "
		 strSql = strSql & " from db_partner.dbo.tbl_vitamin_master as m "
		 strSql = strSql & " inner join db_partner.dbo.tbl_vitamin_detail as d on m.midx = d.midx "
		 strSql = strSql &"	  where  didx = " &didx &" and (vmstatus = 0 or vmstatus = 3)  and d.isusing =1 "
		 dbget.Execute(strSql)
		 
		 strSql = "update db_partner.dbo.tbl_vitamin_detail set  vmstatus="&vmstatus&", adminid='"&session("ssBctID")&"', lastupdate=getdate()  "& vbCrlf
		 strSql = strSql & "  where didx = " &didx 
		 dbget.Execute(strSql)
	 
	else 
	 
		 strSql = "update db_partner.dbo.tbl_vitamin_detail set  vmstatus="&vmstatus&", adminid='"&session("ssBctID")&"', lastupdate=getdate()  "& vbCrlf
		 strSql = strSql & "  where didx = " &didx 
		 dbget.Execute(strSql) 
		
  end if	 
		   
		 
		   
		  response.write "<script>alert('저장 되었습니다.');</script>"
			response.write "<script>location.href='"&sReturnUrl&"'</script>"
			
	 

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
