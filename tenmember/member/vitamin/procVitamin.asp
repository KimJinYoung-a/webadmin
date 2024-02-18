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
dim requestvm,didx 
dim strSql
sMode =  requestCheckvar(request("hidM"),1)
empno =  session("ssBctSn")
requestvm =  requestCheckvar(request("reqVM"),20)
 
if sMode ="I" THEN
		   strSql = " SELECT top 1 midx, totvm,  usevm   from db_partner.dbo.tbl_vitamin_master where empno = '"&empno&"' and isusing = 1 and startday <= getdate() and endday >=getdate() "  
		 rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				midx = rsget("midx")
				totvm = rsget("totvm") 
				usevm = rsget("usevm") 
			END IF
		 rsget.close
		 
		 if midx ="" THEN
		 		response.write "<script>alert('등록가능한 비타민이 존재하지 않습니다.');</script>"
		 		response.write "<script>history.back();</script>"
				response.end
		end if

		 if cdbl(totvm - usevm ) < cdbl(requestvm) THEN 
		 		response.write "<script>alert('신청한 금액이 잔액보다 큰값입니다. 신청금액을 변경해주세요');</script>"
		 		response.write "<script>history.back();</script>"
				response.end
		 end if

		 strSql = "insert into db_partner.dbo.tbl_vitamin_detail ( midx, vmmoney, vmstatus, adminid, lastupdate) "& vbCrlf
		 strSql = strSql & "  values ( "&midx&", '"&requestvm&"', 0, '"&userid&"',getdate() ) " 
		  dbget.Execute(strSql)
		  
		strSql = "select SCOPE_IDENTITY() as didx "
		rsget.Open strSql, dbget, 0
		if not rsget.eof then
			didx = rsget("didx") 
		end if 
		rsget.Close
  
		  strSql = " update db_partner.dbo.tbl_vitamin_master set usevm = usevm+'"&requestvm&"', adminid='"&session("ssBctID")&"', lastupdate=getdate()"& vbCrlf
		  strSql = strSql & " WHERE midx = "&midx
		   dbget.Execute(strSql)
		   
		  response.write "<script>alert('신청 되었습니다. 결재를 등록해주세요');</script>"
			response.write "<script>location.href='/admin/approval/eapp/regeapp.asp?iAidx=351&ieidx=33&iSL="&didx&"&mRP="&requestvm&"';  window.resizeTo( 880, 800 );</script>"
	 
ELSEIF sMode="D" then
	didx =  requestCheckvar(request("didx"),10)
	
	strSql =" update m set m.usevm = m.usevm - d.vmmoney, lastupdate =getdate(), adminid = '"&session("ssBctID")&"' "
	strSql = strSql &" from db_partner.dbo.tbl_vitamin_master as m "
	strSql = strSql &" inner join db_partner.dbo.tbl_vitamin_detail  as d on m.midx = d.midx "
	strSql = strSql &" where m.isusing =1 and d.isusing =1 and d.vmstatus =0 and d.didx =  "&didx 
  dbget.Execute(strSql)
	
	strSql =" update db_partner.dbo.tbl_vitamin_detail "
	strSql = strSql &" set isusing = 0, lastupdate =getdate(), adminid = '"&session("ssBctID")&"' "
	strSql = strSql &" where  didx = "&didx&" and isusing =1 and 	vmstatus =0 "
	 dbget.Execute(strSql)

	response.write "<script>alert('삭제되었습니다.');</script>"
	response.write "<script>location.href='popListVitamin.asp';</script>" 
	 
ELSE
	response.write "<script>alert('데이터처리에 문제가 발생했습니다. 확인 후 재시도해주세요');</script>"
	response.write "<script>history.back(-1);</script>"
END IF
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
