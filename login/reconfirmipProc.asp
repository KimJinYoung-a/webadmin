<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%  
	'// 변수 선언 및 전송값 접수
	dim userid  , sql  
	dim sMode
	dim manageUrl
	dim skey
	
	IF application("Svr_Info")="Dev" THEN
		manageUrl 	 = "http://testwebadmin.10x10.co.kr"
	ELSE
		manageUrl 	 = "http://webadmin.10x10.co.kr"
	END IF

	sMode		= requestCheckVar(trim(request.Form("hidM")),1)
	userid  = requestCheckVar(trim(request.Form("uid")),32)

 
	if userid ="" then
		 response.write("<script>alert('아이디가 입력되지 않았습니다.') ;history.back();</script>") 
     response.End
	end if  
	
SELECT CASE sMode
CASE "A" '비밀번호 인증
	dim sAuthno, dbAuthno
	sAuthno = requestCheckVar(trim(request.Form("sAuthno")),6)
		skey    = requestCheckVar(request.Form("skey"),32)
		
			if sAuthno ="" then
				 response.write("<script>alert('인증번호가 입력되지 않았습니다.') ;</script>") 
		     response.End
			end if 
		    
		    if (LCASE(userid)<>LCASE(session("reauthUID"))) then
                response.write("<script>alert(' 암호화에 문제가 발생했습니다. 확인 후 다시 시도해주세요');</script>") 
            response.end
            end if
            
			if skey <> md5(userid&"TPUAUTH") then
				response.write("<script>alert('아이디 암호화에 문제가 발생했습니다. 확인 후 다시 시도해주세요');</script>") 
				response.end
			end if	
			
			sql =" select top 1 authno from db_partner.dbo.tbl_partner_searchPWD_authno where userid ='"&userid&"' order by idx desc"
			rsget.Open sql,dbget,1
		  if  not rsget.EOF  then
		  	dbAuthno = rsget("authno")
		  end if
		 rsget.close  
		 
		 if dbAuthno <> sAuthno then
		 	response.write("<script>alert('잘못된 인증번호입니다.') ;parent.document.frmAuth.sAuthNo.value='';</script>") 
		   response.End
		 end if
		 
		sql ="update db_partner.dbo.tbl_partner_searchPWD_authno set isSucess='Y' where userid ='"&userid&"' and authno ='"&sAuthno&"'"
		dbget.Execute sql 
        
        ''파트너 인증IP추가. (로그인시 활용) 2017/04/13
        call AddPartnerAuthIpAdd(userid)
        
	Session.Contents.Remove("reauthUID")
  	
	response.write("<script>alert('인증되었습니다. 다시로그인해주시기 바랍니다.');</script>")  
	response.write("<script>top.location.href='/';</script>")  
	response.end

END SELECT


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->