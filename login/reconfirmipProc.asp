<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/login/partner_loginCheck_function.asp"-->
<%  
	'// ���� ���� �� ���۰� ����
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
		 response.write("<script>alert('���̵� �Էµ��� �ʾҽ��ϴ�.') ;history.back();</script>") 
     response.End
	end if  
	
SELECT CASE sMode
CASE "A" '��й�ȣ ����
	dim sAuthno, dbAuthno
	sAuthno = requestCheckVar(trim(request.Form("sAuthno")),6)
		skey    = requestCheckVar(request.Form("skey"),32)
		
			if sAuthno ="" then
				 response.write("<script>alert('������ȣ�� �Էµ��� �ʾҽ��ϴ�.') ;</script>") 
		     response.End
			end if 
		    
		    if (LCASE(userid)<>LCASE(session("reauthUID"))) then
                response.write("<script>alert(' ��ȣȭ�� ������ �߻��߽��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���');</script>") 
            response.end
            end if
            
			if skey <> md5(userid&"TPUAUTH") then
				response.write("<script>alert('���̵� ��ȣȭ�� ������ �߻��߽��ϴ�. Ȯ�� �� �ٽ� �õ����ּ���');</script>") 
				response.end
			end if	
			
			sql =" select top 1 authno from db_partner.dbo.tbl_partner_searchPWD_authno where userid ='"&userid&"' order by idx desc"
			rsget.Open sql,dbget,1
		  if  not rsget.EOF  then
		  	dbAuthno = rsget("authno")
		  end if
		 rsget.close  
		 
		 if dbAuthno <> sAuthno then
		 	response.write("<script>alert('�߸��� ������ȣ�Դϴ�.') ;parent.document.frmAuth.sAuthNo.value='';</script>") 
		   response.End
		 end if
		 
		sql ="update db_partner.dbo.tbl_partner_searchPWD_authno set isSucess='Y' where userid ='"&userid&"' and authno ='"&sAuthno&"'"
		dbget.Execute sql 
        
        ''��Ʈ�� ����IP�߰�. (�α��ν� Ȱ��) 2017/04/13
        call AddPartnerAuthIpAdd(userid)
        
	Session.Contents.Remove("reauthUID")
  	
	response.write("<script>alert('�����Ǿ����ϴ�. �ٽ÷α������ֽñ� �ٶ��ϴ�.');</script>")  
	response.write("<script>top.location.href='/';</script>")  
	response.end

END SELECT


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->