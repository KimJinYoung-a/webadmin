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
        
	 session("AuthUID") = userid
	 session("AuthChk") = "Y"
	
	response.write("<script>top.location.href='/login/searchPwdNew.asp';</script>")  
	response.end
CASE "C" '��й�ȣ ����
	Session.Contents.Remove("AuthUID")
  Session.Contents.Remove("AuthChk") 
	dim  userpass, userpass2 
	dim userpassSec1, userpassSec2
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)
	userpassSec1 = requestCheckVar(trim(request.Form("upwdS1")),32)
	userpassSec2 = requestCheckVar(trim(request.Form("upwdS2")),32)
 
	'�н����� Ȯ��
	if userpass<>userpass2 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"history.back();"&vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	if chkPasswordComplex(userid,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpass) & "\n�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_					
						"history.back();"&vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	if chkPasswordComplex(userid,userpassSec1)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpassSec1) & "\n2�� ��й�ȣ�� ��ȿ���� �ʽ��ϴ�.�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
							"history.back();"&vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
dim Enc_userpass64, Enc_userpass,Enc_2userpass64
Enc_userpass =  MD5(userpass)
Enc_userpass64 = SHA256(MD5(userpass)) 
Enc_2userpass64= SHA256(MD5(userpassSec1))

	'// �н����� ���� 
	sql = "Update [db_partner].[dbo].tbl_partner " & vbCrlf
	sql = sql & " set lastInfoChgDT=getdate(), Enc_password64='" &Enc_userpass64 & "' " & vbCrlf
	sql = sql & " , Enc_password='' " & vbCrlf
	sql = sql & " , Enc_2password64='" & Enc_2userpass64 & "' " & vbCrlf 
	sql = sql & " where id = '" & userid & "'"   
	dbget.Execute(sql)
	
	sql = " IF Not Exists(select * from [db_user].[dbo].tbl_user_n where userid='"&userid&"')" & VbCrlf
  sql = sql & " BEGIN " &VbCrlf
  sql = sql & "     update L" &VbCrlf
  sql = sql & "     set  Enc_userpass64='" & Enc_userpass64 & "'" & VbCrlf
  sql = sql & "     , Enc_userpass=''" & VbCrlf   
  sql = sql & "     from [db_user].[dbo].tbl_logindata L" & VbCrlf   
  sql = sql & "         inner Join [db_user].[dbo].tbl_user_c C" & VbCrlf   
  sql = sql & "         on L.userid=C.userid" & VbCrlf   
  sql = sql & "     where L.userid='" & userid & "'"  & VbCrlf   
  sql = sql & " END " 
  dbget.Execute sql
         
	sql = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&userid&"','"&Left(request.ServerVariables("REMOTE_ADDR"),16)&"','R','',0"
  dbget.Execute sql
    
	response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('����Ǿ����ϴ�.');" &vbCrLf &_
						" location.href='"&manageUrl&"/index.asp'" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	
END SELECT


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->