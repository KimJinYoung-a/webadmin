<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
'// ������ ���� �α� ���� �Լ�
Sub AddLoginLog(param1,param2,param3)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")
    
     ''���� �α��� ���� ���� //2014/07/14 '' tbl_user_tenbyten ����α��� ����
    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&param1&"','"&Left(reFAddr,16)&"','"&param2&"','"&param3&"',0"
    dbget.Execute sqlStr
    
end Sub 

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF

	'�α��� Ȯ��
	if session("ssnTmpUID")="" or isNull(session("ssnTmpUID")) then   ''2017/04/21 ���� (ssBctId => ssnTmpUID)
		Call Alert_Return("�߸��� �����Դϴ�.")
		dbget.close()	:	response.End
	end if

	'// ���� ���� �� ���۰� ����
	dim userid, userpass, userpass2, sql
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)

    if (LCASE(userid)<>LCASE(session("ssnTmpUID"))) then
        Call Alert_Return("�߸��� �����Դϴ�...")
		dbget.close()	:	response.End
    end if

	'�н����� Ȯ��
	if (userpass<>userpass2) then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	if (chkPasswordComplex(userid,userpass)<>"" )then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpass) & "\n�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if


    dim puseridv, iEnc_password64
    sql = "select top 1 IsNULL(userdiv,'') as userdiv , Enc_password64"
    sql = sql + " from [db_partner].[dbo].tbl_partner"
    sql = sql + " where id = '" + userid + "'" + vbCrlf
    rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
    	puseridv = rsget("userdiv")
    	iEnc_password64 = rsget("Enc_password64")
    end if
    rsget.close
    
    if (UCASE(iEnc_password64)=UCASE(SHA256(MD5(userpass)))) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('���� ����Ͻ� ����� ������ ��й�ȣ�� ����Ͻ� �� �����ϴ�.');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if
    
    if (CLNG(puseridv)>=10) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('��� ���� ��� �Ұ�..');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if

	'// �н����� ����
	dbget.beginTrans

	on Error Resume Next
	sql = "Update [db_partner].[dbo].tbl_partner " + vbCrlf
	sql = sql + " set lastInfoChgDT=getdate(), Enc_password64='" & SHA256(MD5(userpass)) & "' " + vbCrlf
	sql = sql + " , Enc_password='' " + vbCrlf
	sql = sql + " where id = '" + userid + "'" + vbCrlf
	dbget.Execute(sql)

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
            
            Call AddLoginLog (userid,"R","") ''�н����� ���� R - flag
  
            
		'@@ �ش� �ε����� �̵�
		    if (session("ssnTmpUID")="10x10") then
                ''������.
                session.Abandon
		        dbget.close()	:	response.End

		    ''����Level
		    elseif (puseridv<=9) then
		        Session.Contents.Remove("ssnTmpUID")
		        
		    	response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/')</script>"
		        dbget.close()	:	response.End
		    else
		        response.write "<script language='javascript'>alert('���Ұ� �����Դϴ�.');</script>"
		        dbget.close()	:	response.End
		    end if    
		    

	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->