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

'   sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
'	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
'	sqlStr = sqlStr + " values(" + VbCrlf
'	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
'	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
'	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
'	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
'	sqlStr = sqlStr + " )" + VbCrlf
'    dbget.Execute sqlStr
    
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
	if session("ssnTmpUIDPartner")="" or isNull(session("ssnTmpUIDPartner")) then
		Call Alert_Return("�߸��� �����Դϴ�.")
		dbget.close()	:	response.End
	end if

	'// ���� ���� �� ���۰� ����
	dim userid, userpass, userpass2, sql
	dim userpassSec1, userpassSec2
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)
	userpassSec1 = requestCheckVar(trim(request.Form("upwdS1")),32)
	userpassSec2 = requestCheckVar(trim(request.Form("upwdS2")),32)

    if (LCASE(userid)<>LCASE(session("ssnTmpUIDPartner"))) then
        Call Alert_Return("�߸��� �����Դϴ�...")
		dbget.close()	:	response.End
    end if

	'�н����� Ȯ��
	if userpass<>userpass2 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('��й�ȣ Ȯ���� Ʋ���ϴ�.\n��Ȯ�� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	if chkPasswordComplex(userid,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpass) & "\n�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	if chkPasswordComplex(userid,userpassSec1)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpassSec1) & "\n2�� ��й�ȣ�� ��ȿ���� �ʽ��ϴ�.�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"	parent.document.forms[0].upwdS1.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwdS2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwdS1.focus();" &vbCrLf &_
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
    
    ''1����� ���� ������ //2017/04/24
    if (UCASE(iEnc_password64)=UCASE(SHA256(MD5(userpass)))) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('���� ����Ͻ� ����� ������ ��й�ȣ�� ����Ͻ� �� �����ϴ�.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if
    
    if (CLNG(puseridv)<10) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('��� ���� ��� �Ұ�..');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if
    
    ''�����ӽ�--------------------------------------------------- 2016/06/20
    dim cuseridv
    sql = "select top 1 IsNULL(userdiv,'') as userdiv "
    sql = sql + " from [db_user].[dbo].tbl_user_c"
    sql = sql + " where userid = '" + userid + "'" + vbCrlf
    
    rsget.CursorLocation = adUseClient
    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
    	cuseridv = rsget("userdiv")
    end if
    rsget.close
            
            
	'// �н����� ����
	dbget.beginTrans

	on Error Resume Next
	
	sql = "update [db_user].[dbo].tbl_logindata" + VbCrlf
	sql = sql + " set  Enc_userpass64='" + SHA256(MD5(userpass)) + "'" + VbCrlf
	sql = sql + " , Enc_userpass=''" + VbCrlf	
	sql = sql + " where userid='" + userid + "'" 
	rsget.Open  sql,dbget,1


	sql = "Update [db_partner].[dbo].tbl_partner " + vbCrlf
	sql = sql + " set lastInfoChgDT=getdate(), Enc_password64='" & SHA256(MD5(userpass)) & "' " + vbCrlf
	sql = sql + " , Enc_password='' " + vbCrlf
	sql = sql + " , Enc_2password64='" & SHA256(MD5(userpassSec1)) & "' " + vbCrlf 
	sql = sql + " where id = '" + userid + "'"    
	dbget.Execute(sql)
 
	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
            
            Call AddLoginLog (userid,"R","") ''�н����� ���� R - flag
            Session.Contents.Remove("ssnTmpUID")
            
            ''���� �ӽ�
        	if (cuseridv="14") then
        	    Session.Contents.Remove("ssnTmpUID")
        	    session("ssUserCDiv")=cuseridv  ''2016/08/11 �߰�
    			response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
            	dbget.close()	:	response.End
        	end if
            ''-----------------------------------------------------------
            
		'@@ �ش� �ε����� �̵�
		    if (userid="10x10") then
                ''������.
                session.Abandon
		        dbget.close()	:	response.End

		    ''����Level
		    elseif (puseridv=999) then
		    	''���� ��ü (yahoo, empas..)
		        response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=9999) then
		    	''�귣�� ��ü
		    	response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=9000) then
		    	''���� ��ü
		    	response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=501) or (puseridv=502) or (puseridv=503) then
		    	response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=101) or (puseridv=111) or (puseridv=112) or (puseridv=201) or (puseridv=301) then
		    	response.write "<script language='javascript'>alert('��й�ȣ�� ����Ǿ����ϴ�. �ٽ÷α��� �� �ּ���.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    else
		        session.Abandon
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