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

    sqlStr = " insert into [db_log].[dbo].tbl_partner_login_log" + VbCrlf
	sqlStr = sqlStr + " (userid,refip,loginSuccess,USBTokenSn)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + param1 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + Left(reFAddr,16) + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param2 + "'," + VbCrlf
	sqlStr = sqlStr + " '" + param3 + "'" + VbCrlf
	sqlStr = sqlStr + " )" + VbCrlf

    dbget.Execute sqlStr
end Sub 

	'�α��� Ȯ��
	if session("ssBctSn")="" or isNull(session("ssBctSn")) then
		Call Alert_Return("�߸��� �����Դϴ�.")
		dbget.close()	:	response.End
	end if

	'// ���� ���� �� ���۰� ����
	dim empno, userpass, userpass2, sql
	empno  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)

    if (LCASE(empno)<>LCASE(session("ssBctSn"))) then
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
	if chkPasswordComplex(empno,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(empno,userpass) & "\n�ٸ� ��й�ȣ�� �Է����ּ���.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if

	'// �н����� ����
	dbget.beginTrans

	on Error Resume Next
	sql = "Update [db_partner].[dbo].tbl_user_tenbyten " + vbCrlf
	sql = sql + " set Enc_emppass64='" & SHA256(MD5(userpass)) & "' " + vbCrlf
	sql = sql + " where empno = '" + empno + "'" + vbCrlf
	dbget.Execute(sql)

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		
		Call AddLoginLog (empno,"R","") ''�н����� ���� R - flag
		response.write "<script language='javascript'>top.location.replace('/tenmember/index.asp')</script>"

	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->