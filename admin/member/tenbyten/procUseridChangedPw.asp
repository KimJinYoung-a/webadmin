<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ѵ��
' History : 2011.01.19 ������ ����
'			2017.09.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<% 
Dim spassword, sUserid, sql
Dim Enc_userpass, Enc_userpass64
Dim objCmd, returnValue
sUserid  = requestCheckvar(Request("uid"),32)
spassword =  requestCheckvar(Request("spw"),32)

If not(C_ADMIN_AUTH or C_PSMngPart) Then
	response.write "<script  type='text/javascript'>"
	response.write "	alert('������ �����ϴ�.');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close() : response.end
end if

if chkPasswordComplexNonID(spassword)<>"" then
	response.write "<script language='javascript'>" &vbCrLf &_
					"	alert('" & chkPasswordComplex(sUserid,spassword) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');" &vbCrLf &_
					" 	history.back();" &vbCrLf &_
					"</script>"
	dbget.close()	:	response.End
end if

Enc_userpass = MD5(spassword)
Enc_userpass64 = SHA256(MD5(spassword))

' ��Ʈ�� ���� ����ħ
sql = "update p" & vbCrlf
sql = sql & " set p.Enc_password64 = '" & Enc_userpass64 & "' " & vbCrlf
sql = sql & " ,p.Enc_password = '' " & vbCrlf
sql = sql & " , p.lastlogindt=getdate()" & vbCrlf
sql = sql & " , p.lastpwchgdt=getdate()" & vbCrlf
sql = sql & " , p.isusing='Y'" & vbCrlf
sql = sql & " from [db_partner].[dbo].tbl_partner p with (nolock)" & vbCrlf
sql = sql & " join [db_partner].[dbo].tbl_user_tenbyten t with (nolock)" & vbCrlf
sql = sql & " 	on p.id = t.userid " & vbCrlf
sql = sql & " where p.id = '" & sUserid & "'" & vbCrlf

'response.write sql & "<br>"
dbget.Execute sql

' ���� ���� ����ħ
sql = "update t" & vbCrlf
sql = sql & " set t.isusing=1" & vbCrlf
sql = sql & " from [db_partner].[dbo].tbl_partner p with (nolock)" & vbCrlf
sql = sql & " join [db_partner].[dbo].tbl_user_tenbyten t with (nolock)" & vbCrlf
sql = sql & " 	on p.id = t.userid " & vbCrlf
sql = sql & " where t.userid = '" & sUserid & "'" & vbCrlf

'response.write sql & "<br>"
dbget.Execute sql

' �н����� Ʋ���� ����� ������.
sql = "update l set l.loginsuccess='C'" & vbCrlf
sql = sql & " from (" & vbCrlf
sql = sql & " 	select max(idx) as maxidx" & vbCrlf
sql = sql & " 	from db_log.dbo.tbl_partner_login_log with (nolock)" & vbCrlf
sql = sql & " 	where userid = '" & sUserid & "'" & vbCrlf
sql = sql & " ) as t" & vbCrlf
sql = sql & " join db_log.dbo.tbl_partner_login_log as l with (nolock)" & vbCrlf
sql = sql & " 	on maxidx=l.idx" & vbCrlf
sql = sql & " 	and l.loginsuccess='N'"

'response.write sql & "<br>"
dbget.Execute sql

sql = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&sUserid&"','"&Left(request.ServerVariables("REMOTE_ADDR"),16)&"','R','',0"
dbget.Execute sql

Call Alert_close ("����Ǿ����ϴ�.")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->