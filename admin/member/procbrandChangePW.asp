<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��й�ȣ ����
' History : 2014.02.03 ������ ����
'			2021.07.16 �ѿ�� ����(2���н����� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<% 
Sub AddLoginLog(param1,param2,param3)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")

    ''���� �α��� ���� ���� //2014/07/14 '' tbl_user_tenbyten ����α��� ����
    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&param1&"','"&Left(reFAddr,16)&"','"&param2&"','"&param3&"',0"
    dbget.Execute sqlStr

end Sub

Dim spassword, sbrandid, sType,spasswordS
Dim Enc_userpass, Enc_userpass64, Enc_2userpass64
Dim objCmd, returnValue
sbrandid  = requestCheckvar(Request("bid"),32)
spassword =  requestCheckvar(Request("spw"),32)
'spasswordS =  requestCheckvar(Request("sPWS1"),32)
sType =  requestCheckvar(Request("spw"),32)
	'������ ���� Ȯ��
	if not (C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or C_MD or C_OFF_part or C_logics_Part) then   
			Call Alert_close ("�����ڸ� ���氡���մϴ�.������ Ȯ�����ּ���")
	end if 	
	
'//�н����� ��å �˻�(2008.12.15;������) 
 if chkPasswordComplex(sbrandid,spassword)<>"" then
 	response.write "<script language='javascript'>" &vbCrLf &_
 					"	alert('" & chkPasswordComplex(sbrandid,spassword) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');" &vbCrLf &_
 					" 	history.back();" &vbCrLf &_
 					"</script>"
 	dbget.close()	:	response.End
 end if
 
' if chkPasswordComplex(sbrandid,spasswordS)<>"" then
'  	response.write "<script language='javascript'>" &vbCrLf &_
'  					"	alert('" & chkPasswordComplex(sbrandid,spasswordS) & "\n��й�ȣ�� Ȯ���� �ٽ� �õ����ּ���.');" &vbCrLf &_
'  					" 	history.back();" &vbCrLf &_
'  					"</script>"
'  	dbget.close()	:	response.End
'  end if
 
 Enc_userpass = MD5(spassword)
 Enc_userpass64 = SHA256(MD5(spassword))
 'Enc_2userpass64= SHA256(MD5(spasswordS))
 	Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partner_changePW]('"&sbrandid&"','"&Enc_userpass&"','"&Enc_userpass64&"','' )}"	' "&Enc_2userpass64&"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    returnValue = objCmd(0).Value
	Set objCmd = nothing
	
	IF returnValue = 0 THEN 
			Call Alert_return ("������ ó���� ������ ������ϴ�. �����ڿ��� �������ּ���")
		response.end
END IF
    CALL AddLoginLog(sbrandid,"C","") ''2016/09/27 '' C Ÿ�� �߰�. ������ �н����� ����. 15�� ��ٸ� ȸ��
	Call Alert_close ("����Ǿ����ϴ�.")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->