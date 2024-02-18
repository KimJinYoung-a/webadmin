<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 비밀번호 변경
' History : 2014.02.03 정윤정 생성
'			2021.07.16 한용민 수정(2차패스워드 제거)
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

    ''최종 로그인 일자 저장 //2014/07/14 '' tbl_user_tenbyten 사번로그인 제외
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
	'관리자 여부 확인
	if not (C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or C_MD or C_OFF_part or C_logics_Part) then   
			Call Alert_close ("관리자만 변경가능합니다.권한을 확인해주세요")
	end if 	
	
'//패스워드 정책 검사(2008.12.15;허진원) 
 if chkPasswordComplex(sbrandid,spassword)<>"" then
 	response.write "<script language='javascript'>" &vbCrLf &_
 					"	alert('" & chkPasswordComplex(sbrandid,spassword) & "\n비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
 					" 	history.back();" &vbCrLf &_
 					"</script>"
 	dbget.close()	:	response.End
 end if
 
' if chkPasswordComplex(sbrandid,spasswordS)<>"" then
'  	response.write "<script language='javascript'>" &vbCrLf &_
'  					"	alert('" & chkPasswordComplex(sbrandid,spasswordS) & "\n비밀번호를 확인후 다시 시도해주세요.');" &vbCrLf &_
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
			Call Alert_return ("데이터 처리에 문제가 생겼습니다. 관리자에게 문의해주세요")
		response.end
END IF
    CALL AddLoginLog(sbrandid,"C","") ''2016/09/27 '' C 타입 추가. 관리자 패스워드 변경. 15분 기다림 회피
	Call Alert_close ("변경되었습니다.")
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->