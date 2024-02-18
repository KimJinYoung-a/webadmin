<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
'// 웹어드민 접속 로그 저장 함수
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
    
     ''최종 로그인 일자 저장 //2014/07/14 '' tbl_user_tenbyten 사번로그인 제외
    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&param1&"','"&Left(reFAddr,16)&"','"&param2&"','"&param3&"',0"
    dbget.Execute sqlStr
    
end Sub 

dim manageUrl
IF application("Svr_Info")="Dev" THEN
	manageUrl 	 = "http://testwebadmin.10x10.co.kr"
ELSE
	manageUrl 	 = "http://webadmin.10x10.co.kr"
END IF

	'로그인 확인
	if session("ssnTmpUIDPartner")="" or isNull(session("ssnTmpUIDPartner")) then
		Call Alert_Return("잘못된 접속입니다.")
		dbget.close()	:	response.End
	end if

	'// 변수 선언 및 전송값 접수
	dim userid, userpass, userpass2, sql
	dim userpassSec1, userpassSec2
	userid  = requestCheckVar(trim(request.Form("uid")),32)
	userpass = requestCheckVar(trim(request.Form("upwd")),32)
	userpass2 = requestCheckVar(trim(request.Form("upwd2")),32)
	userpassSec1 = requestCheckVar(trim(request.Form("upwdS1")),32)
	userpassSec2 = requestCheckVar(trim(request.Form("upwdS2")),32)

    if (LCASE(userid)<>LCASE(session("ssnTmpUIDPartner"))) then
        Call Alert_Return("잘못된 접속입니다...")
		dbget.close()	:	response.End
    end if

	'패스워드 확인
	if userpass<>userpass2 then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('비밀번호 확인이 틀립니다.\n정확한 비밀번호를 입력해주세요.');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	if chkPasswordComplex(userid,userpass)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpass) & "\n다른 비밀번호를 입력해주세요.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
	end if
	
	if chkPasswordComplex(userid,userpassSec1)<>"" then
		response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('" & chkPasswordComplex(userid,userpassSec1) & "\n2차 비밀번호가 유효하지 않습니다.다른 비밀번호를 입력해주세요.');" &vbCrLf &_
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
    
    ''1차비번 동일 사용금지 //2017/04/24
    if (UCASE(iEnc_password64)=UCASE(SHA256(MD5(userpass)))) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('기존 사용하신 비번과 동일한 비밀번호를 사용하실 수 없습니다.');" &vbCrLf &_
						"	parent.document.forms[0].upwd.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd2.value='';" &vbCrLf &_
						"	parent.document.forms[0].upwd.focus();" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if
    
    if (CLNG(puseridv)<10) then
        response.write "<script language='javascript'>" &vbCrLf &_
						"	alert('비번 변경 사용 불가..');" &vbCrLf &_
						"</script>"
		dbget.close()	:	response.End
    end if
    
    ''강사임시--------------------------------------------------- 2016/06/20
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
            
            
	'// 패스워드 변경
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
 
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
            
            Call AddLoginLog (userid,"R","") ''패스워드 수정 R - flag
            Session.Contents.Remove("ssnTmpUID")
            
            ''강사 임시
        	if (cuseridv="14") then
        	    Session.Contents.Remove("ssnTmpUID")
        	    session("ssUserCDiv")=cuseridv  ''2016/08/11 추가
    			response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
            	dbget.close()	:	response.End
        	end if
            ''-----------------------------------------------------------
            
		'@@ 해당 인덱스로 이동
		    if (userid="10x10") then
                ''사용안함.
                session.Abandon
		        dbget.close()	:	response.End

		    ''직원Level
		    elseif (puseridv=999) then
		    	''제휴 업체 (yahoo, empas..)
		        response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=9999) then
		    	''브랜드 업체
		    	response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=9000) then
		    	''강사 업체
		    	response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=501) or (puseridv=502) or (puseridv=503) then
		    	response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    elseif (puseridv=101) or (puseridv=111) or (puseridv=112) or (puseridv=201) or (puseridv=301) then
		    	response.write "<script language='javascript'>alert('비밀번호가 변경되었습니다. 다시로그인 해 주세요.');top.location.replace('" & manageUrl & "/index.asp')</script>"
		        dbget.close()	:	response.End
		    else
		        session.Abandon
		        response.write "<script language='javascript'>alert('사용불가 계정입니다.');</script>"
		        dbget.close()	:	response.End
		    end if

	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.back();" &_
					"</script>"

	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->