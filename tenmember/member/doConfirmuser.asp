<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  인트라넷 개인정보 접근 확인 처리
' History : 2018.08.29 허진원 생성
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
dim menupos, loginType, userid, empno, password, Enc_userpass, Enc_userpass64, sqlStr, chkLogin
menupos = requestCheckVar(request.form("menupos"),8)
loginType = requestCheckVar(request.form("logintype"),3)
password = requestCheckVar(request.form("password"),32)
chkLogin = false

userid = session("ssBctId")
empno = session("ssBctSn")

if password<>"" then
    Enc_userpass = md5(password)
    Enc_userpass64 = SHA256(Enc_userpass)
else
    Call Alert_return("잘못된 접근입니다.[E01]")
    dbget.Close: Response.End
end if

if loginType="id" and userid<>"" then
    '// ID로 계정 확인
    sqlStr = "Select top 1 id from db_partner.dbo.tbl_partner with (noLock) Where id='" & userid & "' and Enc_password64='" & Enc_userpass64 & "'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        chkLogin = true
    end if
    rsget.Close

elseif loginType="emp" and empno<>"" then
    '// 사번으로 계정 확인
    sqlStr = "Select top 1 userid from db_partner.dbo.tbl_user_tenbyten with (noLock) Where empno='" & empno & "' and Enc_emppass64='" & Enc_userpass64 & "'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        chkLogin = true
    end if
    rsget.Close

else
    Call Alert_return("잘못된 접근입니다.[E02]")
    dbget.Close: Response.End
end if

if chkLogin then
    session("chkSCMMyInfoPass") = now()
    response.write "<script>location.replace('" & getSCMSSLURL & "/tenmember/member/modify_myinfo.asp?menupos=" & menupos & "')</script>"
else
    response.write "<script>window.alert('아이디 또는 비밀번호가 맞지 않습니다.');history.go(-1);</script>"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->