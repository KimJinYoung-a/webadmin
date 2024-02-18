<% option Explicit %>
<!-- #include virtual="/mAppadmin/inc/incUTF8.asp" --><?xml version="1.0" encoding="UTF-8" ?>
<result>
<!-- #include virtual="/mAppadmin/inc/incCommon.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
''<!-- #include virtual="/lib/util/base64unicode.asp" -->
Dim conIp, arrIp, tmpIp
conIp = Request.ServerVariables("REMOTE_ADDR")
arrIp = split(conIp,".")
tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")

'// 텐바이텐 서버가 아니면 종료
if Not(tmpIp=>"192168000081" and tmpIp<="192168000090") and Not(tmpIp=>"061252133001" and tmpIp<="061252133127") and Not(tmpIp=>"110093128081" and tmpIp<="110093128099") then
	Response.Write "<retval>N</retval><error>잘못된 접속입니다.</error>" & vbCrLf
	Response.Write "</result>"
	dbget.Close(): Response.End
end if

'// 웹어드민 접속 로그 저장 함수
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

    ''USBTokenSn 길이제약 확인
    ''dbget.Execute sqlStr
end Sub

Dim glb_encuid,glb_tstp,glb_devicekey,glb_appid,glb_enckey
Dim userid, sqlStr
Dim retBUF

glb_encuid  = request("glb_encuid")
glb_tstp    = request("glb_tstp")
glb_devicekey= request("glb_devicekey")
glb_appid   = request("glb_appid")
glb_enckey  = request("glb_enckey")

userid = TENDEC(glb_encuid)

dim newTimeStamp
Dim isConfirmedUser : isConfirmedUser = FALSE
if (MD5(glb_devicekey&userid&glb_tstp)=glb_enckey) then
    sqlStr = " select top 1 * from db_AppNoti.dbo.tbl_tbtpns_register"
    sqlStr = sqlStr & " where userid='"&userid&"'"
    sqlStr = sqlStr & " and appid="&glb_appid&""
    sqlStr = sqlStr & " and regkey='"&glb_devicekey&"'"
    sqlStr = sqlStr & " and AuthDate is Not NULL"
    rsAppNotiget.Open sqlStr,dbAppNotiget,1
    if not rsAppNotiget.EOF  then
        isConfirmedUser = true
    end if
    rsAppNotiget.close

    if (Not isConfirmedUser) then
        retBUF = "<retval>N</retval><error>미인증 디바이스</error>"
    else
        sqlStr = "select top 1 A.id, A.company_name, A.userdiv, A.password " + vbCrlf
        sqlStr = sqlStr + "	, B.part_sn, A.level_sn, B.job_sn, B.username, B.posit_sn, IsNull(B.empno, '') as empno " + vbCrlf
       	sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner as A " + vbCrlf
        sqlStr = sqlStr & " join db_partner.dbo.tbl_user_tenbyten as B" + vbCrlf
        sqlStr = sqlStr & "     ON A.id = B.userid AND B.isUsing = 1" + vbCrlf

        ' 퇴사예정자 처리	' 2018.10.16 한용민
        sqlStr = sqlStr & "	    and (b.statediv ='Y' or (b.statediv ='N' and datediff(dd,b.retireday,getdate())<=0))" & vbcrlf
        sqlStr = sqlStr + " where A.id = '" + userid + "'" + vbCrlf
        sqlStr = sqlStr + " and A.isusing='Y'"
        sqlStr = sqlStr + " and (A.userdiv<9)"

        rsget.Open sqlStr,dbget,1
        if  not rsget.EOF  then
            retBUF = retBUF & "<retval>Y</retval>" & vbCrLf
            retBUF = retBUF & "<mAppBctId><![CDATA[" & userid & "]]></mAppBctId>" & vbCrLf
            retBUF = retBUF & "<mAppBctDiv><![CDATA[" & rsget("userdiv") & "]]></mAppBctDiv>" & vbCrLf
            retBUF = retBUF & "<mAppBctSn><![CDATA[" & rsget("empno") & "]]></mAppBctSn>" & vbCrLf
            retBUF = retBUF & "<mAppBctCname><![CDATA[" & db2html(rsget("company_name")) & "]]></mAppBctCname>" & vbCrLf
            retBUF = retBUF & "<mAppAdminPsn><![CDATA[" & rsget("part_sn") & "]]></mAppAdminPsn>" & vbCrLf
            retBUF = retBUF & "<mAppAdminLsn><![CDATA[" & rsget("level_sn") & "]]></mAppAdminLsn>" & vbCrLf
            retBUF = retBUF & "<mAppAdminPOsn><![CDATA[" & rsget("job_sn") & "]]></mAppAdminPOsn>" & vbCrLf
            retBUF = retBUF & "<mAppAdminPOSITsn><![CDATA[" & rsget("posit_sn") & "]]></mAppAdminPOSITsn>" & vbCrLf

            newTimeStamp =getTimeStampFormat
            retBUF = retBUF & "<mAppencuid><![CDATA[" & TENENC(userid) & "]]></mAppencuid>" & vbCrLf
            retBUF = retBUF & "<mApptstp><![CDATA[" & newTimeStamp & "]]></mApptstp>" & vbCrLf
            retBUF = retBUF & "<mAppenckey><![CDATA[" & MD5(glb_devicekey&userid&newTimeStamp) & "]]></mAppenckey>" & vbCrLf
        end if
        rsget.close

        if (retBUF<>"") then
            Call AddLoginLog (userid,"Y","AT:"&glb_devicekey)
        end if
    end if
else
    retBUF = "<retval>N</retval><error>암호화 오류</error>"
end if

response.write retBUF
%>
</result>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
