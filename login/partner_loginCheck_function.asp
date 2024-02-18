<%
function AddPartnerLoginLogWithGeoIpCode(iuserid,ilogtype,itokentype,igeoipCode)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")

    sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLogWithGeoIpCode '"&iuserid&"','"&Left(reFAddr,16)&"','"&ilogtype&"','"&itokentype&"','"&igeoipCode&"'"
    dbget.Execute sqlStr
end function

function fn_plogin_AddIISLOG(iAddLogs)
    ''addLog �߰� �α� //2016/12/29
    if (request.ServerVariables("QUERY_STRING")<>"") then iAddLogs="&"&iAddLogs
    response.AppendToLog iAddLogs
end function


function lastLoginPwdChgRegDiffDate(iUserid)
    dim sqlStr
    
    lastLoginPwdChgRegDiffDate =0
    
    sqlStr = " select datediff(d,isnull((CASE WHEN isNULL(lastlogindt,'2001-01-01')>isNULL(lastPwChgDT,'2001-01-01') THEN lastlogindt ELSE lastPwChgDT END),regdate),getdate()) as diffDT "
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner "
	sqlStr = sqlStr & " where id ='"&iUserid&"'"
	
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		lastLoginPwdChgRegDiffDate = rsget("diffDT")
	end if
	rsget.close
end function

function IsLongTimeNotLoginUserid(iUserid)
    IsLongTimeNotLoginUserid = (lastLoginPwdChgRegDiffDate(iUserid)>91)
end function
	

function IspartnerLoginRejectIP()
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")
    
    IspartnerLoginRejectIP = FALSE
    
    sqlStr = " select top 1 rejectExpDt from db_partner.[dbo].[tbl_partner_login_rejectIP]"&vbCRLF
    sqlStr = sqlStr&" where refip='"&reFAddr&"'"&vbCRLF
    sqlStr = sqlStr&" and rejectExpDt>getdate()"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		IspartnerLoginRejectIP = TRUE
	end if
	rsget.close
	
end function

function isCriticAuthCheckPartner(oUserid)
    dim iUserid
    iUserid = LCASE(oUserid)
    
    isCriticAuthCheckPartner= FALSE
    
    if (iUserid="mukti001") or (iUserid="alexami") then   ''���ۿ� �հ� ���µǾ�����..
        isCriticAuthCheckPartner = TRUE
        Exit function
    end if
    
    if (iUserid="temp2") or (iUserid="test2") then  ''test;
        isCriticAuthCheckPartner = TRUE
        Exit function
    end if
    
    
end function

''��Ʈ�� �α��ν� ������ �ʿ��� CASE
function IsPartnerAuthRequireIP(iUserid,iGeoCd,is2PwLogin)
    dim sqlStr, reFAddr, authIpExists, lastAuthdate
    reFAddr = request.ServerVariables("REMOTE_ADDR")
    
    IsPartnerAuthRequireIP = FALSE
    
    ''2fact �����̰� ����IP �̸� �ϴ� ���
    if (iGeoCd="KR") and (is2PwLogin) then    
        if (NOT isCriticAuthCheckPartner(iUserid)) THEN 
            Exit function
        End if
    End if
    
    '' ����IP�̰� 1Pw �����̰� , ���ݵ� ��ü�̸� �ϴ� ���. 2017/04/14
    if (iGeoCd="KR") and (NOT is2PwLogin) then
        if (IsBatchVendorAlowIP) then
            Exit function
        end if
    end if
    
    sqlStr = "select userid, convert(Varchar(10),lastAuthdate,21) as lastAuthdate from db_partner.dbo.tbl_partner_loginIP_Authed" &vbCRLF
    sqlStr = sqlStr & " where userid='"&iUserid&"'" &vbCRLF
    sqlStr = sqlStr & " and refip='"&reFAddr&"'"
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
	    authIpExists = True
		lastAuthdate = rsget("lastAuthdate")
	end if
	rsget.close
	
	'' �ش������ �����ϰ�, SMS ������ �޾����� ���.
	If (authIpExists) and (NOT isNULL(lastAuthdate)) then
	    Exit function
	end if
	
	'' �����̰� �ϴ� ����� ����IP �̸� �ϴ� ���.
    If (iGeoCd="KR") and (authIpExists) then
        Exit function
    end if
    
    IsPartnerAuthRequireIP = TRUE
    
end function

''��Ʈ�� �����Ϸ��� ������ �߰�
function AddPartnerAuthIpAdd(iUserid)
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")
    
    sqlStr = "db_partner.dbo.sp_Ten_Partner_LoginAuthIP_ADD '"&iuserid&"','"&Left(reFAddr,16)&"',1"
    dbget.Execute sqlStr
end function

''2���н����尡 �����ȵȰ�� �ñ�.
function Is2ndPwdNotExistsReject(iUserid)
    dim sqlStr
    Is2ndPwdNotExistsReject = False
    
    sqlStr = "select top 1 id, lastLoginDt from [db_partner].[dbo].tbl_partner "&vbCRLF
    sqlStr = sqlStr & " where id ='"&iUserid&"'"&vbCRLF
    sqlStr = sqlStr & " and isNULL(Enc_2password64,'')=''"&vbCRLF
    '''sqlStr = sqlStr & " and datediff(d,regdate,getdate())>31"   '' �ֱ�(1��)��ϵ� �귣��� �糦. => ������� ����.
    
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		Is2ndPwdNotExistsReject = TRUE
	end if
	rsget.close
    
end function

function getGeoIpCountryCode()
    dim geoip, reFAddr
    if (application("Svr_Info")="Dev") then
        getGeoIpCountryCode = "--"
        Exit function
    end if

    reFAddr = request.ServerVariables("REMOTE_ADDR")
    getGeoIpCountryCode = ""
    set geoip = Server.CreateObject("GeoIPCOM.GeoIP")
	geoip.loadDataFile("C:\GeoIP\GeoIP.dat")
	getGeoIpCountryCode = geoip.country_code_by_name(reFAddr)
	set geoip = nothing
end function

'' sabangnet �� ��ġ�� ó���ϴ� ��ü. ���α��ο��� ���.
function IsBatchVendorAlowIP()
    dim sqlStr, reFAddr
    reFAddr = request.ServerVariables("REMOTE_ADDR")
    
    IsBatchVendorAlowIP = FALSE
    sqlStr="select top 1 * from  db_partner.[dbo].[tbl_partner_login_BatchVendorIP]"&vbCRLF
    sqlStr = sqlStr & " where refip='"&reFAddr&"'"&vbCRLF
    sqlStr = sqlStr & " and validdate>getdate()"&vbCRLF
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
	if  not rsget.EOF  then
		IsBatchVendorAlowIP = TRUE
	end if
	rsget.close
	
end function

function getConSVCByUagentOrRefer()
    dim iuserAgnet : iuserAgnet = LCASE(Request.ServerVariables("HTTP_USER_AGENT"))
    dim ireferer : ireferer = LCASE(Request.ServerVariables("HTTP_REFERER"))
    
    dim retVal
    retVal = ""
    
    if Instr(ireferer,"sabangnet.co.kr")>0 then
        retVal = "SABAN"
    elseif Instr(ireferer,"erpia.net")>0 then
        retVal = "ERPIA"
    elseif Instr(ireferer,"next-engine.co.kr")>0 then    
        retVal = "NEENG"
    end if
    
    if (retVal="") then
        if Instr(iuserAgnet,"shopkeeper")>0 then 
            retVal ="SHOPK"
        end if
    end if
    
    getConSVCByUagentOrRefer = retVal  
end function

function IsPreLoginAvailREF()
    dim buf : buf = getConSVCByUagentOrRefer()
    
    IsPreLoginAvailREF= FALSE
    if (buf="SABAN") then
        IsPreLoginAvailREF = TRUE
    end if
end function

'// FrontApi ������Ű ������ - ssnHash�� ���� & ���
Function fnDBSessionCreateV2(frontId)
    dim ssnuserid  : ssnuserid =  frontId
    dim ssnlogindt : ssnlogindt = session("ssnlogindt")

    if (ssnuserid="") or (ssnlogindt="") then Exit function

    Dim ssnkeepAddtime : ssnkeepAddtime = 0
    Dim isessionData : isessionData = fnMakeSessionToDBData(ssnuserid)

    dim iSsnCon : set iSsnCon = CreateObject("ADODB.Connection")
    Dim cmd : set cmd = server.CreateObject("ADODB.Command")

    dim sqlStr
    sqlStr = "db_user.[dbo].[sp_TEN_SSN_CREATE_V2]"

    iSsnCon.Open Application("db_main") ''Ŀ�ؼ� ��Ʈ��.
    cmd.ActiveConnection = iSsnCon
    cmd.CommandText = sqlStr
    cmd.CommandType = adCmdStoredProc

    cmd.Parameters.Append cmd.CreateParameter("@ssnuserid", adVarchar, adParamInput, 32, ssnuserid)
    cmd.Parameters.Append cmd.CreateParameter("@ssnlogindt", adVarchar, adParamInput, 14, ssnlogindt)
    cmd.Parameters.Append cmd.CreateParameter("@lgnchannel", adVarchar, adParamInput, 1, "D")
    cmd.Parameters.Append cmd.CreateParameter("@ssnkeepAddtime", adInteger, adParamInput, , ssnkeepAddtime)
    cmd.Parameters.Append cmd.CreateParameter("@ssndata", adVarWChar, adParamInput, 384, isessionData)
    cmd.Parameters.Append cmd.CreateParameter("@retSsnHash", adVarchar, adParamOutput, 64, "")

    cmd.Execute
    Dim iretSsnHash : iretSsnHash = cmd.Parameters("@retSsnHash").Value
    fnDBSessionCreateV2 = iretSsnHash

    set cmd = Nothing
    iSsnCon.Close
    SET iSsnCon = Nothing

End Function
'// FrontApi ������Ű ������ - ssndata�� ����
Function fnMakeSessionToDBData(frontId)
    Dim retData
    Dim ispliter : ispliter = "||"
    retData = ""
    retData = retData & "ssnuserid=="&frontId&ispliter
    retData = retData & "ssnlogindt=="&session("ssnlogindt")&ispliter
    retData = retData & "ssnusername=="&session("ssBctCname")&ispliter
    retData = retData & "ssnuserdiv=="&session("ssBctDiv")&ispliter
    retData = retData & "ssnuserlevel=="&ChkIif(session("ssBctSn")<>"", "7", "9")&ispliter
	retData = retData & "ssnrealnamecheck==N"&ispliter
    retData = retData & "ssnuseremail=="&replace(session("ssBctEmail"),ispliter,"")&ispliter
    retData = retData & "ssnisAdult==N"&ispliter
    fnMakeSessionToDBData = retData
End Function

Function fnDateTimeToLongTime(icookieLoginDt)
    dim iorginDt : iorginDt = icookieLoginDt
    iorginDt = CDate(iorginDt)

    fnDateTimeToLongTime = Year(iorginDt) & Right("00" & Month(iorginDt),2) & Right("00" & Day(iorginDt),2) & Right("00" & Hour(iorginDt),2) & Right("00" & Minute(iorginDt),2) & Right("00" & Second(iorginDt),2)
End Function
%>