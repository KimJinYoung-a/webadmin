<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incTenRedisSession.asp"-->

<%

GG_RDS_APIURL = "https://dapi.10x10.co.kr" ''"https://52.79.73.177"  
'GG_RDS_APIURL = "https://52.79.73.177"  
GG_RDS_AUTHKEY = "key=lEejpMoNDt1GYzODrwlcwMEDqidUkHdskioU7Tl3bdVeMXNFS13xJimboxKx"
'Call fn_RDS_SSN_SET
'response.end

''response.end
    Dim objXML
    Dim retJson, maykey, mayval
    Dim jsonObj, oJSONoutput

    Dim errrased : errrased = false
    Dim iredisKey : iredisKey = fn_RDS_SSN_KeyGet()
    if LEN(iredisKey)>=1 then 
        Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
        ''setTimeouts (long resolveTimeout, long connectTimeout, long sendTimeout, long receiveTimeout)
        objXML.open "GET", "" & GG_RDS_APIURL&"/api/RedisValues/"&iredisKey, False 
        objXML.SetTimeouts 10*1000, 10*1000, 10*1000, 10*1000
        objXML.setRequestHeader "Authorization", GG_RDS_AUTHKEY
        objXML.setRequestHeader "CONTENT-TYPE", "application/json"

        ''msxml3.dll (0x80072F06)
  ''      objXML.SetOption(2) = 13056 ''
        
        on Error Resume Next
        objXML.send()
        if Err THEN 
            errrased = True
            response.write "err:"&Err.Description

            
        End if
        On Error Goto 0

        if (NOT errrased) then
            response.write  objXML.Status &"<br>"
            if (objXML.Status = "200") then
                retJson = TRIM(objXML.responseText)
            else
                
            end if

            response.write retJson
        end if
        SET objXML = Nothing

    end if

%>