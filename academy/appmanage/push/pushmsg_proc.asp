<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	
	
	dim addParamMsg : addParamMsg=""  ''"param1":"value1","param2":"value2"
''    dim addCount : addCount = Request.Form("params").Count
''    
''	dim kk, iparam, iparamvalue
''	For kk = 1 To addCount
''		iparam = Request.Form("params")(kk)
''		iparamvalue = request.Form("paramvalue")(kk)
''
''		if (iparam<>"" and iparamvalue<>"") then
''			addParamMsg = addParamMsg & CHR(34)&iparam&CHR(34)&":"&CHR(34)&iparamvalue&CHR(34)
''			addParamMsg = addParamMsg & ","
''		end if
''
''	Next
''
''	if (addParamMsg<>"") then
''		addParamMsg = Left(addParamMsg,Len(addParamMsg)-1)
''	end if
    
    
	dim lastadminid : lastadminid = session("ssBctId")

	dim idoc_idx    : idoc_idx	= RequestCheckVar(request("idoc_idx"),10)
	dim stitle      : stitle	= RequestCheckVar(request("stitle"),150)
	dim subtitle    : subtitle	= RequestCheckVar(request("subtitle"),200)
	dim mode        : mode      = RequestCheckVar(request("mode"),20)
    dim appkey      : appkey    = RequestCheckVar(request("appkey"),10)
    dim deviceid    : deviceid  = RequestCheckVar(request("deviceid"),512)    
    dim testlecid   : testlecid = RequestCheckVar(request("testlecid"),32)
    dim multiPsKey  : multiPsKey = RequestCheckVar(request("multiPsKey"),10)
    dim targetgbn   : targetgbn = RequestCheckVar(request("targetgbn"),20)
    
    addParamMsg = CHR(34)&"url"&CHR(34)&":"&CHR(34)&subtitle&CHR(34)
    
	''pushimg = RequestCheckVar(request("pushimg"),200)
	

    ''' ios �޼���(json) �� �ѱ��̴� 256 ����Ʈ�� ���� �� ����
    '' pushtitle + pushurl ���̸� ���� <= 160 (169 ������ ������..)
    
    dim len1, len2, MaxValidLen
    dim sqlStr
    MaxValidLen = 1000 ''186 ''160  2016/11/09
    if (mode="testsendnoti") or (mode="realsendnoti") then
        sqlStr = " select datalength('"&stitle&"') as titleLen, datalength('"&subtitle&"') as urlLen"
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
		If not rsACADEMYget.EOF Then
			len1 = rsACADEMYget("titleLen")
			len2 = rsACADEMYget("urlLen")
		end if
		rsACADEMYget.close
		
		if (len1+len2)>MaxValidLen then
		    response.write "<script>alert('Ÿ��Ʋ ����+URL ���̰� "&MaxValidLen&" ����Ʈ�� �ʰ� �� �� �����ϴ�."&len1&"+"&len2&"="&len1+len2&"');history.back();</script>"
		    dbget.Close()
		    response.end
		end if
    end if
    
	Select Case mode

		Case "testsendnoti"
	        
			''response.write "appkey:"&appkey&"<br>"
			''response.write "deviceid:"&deviceid&"<br>"
			''response.write "stitle:"&stitle&"<br>"
			''response.write "addparams:"&addParamMsg&"<br>"

			if (appkey="") or (deviceid="") or (stitle="")then
				response.write "�ʼ� �� üũ ����"
				dbget.close():response.end
			end if

			if (addParamMsg<>"") then
				sqlStr = "exec [db_academy].[dbo].[sp_ACA_sendPushMsgWithParam_Artist] "&appkey&",'"&deviceid&"','"&stitle&"','"&addParamMsg&"','"&testlecid&"'"

				dbACADEMYget.Execute sqlStr
			else
				sqlStr = "exec [db_academy].[dbo].[sp_ACA_sendPushMsgSimple_Artist] "&appkey&",'"&deviceid&"','"&stitle&"'"

				dbACADEMYget.Execute sqlStr
			end If
			
			''response.write "==================================<br>"
			''response.write "�׽�Ʈ �޽��� �߼ۿ�û �Ǿ����ϴ�.<br>"

			dim sendedMsg:
			sqlStr = "select top 1 *"&VBCRLF
			sqlStr = sqlStr & " from [DBAPPPUSH].db_AppNoti.dbo.tbl_AppPushMsg_Academy_Artist" &VBCRLF
			sqlStr = sqlStr & " where appkey='"&appkey&"'"&VBCRLF
			sqlStr = sqlStr & " and deviceid='"&deviceid&"'"&VBCRLF
			sqlStr = sqlStr & " order by psKey desc"&VBCRLF

			rsACADEMYget.open sqlStr, dbACADEMYget, 1
			If not rsACADEMYget.EOF Then
				sendedMsg = rsACADEMYget("sendMsg")
			end if
			rsACADEMYget.close
			
			If sendedMsg <> "" Then 
				'//�׽�Ʈī��Ʈ �ø�
				''sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
				''sqlStr = sqlStr + " testpush = testpush + 1 " + VbCrlf
				''sqlStr = sqlStr + " where idx = "& idx &""
							
				'response.write sqlStr
				''dbget.Execute sqlStr
			End If 

			'Response.write "sendedMsg:"&sendedMsg
			''Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='�ݱ�'/>"
			''Response.end
            response.write "<script>alert('�׽�Ʈ �޽��� �߼ۿ�û �Ǿ����ϴ�. ');</script>"
        CASE "realsendnoti"
            
            if (targetgbn="") or (stitle="")then
				response.write "�ʼ� �� üũ ����"
				dbACADEMYget.close():response.end
			end if
			
			sqlStr = "exec [db_academy].[dbo].[sp_ACA_sendPushMsgNotice_Artist] "&idoc_idx&",'"&stitle&"','"&addParamMsg&"','"&targetgbn&"'"
			dbACADEMYget.Execute sqlStr
			
			response.write "<script>alert('���� PUSH �޽��� �߼ۿ�û �Ǿ����ϴ�. ');opener.location.reload();self.close();</script>"
				
        CASE ELSE
            response.write "<script>alert('���ǵ��� �ʾ��� "&mode&"');</script>"
			''Response.End 
        
	End Select


%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
