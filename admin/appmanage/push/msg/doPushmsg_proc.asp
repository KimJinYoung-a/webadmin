<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60*20		' 20분
%>
<%
'###########################################################
' Description : 예약 푸시 메시지 작성
' Hieditor : 서동석 생성
'			 2017.03.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim idx, reservationdate , state , mode , reservetime , reservemin , lastadminid, pushimg, pushimg2, pushimg3, pushimg4, pushimg5
Dim reservedate, gaparam, ppidx, makeridarr, itemidarr, keywordarr, bonuscouponidxarr, notclickyn, pushcontents, imgurl, imgurlcount
dim i, userid, yyyymmdd, utmparam, sendedMsg, useridarr, isMsgArr, ParamMsg, objItem, len1, len2, len3, MaxValidLen, stitle , subtitle, sqlStr
dim appkey, message, deviceid, addCount, addParamMsg, istargetMsg,admcomment, totcnt, kk, iparam, iparamvalue
dim noduppDate, noduppDate2, noduppDate3,targetKey,baseIdx, iADODBcmd, intResult, sendranking, replacetagcode, titletemp,contentstemp
dim CUSTOMERIDYN, CUSTOMERNAMEYN, CUSTOMERLEVELNAMEYN, CUSTOMERID,CUSTOMERNAME, CUSTOMERLEVELNAME, replacetagcodearray, replacetagcodetemp, privateYN
	useridarr = requestcheckvar(request("useridarr"),256)
	makeridarr = request("makeridarr")
	itemidarr = request("itemidarr")
	keywordarr = request("keywordarr")
	bonuscouponidxarr = request("bonuscouponidxarr")
	notclickyn = requestcheckvar(request("notclickyn"),10)
	appkey = requestcheckvar(request("appkey"),10)
	message = requestcheckvar(request("message"),800)
	deviceid = LEFT(replace(request("deviceid"),"'",""),200)  '' -- 이 치환 되면 안됨..
	idx				= RequestCheckVar(request("idx"),10)
	stitle			= RequestCheckVar(request("stitle"),800)
	pushcontents			= RequestCheckVar(request("pushcontents"),3000)
	subtitle		= RequestCheckVar(request("subtitle"),500)
	reservationdate = RequestCheckVar(request("reservationdate"),10)
	reservetime		= RequestCheckVar(request("time1"),2)
	reservemin		= RequestCheckVar(request("time2"),2)
	state			= RequestCheckVar(request("state"),2)
	istargetMsg     = RequestCheckVar(request("istargetMsg"),1)
	admcomment      = RequestCheckVar(request("admcomment"),200)
	noduppDate      = RequestCheckVar(request("noduppDate"),10)
	noduppDate2      = RequestCheckVar(request("noduppDate2"),10)
	noduppDate3      = RequestCheckVar(request("noduppDate3"),10)
	targetKey       = RequestCheckVar(request("targetKey"),10)
	sendranking       = RequestCheckVar(getNumeric(request("sendranking")),10)
	imgurl = request("imgurl")

imgurlcount=0
addCount = Request.Form("params").Count
addParamMsg=""  ''"param1":"value1","param2":"value2"
lastadminid = session("ssBctId")
privateYN="N"

For kk = 1 To addCount
	iparam = Request.Form("params")(kk)
	iparamvalue = request.Form("paramvalue")(kk)

	if (iparam<>"" and iparamvalue<>"") then
		addParamMsg = addParamMsg & CHR(34)&iparam&CHR(34)&":"&CHR(34)&iparamvalue&CHR(34)
		addParamMsg = addParamMsg & ","
	end if
Next

' 다중이미지 작업
imgurlcount = Request.Form("param2").Count
if imgurlcount>0 then
	addParamMsg = addParamMsg & CHR(34)&trim(Request.Form("param2")(1))&CHR(34)&":["
	For i = 1 To imgurlcount
		iparam = trim(Request.Form("param2")(i))
		iparamvalue = trim(request.Form("array-image-url")(i))

		if (iparam<>"" and iparamvalue<>"") then
			addParamMsg = addParamMsg &CHR(34)&iparamvalue&CHR(34)
			addParamMsg = addParamMsg & ","
		else
			response.write "<script type='text/javascript'>alert('이미지를 정확하게 입력해 주세요.');</script>"
			session.codePage = 949
			response.write "<script type='text/javascript'>history.back();</script>"
			dbget.Close() : response.end
		end if
	Next
	if (right(addParamMsg,1)=",") then
		addParamMsg = Left(addParamMsg,Len(addParamMsg)-1)
	end if
	addParamMsg = addParamMsg & "]"
end if

if (right(addParamMsg,1)=",") then
	addParamMsg = Left(addParamMsg,Len(addParamMsg)-1)
end if

reservedate		= reservationdate &" "& reservetime &":"& reservemin &":000" '예약일
'yyyymmdd = year(date()) & Format00(2,Month(date())) & Format00(2,day(date()))
yyyymmdd = replace(reservationdate,"-","")

if istargetMsg="" then istargetMsg="0"
if noduppDate="on" then noduppDate="1"
if noduppDate2="on" then noduppDate2="1"
if noduppDate3="on" then noduppDate3="1"
if notclickyn="on" then
	notclickyn="Y"
else
	notclickyn="N"
end if
if istargetMsg="0" then
    targetKey = ""
    baseIdx   = ""
end if
if sendranking="" then sendranking="6"
if (targetKey="") or (targetKey="1") then baseIdx   = ""

'Response.write reservedate
'Response.end

pushimg = RequestCheckVar(request("pushimg"),200)
pushimg2 = RequestCheckVar(request("pushimg2"),200)
pushimg3 = RequestCheckVar(request("pushimg3"),200)
pushimg4 = RequestCheckVar(request("pushimg4"),200)
pushimg5 = RequestCheckVar(request("pushimg5"),200)
mode = RequestCheckVar(request("mode"),32)

' PUSH메시지 크기
''' ios 메세지(json) 의 총길이는 256 바이트를 넘을 수 없음
'' pushtitle + pushurl 길이를 제한 <= 160 (169 까지는 나갔음..)
'MaxValidLen = 186 ''160  2016/11/09

' iOS 8 이상, Android 4 이상에서 전체 전송 제한 크기는 4Kb로 확인됨
'   다만 UTF8 한글 특성상 한글 한글자는 3byte가 할당되므로 1300자 정도 넣을 있으며
'   통신에 기타 해더및 추가 정보를 빼면 약 800자정도 넣을 수 있는것으로 판단됨
MaxValidLen = 800		' 2018.08.21
if (mode="mInsert") or (mode="mEdit") then
    sqlStr = " select len(N'"&stitle&"') as titleLen, len(N'"&subtitle&"') as urlLen, len(N'"& pushcontents &"') as pushcontents"

	'response.write sqlStr & "<br>"
    rsget.open sqlStr, dbget, 1
	If not rsget.EOF Then
		len1 = rsget("titleLen")
		len2 = rsget("urlLen")
		len3 = rsget("pushcontents")
	end if
	rsget.close
	
	if (len1+len2+len3)>MaxValidLen then
	    'response.write "<script type='text/javascript'>alert('타이틀 길이+URL 길이가 "&MaxValidLen&" 바이트를 초과 할 수 없습니다."&len1&"+"&len2&"="&len1+len2&"');history.back();</script>"
		response.write "<script type='text/javascript'>alert('타이틀 길이+URL 길이가 "& MaxValidLen &"글자를 초과 할 수 없습니다."&len1&"+"&len2&"+"&len3&"="&len1+len2+len3&"');</script>"
		session.codePage = 949
		response.write "<script type='text/javascript'>history.back();</script>"
	    dbget.Close() : response.end
	end if

	' if instr(subtitle,".asp") < 1 then
	' 	response.write "<script type='text/javascript'>alert('푸시메세지 링크에 .asp 가 없습니다.');</script>"
	' 	session.codePage = 949
	' 	response.write "<script type='text/javascript'>history.back();</script>"
	' 	dbget.Close() : response.end
	' end if
	' if instr(subtitle,"?") > 0 or instr(subtitle,"&") > 0 then
	' 	if instr(subtitle,".asp?") < 1 and instr(subtitle,"/?") < 1 then
	' 		response.write "<script type='text/javascript'>alert('푸시메세지 링크 형식이 잘못되었습니다[0]');</script>"
	' 		session.codePage = 949
	' 		response.write "<script type='text/javascript'>history.back();</script>"
	' 		dbget.Close() : response.end
	' 	end If
	' end if
	if instr(subtitle,"이벤트번호") > 0 or instr(subtitle,"이벤트코드") > 0 or instr(subtitle,"상품번호") > 0 or instr(subtitle,"상품코드") > 0 then
		response.write "<script type='text/javascript'>alert('푸시메세지 링크에 한글로된 잘못된 형식이 있습니다.');</script>"
		session.codePage = 949
		response.write "<script type='text/javascript'>history.back();</script>"
	    dbget.Close() : response.end
	end if
end if

Select Case mode
	Case "mInsert"
		sqlStr = " select count(*) as cnt from db_contents.dbo.tbl_app_push_reserve where isusing = 'Y' and istargetMsg=0 and reservedate = '" & reservedate & "'"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			totcnt = rsget("cnt")
		end if
		rsget.close
	 
		'/ 해당시간 등록 푸시가 없거나, 수기 푸시면 묻지도 따지지도 말고 등록되게
		If totcnt = "0" or targetKey="9999" then	 
			if makeridarr <> "" then
				if checkNotValidHTML(makeridarr) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('브랜드ID에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
				end if
			end if
			if itemidarr <> "" then
				if checkNotValidHTML(itemidarr) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('상품코드에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
				end if
			end if
			if keywordarr <> "" then
				if checkNotValidHTML(keywordarr) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('키워드에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
				end if
			end if
			if bonuscouponidxarr <> "" then
				if checkNotValidHTML(bonuscouponidxarr) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('보너스쿠폰번호에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
				end if
			end if

			if trim(makeridarr)<>"" then
				makeridarr = replace(replace(replace(trim(makeridarr),".",","),"""",""),"'","")		'오타 교정
				if left(makeridarr,1)="," then makeridarr = mid(makeridarr,2,len(makeridarr))		'오타 교정
				if right(makeridarr,1)="," then makeridarr = left(makeridarr,len(makeridarr)-1)		'오타 교정
				makeridarr = """" & replace(makeridarr,",",""",""") & """"		'디비에 우선 브랜드 사이에 "," 넣어서 때려 넣음. 받는 부분에서 replace(makeridarr,"""","'") 해야함
			end if
			if trim(itemidarr)<>"" then
				itemidarr = replace(replace(replace(trim(itemidarr),".",","),"""",""),"'","")		'오타 교정
				if left(itemidarr,1)="," then itemidarr = mid(itemidarr,2,len(itemidarr))		'오타 교정
				if right(itemidarr,1)="," then itemidarr = left(itemidarr,len(itemidarr)-1)		'오타 교정
			end if
			if trim(keywordarr)<>"" then
				keywordarr = replace(replace(replace(trim(keywordarr),".",","),"""",""),"'","")		'오타 교정
				if left(keywordarr,1)="," then keywordarr = mid(keywordarr,2,len(keywordarr))		'오타 교정
				if right(keywordarr,1)="," then keywordarr = left(keywordarr,len(keywordarr)-1)		'오타 교정
				keywordarr = """" & replace(keywordarr,",",""",""") & """"		'디비에 우선 키워드 사이에 "," 넣어서 때려 넣음. 받는 부분에서 replace(keywordarr,"""","'") 해야함
			end if
			if trim(bonuscouponidxarr)<>"" then
				bonuscouponidxarr = replace(replace(replace(trim(bonuscouponidxarr),".",","),"""",""),"'","")		'오타 교정
				if left(bonuscouponidxarr,1)="," then bonuscouponidxarr = mid(bonuscouponidxarr,2,len(bonuscouponidxarr))		'오타 교정
				if right(bonuscouponidxarr,1)="," then bonuscouponidxarr = left(bonuscouponidxarr,len(bonuscouponidxarr)-1)		'오타 교정
			end if
			stitle = trim(stitle)
			pushcontents = trim(pushcontents)

			if stitle="" or isnull(stitle) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목을 입력해 주세요.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			if pushcontents="" or isnull(pushcontents) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('내용을 입력해 주세요.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			if stitle<>"" then
				stitle = replace(stitle,vbcrlf,"")

				if checkNotValidHTML(stitle) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
			end if
			if pushcontents<>"" then
				pushcontents = replace(pushcontents,vbcrlf,"\n")

				if checkNotValidHTML(pushcontents) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
			end if

			' 타켓푸시의 경우 사용가능한 치환코드를 가져온다.
			if istargetMsg="1" then
				replacetagcode=""
				sqlStr = "SELECT" & vbcrlf
				sqlStr = sqlStr & " q.targetKey,q.targetName,q.targetQuery,q.isusing,q.repeatpushyn, q.target_procedureyn, q.replacetagcode" & vbcrlf
				sqlStr = sqlStr & " From db_contents.[dbo].[tbl_app_targetQuery] q with (readuncommitted)" & vbcrlf
				sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
				sqlStr = sqlStr & " and q.targetkey=N'"& targetkey &"'" & vbcrlf

				'response.write sqlStr &"<br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				IF not rsget.EOF THEN
					replacetagcode 	= trim(db2html(rsget("replacetagcode")))
				End IF			
				rsget.Close	

			' 전체푸시의 경우 아이디와 이름만 사용가능
			else
				replacetagcode="${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}"
			end if

			titletemp = stitle
			contentstemp = pushcontents

			replacetagcodearray=""
			replacetagcodetemp=""
			if replacetagcode<>"" and not(isnull(replacetagcode)) then
				replacetagcodearray = split(replacetagcode,",")

				if isarray(replacetagcodearray) then
					for i = 0  to ubound(replacetagcodearray)
						replacetagcodetemp = trim(replacetagcodearray(i))
						if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
							titletemp = replace(titletemp,replacetagcodetemp,"")
							contentstemp = replace(contentstemp,replacetagcodetemp,"")
						end if
					next
				end if
			end if

			if instr(titletemp,"${")>0 or instr(contentstemp,"${")>0 then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if

			privateYN="N"
			if instr(stitle,"${")>0 or instr(pushcontents,"${")>0 or instr(subtitle,"${")>0 then
				privateYN="Y"
			end if

			sqlStr = " insert into db_contents.dbo.tbl_app_push_reserve (" + VbCrlf
			sqlStr = sqlStr & " pushtitle , pushurl , state , reservedate, pushimg, istargetMsg, admcomment, noduppDate, noduppDate2, noduppDate3, targetKey, baseIdx" & VbCrlf
			sqlStr = sqlStr & " , makeridarr, itemidarr, keywordarr, bonuscouponidxarr, notclickyn, regadminid, lastadminid, pushcontents, pushimg2, pushimg3, pushimg4" & VbCrlf
			sqlStr = sqlStr & " , pushimg5, sendranking, privateYN) values (" & VbCrlf
			sqlStr = sqlStr + " N'" + stitle + "' ,N'" + subtitle + "' ," + state + " ,'"& reservedate &"', N'"& pushimg &"',"&istargetMsg&", N'"&html2db(admcomment)&"'" + VbCrlf
			sqlStr = sqlStr + " ,"&CHKIIF(noduppDate<>"1","NULL",noduppDate) & "," & CHKIIF(noduppDate2<>"1","NULL",noduppDate2) & "," & CHKIIF(noduppDate3<>"1","NULL",noduppDate3)+ VbCrlf
			sqlStr = sqlStr + " ,"&CHKIIF(targetKey="","NULL",targetKey)&" ,"&CHKIIF(baseIdx="","NULL",baseIdx) &", N'"& makeridarr &"', N'"& itemidarr &"'" & VbCrlf
			sqlStr = sqlStr & " , N'"& keywordarr &"', N'"& bonuscouponidxarr &"', '"& notclickyn &"', N'"& lastadminid &"', N'"& lastadminid &"',N'" + pushcontents + "'" & VbCrlf
			sqlStr = sqlStr & " , N'"& pushimg2 &"', N'"& pushimg3 &"', N'"& pushimg4 &"', N'"& pushimg5 &"', "& sendranking &",'"& privateYN &"'" & VbCrlf
			sqlStr = sqlStr + " )" + VbCrlf

			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

            '' gaparam 추가 2016/11/09 ------------------------------------------------------------------
		    if InStr(subtitle,"gaparam")<1 then
		        gaparam = ""

		        sqlStr = "select top 1 idx, isNULL(targetKey,0) as targetKey from db_contents.dbo.tbl_app_push_reserve where pushurl='"&subtitle&"' order by idx desc"

				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    			If not rsget.EOF Then
    			    ppidx = rsget("idx")
    				gaparam = "gaparam=push_"&ppidx&"_"&rsget("targetKey")
					'utmparam = "utm_source=10x10&utm_medium=push&utm_campaign="& yyyymmdd &"_"&ppidx&"_"&rsget("targetKey")
    			end if
    			rsget.close

		        if (gaparam<>"") then
		            if InStr(subtitle,"?")>0 then
						gaparam = "&"&gaparam
		                'gaparam = "&"&gaparam&"&"&utmparam
		            else
		                gaparam = "?"&gaparam
						'gaparam = "?"&gaparam&"&"&utmparam
		            end if
		            
		            sqlStr = " update db_contents.dbo.tbl_app_push_reserve" + VbCrlf
		            sqlStr = sqlStr + " set pushurl=N'"&subtitle&gaparam&"'"  + VbCrlf
					sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
					sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
		            sqlStr = sqlStr + " idx="&ppidx
		            
		            dbget.Execute sqlStr
		        end if
		    end if

		'	dim referer
		'	referer = request.ServerVariables("HTTP_REFERER")
			response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
			session.codePage = 949
			Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
			dbget.close()	:	response.End
		Else
			response.write "<script type='text/javascript'>alert('같은시간에 전체 발송 메세지가 이미 존재합니다.');</script>"
			session.codePage = 949
			Response.write "<script type='text/javascript'>history.back(-1);</script>"
			dbget.close()	:	response.End
		End If 

	' 발송 선택기기 1개
	Case "test_insert"
		response.write "appkey:"&appkey&"<br>"
		response.write "deviceid:"&deviceid&"<br>"
		response.write "message:"&message&"<br>"
		response.write "pushcontents:"&pushcontents&"<br>"
		response.write "addparams:"&addParamMsg&"<br>"

		if (appkey="") or (deviceid="") then
			response.write "필수 값 체크 오류"
			session.codePage = 949
			dbget.close():response.end
		end if
		message = trim(message)
		pushcontents = trim(pushcontents)

		if message="" or isnull(message) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if pushcontents="" or isnull(pushcontents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('내용을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if message<>"" then
			message = replace(message,vbcrlf,"")

			if checkNotValidHTML(message) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		if pushcontents<>"" then
			pushcontents = replace(pushcontents,vbcrlf,"\n")

			if checkNotValidHTML(pushcontents) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if

		CUSTOMERIDYN="N"
		CUSTOMERNAMEYN="N"
		CUSTOMERLEVELNAMEYN="N"

		' 실제 고객의 데이터로 치환을 위한 체크
		if replace(message,"${CUSTOMERID}","")<>message then CUSTOMERIDYN="Y"
		if replace(pushcontents,"${CUSTOMERID}","")<>pushcontents then CUSTOMERIDYN="Y"
		if replace(message,"${CUSTOMERNAME}","")<>message then CUSTOMERNAMEYN="Y"
		if replace(pushcontents,"${CUSTOMERNAME}","")<>pushcontents then CUSTOMERNAMEYN="Y"
		if replace(message,"${CUSTOMERLEVELNAME}","")<>message then CUSTOMERLEVELNAMEYN="Y"
		if replace(pushcontents,"${CUSTOMERLEVELNAME}","")<>pushcontents then CUSTOMERLEVELNAMEYN="Y"

		'if len(session("ssBctId"))>1 then
		'	CUSTOMERID=LEFT(session("ssBctId"),LEN(session("ssBctId"))-1)&"*"
		'else
			CUSTOMERID=session("ssBctId")
		'end if
		'if len(session("ssBctCname"))>1 then
		'	CUSTOMERNAME=LEFT(session("ssBctCname"),LEN(session("ssBctCname"))-1)&"*"
		'else
			CUSTOMERNAME=session("ssBctCname")
		'end if
		' 실제 고객의 데이터로 치환
		if CUSTOMERIDYN="Y" or CUSTOMERNAMEYN="Y" then
			message = replace(message,"${CUSTOMERID}",CUSTOMERID)
			pushcontents = replace(pushcontents,"${CUSTOMERID}",CUSTOMERID)
			message = replace(message,"${CUSTOMERNAME}",CUSTOMERNAME)
			pushcontents = replace(pushcontents,"${CUSTOMERNAME}",CUSTOMERNAME)
		end if
		if CUSTOMERLEVELNAMEYN="Y" then
			sqlStr = "SELECT"
			sqlStr = sqlStr & " isnull((case when L.userlevel = 0 then 'WHITE'"
			sqlStr = sqlStr & " 	when L.userlevel = 1 then 'RED'"
			sqlStr = sqlStr & " 	when L.userlevel = 2 then 'VIP'"
			sqlStr = sqlStr & " 	when L.userlevel = 3 then 'VIP GOLD'"
			sqlStr = sqlStr & " 	when L.userlevel = 4 then 'VVIP'"
			sqlStr = sqlStr & " 	when L.userlevel = 7 then 'STAFF'"
			sqlStr = sqlStr & " 	when L.userlevel = 8 then 'FAMILY'"
			sqlStr = sqlStr & " 	when L.userlevel = 9 then 'BIZ'"
			sqlStr = sqlStr & " end),'비회원') as userlevelname"
			sqlStr = sqlStr & " From db_user.dbo.tbl_logindata l with (readuncommitted)"
			sqlStr = sqlStr & " WHERE l.userid='"& session("ssBctId") &"'"

			'response.write sqlStr &"<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				CUSTOMERLEVELNAME = trim(db2html(rsget("userlevelname")))
			else
				CUSTOMERLEVELNAME = "비회원"
			End IF			
			rsget.Close
			
			message = replace(message,"${CUSTOMERLEVELNAME}",CUSTOMERLEVELNAME)
			pushcontents = replace(pushcontents,"${CUSTOMERLEVELNAME}",CUSTOMERLEVELNAME)
		end if

		if (addParamMsg<>"") then
			sqlStr = "exec [db_contents].[dbo].[sp_Ten_sendPushMsgWithParam] "&appkey&",'"&deviceid&"',N'"&message&"','"&addParamMsg&"','"& session("ssBctId") &"',N'"& pushcontents &"',N'Y'"

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr
		else
			sqlStr = "exec [db_contents].[dbo].[sp_Ten_sendPushMsgSimple] "&appkey&",'"&deviceid&"',N'"&message&"',N'"& pushcontents &"',N'Y'"

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr
		end If
		
		response.write "==================================<br>"
		response.write "테스트 메시지 발송요청 되었습니다.<br>"

		sqlStr = "select top 1 *"&VBCRLF
		sqlStr = sqlStr & " from [DBAPPPUSH].db_AppNoti.dbo.tbl_AppPushMsg" &VBCRLF
		sqlStr = sqlStr & " where appkey='"&appkey&"'"&VBCRLF
		sqlStr = sqlStr & " and deviceid='"&deviceid&"'"&VBCRLF
		sqlStr = sqlStr & " order by psKey desc"&VBCRLF

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			sendedMsg = rsget("sendMsg")
		end if
		rsget.close
		
		If sendedMsg <> "" Then 
			'//테스트카운트 올림
			sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
			sqlStr = sqlStr + " testpush = testpush + 1 " + VbCrlf
			sqlStr = sqlStr + " where idx = "& idx &""
						
			'response.write sqlStr
			dbget.Execute sqlStr
		End If 

		'Response.write "sendedMsg:"&sendedMsg
		Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='닫기'/>"
		session.codePage = 949
		dbget.close():response.end

	' 발송 등록된 전체기기
	Case "test_allinsert"
		response.write "userid:"& useridarr &"<br>"
		response.write "message:"&message&"<br>"
		response.write "pushcontents:"&pushcontents&"<br>"
		response.write "addparams:"&addParamMsg&"<br>"

		if (useridarr="") then
			response.write "필수 값 체크 오류"
			session.codePage = 949
			dbget.close():response.end
		end if

		message = trim(message)
		pushcontents = trim(pushcontents)

		if message="" or isnull(message) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if pushcontents="" or isnull(pushcontents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('내용을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if message<>"" then
			message = replace(message,vbcrlf,"")

			if checkNotValidHTML(message) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		if pushcontents<>"" then
			pushcontents = replace(pushcontents,vbcrlf,"\n")

			if checkNotValidHTML(pushcontents) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		useridarr = replace(useridarr,"'","")     
		useridarr = replace(useridarr,"""","")
		useridarr = replace(useridarr,".",",")
		useridarr = replace(useridarr,",,",",")

		useridarr = "'" & replace(useridarr,",","','") & "'"

		IF LEFT(message,1)="{" and RIGHT(message,1)="}" then
			isMsgArr = 1
		ELSE
			isMsgArr = 0
		end if

		CUSTOMERIDYN="N"
		CUSTOMERNAMEYN="N"
		CUSTOMERLEVELNAMEYN="N"

		' 실제 고객의 데이터로 치환을 위한 체크
		if replace(message,"${CUSTOMERID}","")<>message then CUSTOMERIDYN="Y"
		if replace(pushcontents,"${CUSTOMERID}","")<>pushcontents then CUSTOMERIDYN="Y"
		if replace(message,"${CUSTOMERNAME}","")<>message then CUSTOMERNAMEYN="Y"
		if replace(pushcontents,"${CUSTOMERNAME}","")<>pushcontents then CUSTOMERNAMEYN="Y"
		if replace(message,"${CUSTOMERLEVELNAME}","")<>message then CUSTOMERLEVELNAMEYN="Y"
		if replace(pushcontents,"${CUSTOMERLEVELNAME}","")<>pushcontents then CUSTOMERLEVELNAMEYN="Y"

		sqlStr = "select"
		sqlStr = sqlStr & " u.userid, u.username"
		sqlStr = sqlStr & " , isnull((case when L.userlevel = 0 then 'WHITE'"
		sqlStr = sqlStr & " 	when L.userlevel = 1 then 'RED'"
		sqlStr = sqlStr & " 	when L.userlevel = 2 then 'VIP'"
		sqlStr = sqlStr & " 	when L.userlevel = 3 then 'VIP GOLD'"
		sqlStr = sqlStr & " 	when L.userlevel = 4 then 'VVIP'"
		sqlStr = sqlStr & " 	when L.userlevel = 7 then 'STAFF'"
		sqlStr = sqlStr & " 	when L.userlevel = 8 then 'FAMILY'"
		sqlStr = sqlStr & " 	when L.userlevel = 9 then 'BIZ'"
		sqlStr = sqlStr & " end),'비회원') as userlevelname"
		sqlStr = sqlStr & " into #tmpuser"
		sqlStr = sqlStr & " from db_user.dbo.tbl_user_n u with (nolock)"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_logindata l with (nolock)"
		sqlStr = sqlStr & " 	on u.userid=l.userid"
		sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

		sqlStr = "insert into [DBAPPPUSH].db_AppNoti.dbo.tbl_AppPushMsg_NoLock (appkey,multiPsKey,sendState,deviceid,sendMsg,userid,targetKey, repeatpushyn, repeatidx, sendranking, testpushyn) " & vbcrlf
		sqlStr = sqlStr & " SELECT " & vbcrlf
		sqlStr = sqlStr & " a.appKey, a.multiPsKey, a.sendState, a.deviceid" & vbcrlf
		sqlStr = sqlStr & " , convert(nvarchar(max)," & vbcrlf
		sqlStr = sqlStr & " 	replace(" & vbcrlf
		sqlStr = sqlStr & " 		replace(" & vbcrlf
		'sqlStr = sqlStr & " 			replace(msg1 + msg2 + msg3,'${CUSTOMERID}',(CASE when isnull(a.userid,'')='' then '고객' WHEN LEN(isnull(a.userid,''))>1 THEN LEFT(isnull(a.userid,''),LEN(isnull(a.userid,''))-1)+N'*' ELSE isnull(a.userid,'') END))" & vbcrlf
		sqlStr = sqlStr & " 			replace(msg1 + msg2 + msg3,'${CUSTOMERID}',(CASE when isnull(a.userid,'')='' then '고객' ELSE isnull(a.userid,'') END))" & vbcrlf
		'sqlStr = sqlStr & " 		,'${CUSTOMERNAME}',(CASE when isnull(a.username,'')='' then '고객' WHEN LEN(isnull(a.username,''))>1 THEN LEFT(isnull(a.username,''),LEN(isnull(a.username,''))-1)+N'*' ELSE isnull(a.username,'') END))" & vbcrlf
		sqlStr = sqlStr & " 		,'${CUSTOMERNAME}',(CASE when isnull(a.username,'')='' then '고객' ELSE isnull(a.username,'') END))" & vbcrlf
		sqlStr = sqlStr & " 	,'${CUSTOMERLEVELNAME}',isnull(a.userlevelname,'비회원'))" & vbcrlf
		sqlStr = sqlStr & " ) AS sendMsg" & vbcrlf
		sqlStr = sqlStr & " , a.userid, NULL, 'N', NULL, 3, N'Y'" & vbcrlf		' , a.regIdx 
		sqlStr = sqlStr & " FROM ( " & vbcrlf
		sqlStr = sqlStr & " 	SELECT " & vbcrlf
		sqlStr = sqlStr & " 	r.appKey " & vbcrlf
		sqlStr = sqlStr & " 	, 0 AS multiPsKey " & vbcrlf
		sqlStr = sqlStr & " 	, 0 AS sendState, r.deviceid " & vbcrlf

		if isMsgArr=1 then
			sqlStr = sqlStr & " 	, N'{""title"":"& message &",""noti"":"& pushcontents &"' as msg1 " & vbcrlf
		else
			sqlStr = sqlStr & " 	, N'{""title"":"""& replace(message,"""","\""") &""",""noti"":"""& replace(pushcontents,"""","\""") &"""' as msg1 " & vbcrlf
		end if

		sqlStr = sqlStr & " 	, N',' + N'"& addParamMsg &"' AS msg2 " & vbcrlf
		sqlStr = sqlStr & " 	, (CASE WHEN appKey = 5 THEN N',""did"":""'+ r.deviceid +N'""' else N'' end) + N'}' as msg3 " & vbcrlf
		sqlStr = sqlStr & " 	, T.userid, R.regIdx, R.isusing, R.pushyn " & vbcrlf
		sqlStr = sqlStr & " 	, ROW_NUMBER() OVER (PARTITION BY T.userid ORDER BY ISNULL(R.lastupdate,R.regdate) DESC) device_rnk " & vbcrlf
		sqlStr = sqlStr & " 	, t.username, t.userlevelname" & vbcrlf
		sqlStr = sqlStr & " 	FROM #tmpuser T " & vbcrlf
		sqlStr = sqlStr & " 	JOIN db_contents.dbo.tbl_app_regInfo R " & vbcrlf
		sqlStr = sqlStr & " 		on T.userid=R.userid " & vbcrlf
		sqlStr = sqlStr & " 	WHERE 1=1 " & vbcrlf
		'sqlStr = sqlStr & " 	and ISNULL(pushyn,'')<>'N' " & vbcrlf
		sqlStr = sqlStr & " 	AND R.isusing='Y' " & vbcrlf
		sqlStr = sqlStr & " 	AND ((R.appkey=6 and R.appVer>='36') or (R.appkey=5 and R.appVer>='1')) " & vbcrlf
		sqlStr = sqlStr & " )  as a " & vbcrlf
		'sqlStr = sqlStr & " WHERE device_rnk = 1 " & vbcrlf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

		response.write "==================================<br>"
		response.write "테스트 메시지 발송요청 되었습니다.<br>"

		'//테스트카운트 올림
		sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
		sqlStr = sqlStr + " testpush = testpush + "& ubound(split(useridarr,","))+1 &" " + VbCrlf
		sqlStr = sqlStr + " where idx = "& idx &""
					
		'response.write sqlStr
		dbget.Execute sqlStr

		'Response.write "sendedMsg:"&sendedMsg
		Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='닫기'/>"
		Response.end

	Case "mEdit"
		if makeridarr <> "" then
			if checkNotValidHTML(makeridarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('브랜드ID에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if
		if itemidarr <> "" then
			if checkNotValidHTML(itemidarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('상품코드에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if
		if keywordarr <> "" then
			if checkNotValidHTML(keywordarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('키워드에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if
		if bonuscouponidxarr <> "" then
			if checkNotValidHTML(bonuscouponidxarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('보너스쿠폰번호에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if

		if trim(makeridarr)<>"" then
			makeridarr = replace(replace(replace(trim(makeridarr),".",","),"""",""),"'","")		'오타 교정
			if left(makeridarr,1)="," then makeridarr = mid(makeridarr,2,len(makeridarr))		'오타 교정
			if right(makeridarr,1)="," then makeridarr = left(makeridarr,len(makeridarr)-1)		'오타 교정
			makeridarr = """" & replace(makeridarr,",",""",""") & """"		'디비에 우선 브랜드 사이에 "," 넣어서 때려 넣음. 받는 부분에서 replace(makeridarr,"""","'") 해야함
		end if
		if trim(itemidarr)<>"" then
			itemidarr = replace(replace(replace(trim(itemidarr),".",","),"""",""),"'","")		'오타 교정
			if left(itemidarr,1)="," then itemidarr = mid(itemidarr,2,len(itemidarr))		'오타 교정
			if right(itemidarr,1)="," then itemidarr = left(itemidarr,len(itemidarr)-1)		'오타 교정
		end if
		if trim(keywordarr)<>"" then
			keywordarr = replace(replace(replace(trim(keywordarr),".",","),"""",""),"'","")		'오타 교정
			if left(keywordarr,1)="," then keywordarr = mid(keywordarr,2,len(keywordarr))		'오타 교정
			if right(keywordarr,1)="," then keywordarr = left(keywordarr,len(keywordarr)-1)		'오타 교정
			keywordarr = """" & replace(keywordarr,",",""",""") & """"		'디비에 우선 키워드 사이에 "," 넣어서 때려 넣음. 받는 부분에서 replace(keywordarr,"""","'") 해야함
		end if
		if trim(bonuscouponidxarr)<>"" then
			bonuscouponidxarr = replace(replace(replace(trim(bonuscouponidxarr),".",","),"""",""),"'","")		'오타 교정
			if left(bonuscouponidxarr,1)="," then bonuscouponidxarr = mid(bonuscouponidxarr,2,len(bonuscouponidxarr))		'오타 교정
			if right(bonuscouponidxarr,1)="," then bonuscouponidxarr = left(bonuscouponidxarr,len(bonuscouponidxarr)-1)		'오타 교정
		end if
		stitle = trim(stitle)
		pushcontents = trim(pushcontents)

		if stitle="" or isnull(stitle) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if pushcontents="" or isnull(pushcontents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('내용을 입력해 주세요.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
		if stitle<>"" then
			stitle = replace(stitle,vbcrlf,"")

			if checkNotValidHTML(stitle) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		if pushcontents<>"" then
			pushcontents = replace(pushcontents,vbcrlf,"\n")

			if checkNotValidHTML(pushcontents) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if

		' 타켓푸시의 경우 사용가능한 치환코드를 가져온다.
		if istargetMsg="1" then
			replacetagcode=""
			sqlStr = "SELECT" & vbcrlf
			sqlStr = sqlStr & " q.targetKey,q.targetName,q.targetQuery,q.isusing,q.repeatpushyn, q.target_procedureyn, q.replacetagcode" & vbcrlf
			sqlStr = sqlStr & " From db_contents.[dbo].[tbl_app_targetQuery] q with (readuncommitted)" & vbcrlf
			sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
			sqlStr = sqlStr & " and q.targetkey=N'"& targetkey &"'" & vbcrlf

			'response.write sqlStr &"<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				replacetagcode 	= trim(db2html(rsget("replacetagcode")))
			End IF			
			rsget.Close	

		' 전체푸시의 경우 아이디와 이름만 사용가능
		else
			replacetagcode="${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}"
		end if

		titletemp = stitle
		contentstemp = pushcontents

		replacetagcodearray=""
		replacetagcodetemp=""
		if replacetagcode<>"" and not(isnull(replacetagcode)) then
			replacetagcodearray = split(replacetagcode,",")

			if isarray(replacetagcodearray) then
				for i = 0  to ubound(replacetagcodearray)
					replacetagcodetemp = trim(replacetagcodearray(i))
					if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
						titletemp = replace(titletemp,replacetagcodetemp,"")
						contentstemp = replace(contentstemp,replacetagcodetemp,"")
					end if
				next
			end if
		end if

		if instr(titletemp,"${")>0 or instr(contentstemp,"${")>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		privateYN="N"
		if instr(stitle,"${")>0 or instr(pushcontents,"${")>0 or instr(subtitle,"${")>0 then
			privateYN="Y"
		end if

       '' 타겟 대상이 바뀌면 재타겟을 해야함..
        dim PreTargetKey
        sqlStr = "select targetKey from db_contents.dbo.tbl_app_push_reserve"
        sqlStr = sqlStr + " where idx = "& idx

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			PreTargetKey = rsget("targetKey")
			if isNULL(PreTargetKey) then PreTargetKey=""
		end if
		rsget.close
		
		if (CStr(PreTargetKey)<>CStr(targetKey)) then
		    rw PreTargetKey&","&targetKey
		    sqlStr = "delete from db_contents.dbo.tbl_app_push_TargetTemp" + VbCrlf
            sqlStr = sqlStr & "where rsvIdx="& idx
            dbget.Execute sqlStr
		end if

		sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
		sqlStr = sqlStr + " pushtitle = N'"& stitle &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushurl = N'"& subtitle &"'" + VbCrlf
		sqlStr = sqlStr + " ,state = "& state &"" + VbCrlf
		sqlStr = sqlStr + " ,reservedate = '"& reservedate &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg = N'"& pushimg &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg2 = N'"& pushimg2 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg3 = N'"& pushimg3 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg4 = N'"& pushimg4 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg5 = N'"& pushimg5 &"'" + VbCrlf
		sqlStr = sqlStr + " ,istargetMsg= '"& istargetMsg &"'" + VbCrlf
		sqlStr = sqlStr + " ,admcomment= N'"& html2db(admcomment) &"'" + VbCrlf
		sqlStr = sqlStr + " ,noduppDate= "& CHKIIF(noduppDate<>"1","NULL",noduppDate) &"" + VbCrlf
		sqlStr = sqlStr + " ,noduppDate2= "& CHKIIF(noduppDate2<>"1","NULL",noduppDate2) &"" + VbCrlf
		sqlStr = sqlStr + " ,noduppDate3= "& CHKIIF(noduppDate3<>"1","NULL",noduppDate3) &"" + VbCrlf
		sqlStr = sqlStr + " ,targetKey= "& CHKIIF(targetKey="","NULL",targetKey) &"" + VbCrlf
		sqlStr = sqlStr + " ,baseIdx= "& CHKIIF(baseIdx="","NULL",baseIdx) &"" + VbCrlf
		sqlStr = sqlStr + " ,notclickyn= '"& notclickyn &"'" + VbCrlf

		if (CStr(PreTargetKey)<>CStr(targetKey)) then
		    sqlStr = sqlStr + " ,mayTargetCnt=0" + VbCrlf
		    sqlStr = sqlStr + " ,targetState=0" + VbCrlf
		end if

		sqlStr = sqlStr & " , makeridarr=N'"& makeridarr &"'" & VbCrlf
		sqlStr = sqlStr & " , itemidarr=N'"& itemidarr &"'" & VbCrlf
		sqlStr = sqlStr & " , keywordarr=N'"& keywordarr &"'" & VbCrlf
		sqlStr = sqlStr & " , bonuscouponidxarr=N'"& bonuscouponidxarr &"'" & VbCrlf
		sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
		sqlStr = sqlStr + " , lastupdate=getdate()"  + VbCrlf
		sqlStr = sqlStr + " , sendranking = "& sendranking &"" + VbCrlf
		sqlStr = sqlStr + " , pushcontents = N'"& pushcontents &"'" + VbCrlf
		sqlStr = sqlStr + " , privateYN = '"& privateYN &"' where" + VbCrlf
		sqlStr = sqlStr + " idx = "& idx

		'response.write sqlStr & "<br>"
		'response.end
		dbget.Execute sqlStr

		'' gaparam 추가 2016/11/09 ------------------------------------------------------------------
		if InStr(subtitle,"gaparam")<1 then
			gaparam = ""

			sqlStr = "select top 1 idx, isNULL(targetKey,0) as targetKey from db_contents.dbo.tbl_app_push_reserve where idx="&idx&""

			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				ppidx = rsget("idx")
				gaparam = "gaparam=push_"&ppidx&"_"&rsget("targetKey")
				'utmparam = "utm_source=10x10&utm_medium=push&utm_campaign="& yyyymmdd &"_"&ppidx&"_"&rsget("targetKey")
			end if
			rsget.close

			if (gaparam<>"") then
				if InStr(subtitle,"?")>0 then
					gaparam = "&"&gaparam
					'gaparam = "&"&gaparam&"&"&utmparam
				else
					gaparam = "?"&gaparam
					'gaparam = "?"&gaparam&"&"&utmparam
				end if
				
				sqlStr = " update db_contents.dbo.tbl_app_push_reserve" + VbCrlf
				sqlStr = sqlStr + " set pushurl=N'"&subtitle&gaparam&"'"  + VbCrlf
				sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
				sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
				sqlStr = sqlStr + " idx="&idx
				
				dbget.Execute sqlStr
			end if
		end if

		response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
		dbget.close()	:	response.End
	
	Case "state"
		if idx = "" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('푸시번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		'' 조건 추가 state0=>1 로 변경시(발송예약), 타겟 상태 check
		'' 타겟 발송인경우 istargetMsg=1 & targetstate=0 & targetKey<>9999(수기) & 발송예약일이 6시간 이후인경우 타겟 예약상태로 변경함.
		'' 발송 엔진에서 타겟 상태가 타겟 예약인경우 발송 30분 전에 타게팅함.
		
		sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
		sqlStr = sqlStr + " state = '"& state &"' where" + VbCrlf
		sqlStr = sqlStr + " idx = "& idx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('발송상태가 변경 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	Case "del"
		sqlStr = " update db_contents.dbo.tbl_app_push_reserve set " + VbCrlf
		sqlStr = sqlStr + " isusing = 'N'" + VbCrlf
		sqlStr = sqlStr + " , lastadminid='"& lastadminid &"'"  + VbCrlf
		sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
		sqlStr = sqlStr + " idx = "& idx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('사용여부가 변경 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	'/타게팅
	Case "target"	
		sqlStr = " exec db_contents.[dbo].[sp_Ten_Push_Reserve_MakeTarget] "&idx

		'response.write sqlStr & "<Br>"
		dbget.CommandTimeout = 60*10   ' 10분
		dbget.Execute sqlStr

		response.write "<script type='text/javascript'>alert('타겟 설정 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	'/관리자 리타게팅
	Case "retarget"	
	    sqlStr = "update db_contents.dbo.tbl_app_push_reserve set targetState=0 where idx="&idx
	    dbget.Execute sqlStr
	    
		sqlStr = " exec db_contents.[dbo].[sp_Ten_Push_Reserve_MakeTarget] "&idx
		dbget.CommandTimeout = 60*10   ' 10분
		dbget.Execute sqlStr
		response.write "<script type='text/javascript'>alert('타겟 설정 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	Case "abtarget"	
	    sqlStr = "db_contents.[dbo].[sp_Ten_Push_Reserve_MakeTarget_DIVIDE]"
	    
	    set iADODBcmd = server.CreateObject("ADODB.Command")
	    iADODBcmd.ActiveConnection = dbget
        iADODBcmd.CommandText = sqlStr
        iADODBcmd.CommandType = adCmdStoredProc

	    iADODBcmd.Parameters.Append iADODBcmd.CreateParameter("returnValue", adInteger, adParamReturnValue)
        iADODBcmd.Parameters.Append iADODBcmd.CreateParameter("@targetIDX", adInteger, adParamInput, , idx)
	    
	    iADODBcmd.Execute

        intResult = iADODBcmd.Parameters("returnValue").Value

        set iADODBcmd = Nothing
        dbget.Close

        if (intResult<1) then
            response.write "<script type='text/javascript'>alert('[ERR:"&intResult&"] 타겟 분리설정중 오류.');</script>"
        else
            response.write "<script type='text/javascript'>alert('["&intResult&"]B타겟 설정 되었습니다. 발송일정을 수정하시기바랍니다.');</script>"
        	session.codePage = 949
		    Response.write "<script type='text/javascript'>parent.location.href='/admin/appmanage/push/msg/poppushmsg_edit.asp?idx="&intResult&"';</script>"
			dbget.close()	:	response.End
        end if

		session.codePage = 949
		dbget.close()	:	response.End

    CASE ELSE
        response.write "<script type='text/javascript'>alert('정의되지 않았음 "&mode&"');</script>"
		session.codePage = 949
		dbget.close()	:	response.End
End Select
%>

<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->