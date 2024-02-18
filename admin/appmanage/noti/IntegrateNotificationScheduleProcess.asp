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
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim sIdx, notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents
dim pushurl, templateCode, contents, button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type
dim failed_subject, failed_msg, etc_template_code, member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate
dim adminUserid, isusing, MaxValidLen, len1, len2, len3
dim sqlStr, mode, i, timeCheck, menupos, replacetagcode, currentTime, sendReserveTime
dim addCount, addParamMsg, kk, iparam, iparamvalue, useridarr, time1, time2
dim titletemp, contentstemp, failed_subjecttemp, failed_msgtemp, replacetagcodearray, replacetagcodetemp
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	sIdx = requestcheckvar(getNumeric(trim(request("sIdx"))),10)
    notiType=requestcheckvar(trim(request("notiType")),32)
    linkCode=requestcheckvar(getNumeric(trim(request("linkCode"))),10)
    startDate = requestcheckvar(trim(request("startDate")),32)
    endDate = requestcheckvar(trim(request("endDate")),32)
    time1=requestcheckvar(trim(request("time1")),2)
    time2=requestcheckvar(trim(request("time2")),2)
    pushIsusing=requestcheckvar(trim(request("pushIsusing")),1)
    kakaoAlrimIsusing=requestcheckvar(trim(request("kakaoAlrimIsusing")),1)
    pushtitle=requestcheckvar(trim(request("pushtitle")),800)
    pushcontents=requestcheckvar(trim(request("pushcontents")),4000)
    pushurl=requestcheckvar(trim(request("pushurl")),500)
    templateCode=requestcheckvar(trim(request("template_code")),32)
    contents=requestcheckvar(trim(request("contents")),1000)
    button_name=requestcheckvar(trim(request("button_name")),64)
    button_url_mobile=requestcheckvar(trim(request("button_url_mobile")),256)
    button_name2=requestcheckvar(trim(request("button_name2")),64)
    button_url_mobile2=requestcheckvar(trim(request("button_url_mobile2")),256)
    failed_type=requestcheckvar(trim(request("failed_type")),3)
    failed_subject=requestcheckvar(trim(request("failed_subject")),50)
    failed_msg=requestcheckvar(trim(request("failed_msg")),1000)
    etc_template_code=requestcheckvar(trim(request("etc_template_code")),32)
    member_smsok_checkyn=requestcheckvar(trim(request("member_smsok_checkyn")),1)
    member_kakaoalrimyn_checkyn=requestcheckvar(trim(request("member_kakaoalrimyn_checkyn")),1)   
    isusing=requestcheckvar(trim(request("isusing")),1)
    mode = RequestCheckVar(request("mode"),32)
	useridarr = requestcheckvar(request("useridarr"),256)

adminUserid=session("ssBctId")
timeCheck = false
member_smsok_checkyn="Y"

addCount = Request.Form("params").Count
addParamMsg=""  ''"param1":"value1","param2":"value2"
For kk = 1 To addCount
	iparam = Request.Form("params")(kk)
	iparamvalue = request.Form("paramvalue")(kk)

	if (iparam<>"" and iparamvalue<>"") then
		addParamMsg = addParamMsg & CHR(34)&iparam&CHR(34)&":"&CHR(34)&iparamvalue&CHR(34)
		addParamMsg = addParamMsg & ","
	end if
Next

if (right(addParamMsg,1)=",") then
	addParamMsg = Left(addParamMsg,Len(addParamMsg)-1)
end if

if (mode="mInsert") or (mode="mEdit") then
    if pushIsusing="" or isnull(pushIsusing) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('푸시 사용 여부를 입력해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if kakaoAlrimIsusing="" or isnull(kakaoAlrimIsusing) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('카카오 알림톡 사용 여부를 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if notiType="" or isnull(notiType) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('구분을 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if linkCode="" or isnull(linkCode) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('관련코드를 등록해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if startDate="" or isnull(startDate) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('기간 시작일을 등록해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if endDate="" or isnull(endDate) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('기간 종료일을 등록해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if time1="" or isnull(time1) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('발송시간 시간을 정확하게 입력해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if time2="" or isnull(time2) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('발송시간 분을 정확하게 입력해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

	timeCheck = false
	if time2 = 00 or time2 = 10 or time2 = 20 or time2 = 30 or time2 = 40 or time2 = 50 then
		timeCheck = true
	end if
	if not(timeCheck) then
        if not(C_ADMIN_AUTH) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('발송은 10분 단위로 등록 하실수 있습니다.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
	end if
    reserveTime = time1 & ":" & time2

    if isusing="" or isnull(isusing) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('알림사용여부를 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    if kakaoAlrimIsusing="Y" then
        if contents="" or isnull(contents) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('카카오 알림톡 내용을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
        if instr(contents,"#{") then
            response.write "<script type='text/javascript'>"
            response.write "	alert('카카오 알림톡 템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
		if checkNotValidHTML(contents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('카카오 알림톡 내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if
        if button_name<>"" and not(isnull(button_name)) then
            button_name = replace(button_name,vbcrlf,"")
        end if
        if button_url_mobile<>"" and not(isnull(button_url_mobile)) then
            button_url_mobile = replace(button_url_mobile,vbcrlf,"")
        end if
        if button_name2<>"" and not(isnull(button_name2)) then
            button_name2 = replace(button_name2,vbcrlf,"")
        end if
        if button_url_mobile2<>"" and not(isnull(button_url_mobile2)) then
            button_url_mobile2 = replace(button_url_mobile2,vbcrlf,"")
        end if
        ' 수기템플릿
        if templateCode="etc-9999" then
            if etc_template_code="" or isnull(etc_template_code) then
                response.write "<script type='text/javascript'>"
                response.write "	alert('수기템플릿코드를 입력해 주세요.');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if
            templateCode = etc_template_code
        else
            etc_template_code=""
        end if

		replacetagcode=""
		sqlStr = "SELECT" & vbcrlf
		sqlStr = sqlStr & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
		sqlStr = sqlStr & " and q.targetkey=N'999'" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			replacetagcode 	= trim(db2html(rsget("replacetagcode")))
		End IF			
		rsget.Close	

		contentstemp = contents
		failed_subjecttemp = failed_subject
		failed_msgtemp = failed_msg

		if replacetagcode<>"" and not(isnull(replacetagcode)) then
			replacetagcodearray = split(replacetagcode,",")

			if isarray(replacetagcodearray) then
				for i = 0  to ubound(replacetagcodearray)
					replacetagcodetemp = trim(replacetagcodearray(i))
					if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
						contentstemp = replace(contentstemp,replacetagcodetemp,"")
						failed_subjecttemp = replace(failed_subjecttemp,replacetagcodetemp,"")
						failed_msgtemp = replace(failed_msgtemp,replacetagcodetemp,"")
					end if
				next
			end if
		end if

		if instr(contentstemp,"${")>0 or instr(failed_subjecttemp,"${")>0 or instr(failed_msgtemp,"${")>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

        ' 이상한 데이터가 있을까바 데이터 가공
        contents = replace(contents,"'","")
    end if

    if pushIsusing="Y" then
        if pushtitle="" or isnull(pushtitle) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('푸시 제목을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
        if pushcontents="" or isnull(pushcontents) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('푸시 내용을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
        if pushtitle<>"" then
            pushtitle = replace(pushtitle,vbcrlf,"")

            if checkNotValidHTML(pushtitle) then
                response.write "<script type='text/javascript'>"
                response.write "	alert('푸시 제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if
        end if
        if pushcontents<>"" then
            pushcontents = replace(pushcontents,vbcrlf,"\n")

            if checkNotValidHTML(pushcontents) then
                response.write "<script type='text/javascript'>"
                response.write "	alert('푸시 내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
                response.write "</script>"
                session.codePage = 949
                dbget.close()	:	response.End
            end if
        end if
        if pushurl="" or isnull(pushurl) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('푸시링크을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
        replacetagcode="${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}"

        titletemp = pushtitle
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
            response.write "	alert('푸시 제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if

        ' PUSH메시지 크기
        ''' ios 메세지(json) 의 총길이는 256 바이트를 넘을 수 없음
        '' pushtitle + pushurl 길이를 제한 <= 160 (169 까지는 나갔음..)
        'MaxValidLen = 186 ''160  2016/11/09

        ' iOS 8 이상, Android 4 이상에서 전체 전송 제한 크기는 4Kb로 확인됨
        '   다만 UTF8 한글 특성상 한글 한글자는 3byte가 할당되므로 1300자 정도 넣을 있으며
        '   통신에 기타 해더및 추가 정보를 빼면 약 800자정도 넣을 수 있는것으로 판단됨
        MaxValidLen = 800		' 2018.08.21
        sqlStr = " select len(N'"& pushtitle &"') as titleLen, len(N'"& pushurl &"') as urlLen, len(N'"& pushcontents &"') as pushcontents"

        'response.write sqlStr & "<br>"
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
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
        ' if instr(pushurl,".asp") < 1 then
        '     response.write "<script type='text/javascript'>alert('푸시메세지 링크에 .asp 가 없습니다.');</script>"
        '     session.codePage = 949
        '     response.write "<script type='text/javascript'>history.back();</script>"
        '     dbget.Close() : response.end
        ' end if
        ' if instr(pushurl,"?") > 0 or instr(pushurl,"&") > 0 then
        '     if instr(pushurl,".asp?") < 1 and instr(pushurl,"/?") < 1 then
        '         response.write "<script type='text/javascript'>alert('푸시메세지 링크 형식이 잘못되었습니다[0]');</script>"
        '         session.codePage = 949
        '         response.write "<script type='text/javascript'>history.back();</script>"
        '         dbget.Close() : response.end
        '     end If
        ' end if
        if instr(pushurl,"이벤트번호") > 0 or instr(pushurl,"이벤트코드") > 0 or instr(pushurl,"상품번호") > 0 or instr(pushurl,"상품코드") > 0 then
            response.write "<script type='text/javascript'>alert('푸시메세지 링크에 한글로된 잘못된 형식이 있습니다.');</script>"
            session.codePage = 949
            response.write "<script type='text/javascript'>history.back();</script>"
            dbget.Close() : response.end
        end if
    end if
end if

currentTime = date & " " & Format00(2,hour(now())) & ":" & Format00(2,minute(now())) & ":" & Format00(2,Second(now()))

If mode = "mInsert" then
    sendReserveTime = date & " " & reserveTime & ":00"
    if datediff("d",startDate,date())>=0 and datediff("d",endDate,date())<=0 then
        if datediff("n",currentTime,sendReserveTime) >= 0 and datediff("n",currentTime,sendReserveTime) < 32 then
            response.write "<script type='text/javascript'>alert('발송시간 30분전에 대상자가 푸시나 알림톡 매뉴에 생성 됩니다.\n발송시간을 다시 지정해 주세요.');</script>"
            session.codePage = 949
            response.write "<script type='text/javascript'>history.back();</script>"
            dbget.Close() : response.end
        end if
    end if

    sqlStr = "insert into db_contents.dbo.tbl_IntegrateNotificationSchedule (" & vbcrlf
    sqlStr = sqlStr & " notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents" & vbcrlf
    sqlStr = sqlStr & " , pushurl, templateCode, contents, button_name, button_url_mobile, button_name2" & vbcrlf
    sqlStr = sqlStr & " , button_url_mobile2, failed_type, failed_subject, failed_msg, etc_template_code" & vbcrlf
    sqlStr = sqlStr & " , member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate, lastUpdate, adminUserid" & vbcrlf
    sqlStr = sqlStr & " , lastUserid, isusing) values (" & vbcrlf
    sqlStr = sqlStr & " N'"& notiType &"',"& linkCode &",N'"& startDate &" 00:00:00',N'"& endDate &" 23:59:59',N'"& reserveTime &"'" & vbcrlf
    sqlStr = sqlStr & " ,N'"& pushIsusing &"',N'"& kakaoAlrimIsusing &"',N'"& pushtitle &"',N'"& pushcontents &"',N'"& pushurl &"'" & vbcrlf
    sqlStr = sqlStr & " ,N'"& templateCode &"',N'"& contents &"',N'"& button_name &"',N'"& button_url_mobile &"',N'"& button_name2 &"'" & vbcrlf
    sqlStr = sqlStr & " ,N'"& button_url_mobile2 &"',N'"& failed_type &"',N'"& failed_subject &"',N'"& failed_msg &"',N'"& etc_template_code &"'" & vbcrlf
    sqlStr = sqlStr & " ,N'"& member_smsok_checkyn &"',N'"& member_kakaoalrimyn_checkyn &"',getdate(),getdate(),N'"& adminUserid &"'" & vbcrlf
    sqlStr = sqlStr & " ,N'"& adminUserid &"',N'"& isusing &"'" & vbcrlf
    sqlStr = sqlStr & " )" & vbcrlf

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('저장 되었습니다.');</script>"
    session.codePage = 949
    Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End

elseIf mode = "mEdit" then
    if sidx="" or isnull(sidx) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('수정을 위한 구분자가 없습니다.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    sqlStr = " select reserveTime from db_contents.dbo.tbl_IntegrateNotificationSchedule where sidx="& sidx &""

    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If not rsget.EOF Then
        sendReserveTime = date & " " & rsget("reserveTime") & ":00"
    end if
    rsget.close

    ' 기존에 입력되어 있던 발송시간 중복 발송 체크
    if datediff("d",startDate,date())>=0 and datediff("d",endDate,date())<=0 then
        if datediff("n",currentTime,sendReserveTime) >= 0 and datediff("n",currentTime,sendReserveTime) < 32 then
            response.write "<script type='text/javascript'>alert('기존에 입력되어 있던 발송시간 기준으로 발송대상자가 이미 푸시나 알림톡 발송 매뉴에 작성 되었습니다.');</script>"
            session.codePage = 949
            response.write "<script type='text/javascript'>history.back();</script>"
            dbget.Close() : response.end
        end if
    end if

    ' 수정요청된 발송시간 체크
    sendReserveTime = date & " " & reserveTime & ":00"
    if datediff("d",startDate,date())>=0 and datediff("d",endDate,date())<=0 then
        if datediff("n",currentTime,sendReserveTime) >= 0 and datediff("n",currentTime,sendReserveTime) < 32 then
            response.write "<script type='text/javascript'>alert('발송시간 30분전에 대상자가 푸시나 알림톡 매뉴에 생성 됩니다.\n발송시간을 다시 지정해 주세요.');</script>"
            session.codePage = 949
            response.write "<script type='text/javascript'>history.back();</script>"
            dbget.Close() : response.end
        end if
    end if

    sqlStr="update db_contents.dbo.tbl_IntegrateNotificationSchedule" & vbcrlf
    sqlStr = sqlStr & " set notiType='"& notiType &"'" & vbcrlf
    sqlStr = sqlStr & " , linkCode="& linkCode &"" & vbcrlf
    sqlStr = sqlStr & " , startDate=N'"& startDate &" 00:00:00'" & vbcrlf
    sqlStr = sqlStr & " , endDate=N'"& endDate &" 23:59:59'" & vbcrlf
    sqlStr = sqlStr & " , reserveTime=N'"& reserveTime &"'" & vbcrlf
    sqlStr = sqlStr & " , pushIsusing=N'"& pushIsusing &"'" & vbcrlf
    sqlStr = sqlStr & " , kakaoAlrimIsusing=N'"& kakaoAlrimIsusing &"'" & vbcrlf
    sqlStr = sqlStr & " , pushtitle=N'"& pushtitle &"'" & vbcrlf
    sqlStr = sqlStr & " , pushcontents=N'"& pushcontents &"'" & vbcrlf
    sqlStr = sqlStr & " , pushurl=N'"& pushurl &"'" & vbcrlf
    sqlStr = sqlStr & " , templateCode=N'"& templateCode &"'" & vbcrlf
    sqlStr = sqlStr & " , contents=N'"& contents &"'" & vbcrlf
    sqlStr = sqlStr & " , button_name=N'"& button_name &"'" & vbcrlf
    sqlStr = sqlStr & " , button_url_mobile=N'"& button_url_mobile &"'" & vbcrlf
    sqlStr = sqlStr & " , button_name2=N'"& button_name2 &"'" & vbcrlf
    sqlStr = sqlStr & " , button_url_mobile2=N'"& button_url_mobile2 &"'" & vbcrlf
    sqlStr = sqlStr & " , failed_type=N'"& failed_type &"'" & vbcrlf
    sqlStr = sqlStr & " , failed_subject=N'"& failed_subject &"'" & vbcrlf
    sqlStr = sqlStr & " , failed_msg=N'"& failed_msg &"'" & vbcrlf
    sqlStr = sqlStr & " , etc_template_code=N'"& etc_template_code &"'" & vbcrlf
    sqlStr = sqlStr & " , member_smsok_checkyn=N'"& member_smsok_checkyn &"'" & vbcrlf
    sqlStr = sqlStr & " , member_kakaoalrimyn_checkyn=N'"& member_kakaoalrimyn_checkyn &"'" & vbcrlf
    sqlStr = sqlStr & " , lastUpdate=getdate()" & vbcrlf
    sqlStr = sqlStr & " , lastUserid=N'"& adminUserid &"'" & vbcrlf
    sqlStr = sqlStr & " , isusing=N'"& isusing &"' where" & vbcrlf
    sqlStr = sqlStr & " sidx="& sidx &""

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('수정 되었습니다.');</script>"
    session.codePage = 949
    Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End

elseIf mode = "pushTestSend" then
    response.write "userid:"& useridarr &"<br>"
    response.write "pushtitle:"&pushtitle&"<br>"
    response.write "pushcontents:"&pushcontents&"<br>"
    response.write "addparams:"&addParamMsg&"<br>"

    if (useridarr="") then
        response.write "필수 값 체크 오류"
        session.codePage = 949
        dbget.close():response.end
    end if

    pushtitle = trim(pushtitle)
    pushcontents = trim(pushcontents)

    if pushtitle="" or isnull(pushtitle) then
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
    if pushtitle<>"" then
        pushtitle = replace(pushtitle,vbcrlf,"")

        if checkNotValidHTML(pushtitle) then
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

    'IF LEFT(message,1)="{" and RIGHT(message,1)="}" then
    '    isMsgArr = 1
    'ELSE
    '    isMsgArr = 0
    'end if

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
    sqlStr = sqlStr & " 	    replace(" & vbcrlf
    'sqlStr = sqlStr & " 		    replace(msg1 + msg2 + msg3,'${CUSTOMERID}',(CASE when isnull(a.userid,'')='' then '고객' WHEN LEN(isnull(a.userid,''))>1 THEN LEFT(isnull(a.userid,''),LEN(isnull(a.userid,''))-1)+N'*' ELSE isnull(a.userid,'') END))" & vbcrlf
    sqlStr = sqlStr & " 		    replace(msg1 + msg2 + msg3,'${CUSTOMERID}',(CASE when isnull(a.userid,'')='' then '고객' ELSE isnull(a.userid,'') END))" & vbcrlf
    'sqlStr = sqlStr & " 	    ,'${CUSTOMERNAME}',(CASE when isnull(a.username,'')='' then '고객' WHEN LEN(isnull(a.username,''))>1 THEN LEFT(isnull(a.username,''),LEN(isnull(a.username,''))-1)+N'*' ELSE isnull(a.username,'') END))" & vbcrlf
    sqlStr = sqlStr & " 	    ,'${CUSTOMERNAME}',(CASE when isnull(a.username,'')='' then '고객' ELSE isnull(a.username,'') END))" & vbcrlf
	sqlStr = sqlStr & " 	,'${CUSTOMERLEVELNAME}',isnull(a.userlevelname,'비회원'))" & vbcrlf
    sqlStr = sqlStr & " ) AS sendMsg" & vbcrlf
    sqlStr = sqlStr & " , a.userid, NULL, 'N', NULL, 3, N'Y'" & vbcrlf		' , a.regIdx 
    sqlStr = sqlStr & " FROM ( " & vbcrlf
    sqlStr = sqlStr & " 	SELECT " & vbcrlf
    sqlStr = sqlStr & " 	r.appKey " & vbcrlf
    sqlStr = sqlStr & " 	, 0 AS multiPsKey " & vbcrlf
    sqlStr = sqlStr & " 	, 0 AS sendState, r.deviceid " & vbcrlf

    'if isMsgArr=1 then
    '    sqlStr = sqlStr & " 	, N'{""title"":"& pushtitle &",""noti"":"& pushcontents &"' as msg1 " & vbcrlf
    'else
        sqlStr = sqlStr & " 	, N'{""title"":"""& replace(pushtitle,"""","\""") &""",""noti"":"""& replace(pushcontents,"""","\""") &"""' as msg1 " & vbcrlf
    'end if

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
    sqlStr = " update db_contents.dbo.tbl_IntegrateNotificationSchedule" & VbCrlf
    sqlStr = sqlStr & " set pushTestCount = pushTestCount + "& ubound(split(useridarr,","))+1 &" where" & VbCrlf
    sqlStr = sqlStr & " sIdx = "& sIdx &""
                
    'response.write sqlStr
    dbget.Execute sqlStr

    'Response.write "sendedMsg:"&sendedMsg
    Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='닫기'/>"
    Response.end

elseIf mode = "kakaoAlrimTestSend" then
    if (useridarr="" or isnull(useridarr)) then
        response.write "필수 값 체크 오류. 발송 아이디가 없습니다."
        session.codePage = 949
        dbget.close():response.end
    end if
    if (sIdx="" or isnull(sIdx)) then
        response.write "필수 값 체크 오류. 발송키가 없습니다."
        session.codePage = 949
        dbget.close():response.end
    end if

    if contents="" or isnull(contents) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('카카오 알림톡 내용을 입력해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if instr(contents,"#{") then
        response.write "<script type='text/javascript'>"
        response.write "	alert('템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if checkNotValidHTML(contents) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    if isnull(button_name) then button_name=""
    if isnull(button_url_mobile) then button_url_mobile=""
    if isnull(button_name2) then button_name2=""
    if isnull(button_url_mobile2) then button_url_mobile2=""

    if button_name<>"" and not(isnull(button_name)) then
        button_name = replace(button_name,vbcrlf,"")
    end if
    if button_url_mobile<>"" and not(isnull(button_url_mobile)) then
        button_url_mobile = replace(button_url_mobile,vbcrlf,"")
    end if
    if button_name2<>"" and not(isnull(button_name2)) then
        button_name2 = replace(button_name2,vbcrlf,"")
    end if
    if button_url_mobile2<>"" and not(isnull(button_url_mobile2)) then
        button_url_mobile2 = replace(button_url_mobile2,vbcrlf,"")
    end if
    ' 수기템플릿
    if templateCode="etc-9999" then
        if etc_template_code="" or isnull(etc_template_code) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('수기템플릿코드를 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
        templateCode = etc_template_code
    else
        etc_template_code=""
    end if

    replacetagcode=""
    sqlStr = "SELECT" & vbcrlf
    sqlStr = sqlStr & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
    sqlStr = sqlStr & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
    sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
    sqlStr = sqlStr & " and q.targetkey=N'999'" & vbcrlf

    'response.write sqlStr &"<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF not rsget.EOF THEN
        replacetagcode 	= trim(db2html(rsget("replacetagcode")))
    End IF			
    rsget.Close	

    contentstemp = contents
    failed_subjecttemp = failed_subject
    failed_msgtemp = failed_msg

    if replacetagcode<>"" and not(isnull(replacetagcode)) then
        replacetagcodearray = split(replacetagcode,",")

        if isarray(replacetagcodearray) then
            for i = 0  to ubound(replacetagcodearray)
                replacetagcodetemp = trim(replacetagcodearray(i))
                if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
                    contentstemp = replace(contentstemp,replacetagcodetemp,"")
                    failed_subjecttemp = replace(failed_subjecttemp,replacetagcodetemp,"")
                    failed_msgtemp = replace(failed_msgtemp,replacetagcodetemp,"")
                end if
            next
        end if
    end if

    if instr(contentstemp,"${")>0 or instr(failed_subjecttemp,"${")>0 or instr(failed_msgtemp,"${")>0 then
        response.write "<script type='text/javascript'>"
        response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    ' 이상한 데이터가 있을까바 데이터 가공
    contents = replace(contents,"'","")

    useridarr = replace(useridarr,"'","")     
    useridarr = replace(useridarr,"""","")
    useridarr = replace(useridarr,".",",")
    useridarr = replace(useridarr,",,",",")
    useridarr = "'" & replace(useridarr,",","','") & "'"

    response.write "카카오톡 알림톡 테스트 발송" & "<br>"

    sqlStr = "SELECT distinct" & vbcrlf
    sqlStr = sqlStr & " getdate() as REQDATE,N'1' as STATUS," & vbcrlf
    sqlStr = sqlStr & " replace(isnull(u.usercell,''),'-','') as PHONE," & vbcrlf		' 수신자 휴대폰 번호
    sqlStr = sqlStr & " N'1644-6030' as CALLBACK" & vbcrlf	' 발신자 번호
    ' 알림톡 내용
    sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
    sqlStr = sqlStr & "     replace(" & vbcrlf
    sqlStr = sqlStr & "         replace(" & vbcrlf
    sqlStr = sqlStr & "             replace(" & vbcrlf
    sqlStr = sqlStr & "                 replace(" & vbcrlf
    'sqlStr = sqlStr & " 	    			replace(N'"& contents &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
    sqlStr = sqlStr & " 	    			replace(N'"& contents &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
    sqlStr = sqlStr & "                 ,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
    sqlStr = sqlStr & "             ,'${CUSTOMERLEVELNAME}'" & vbcrlf
    sqlStr = sqlStr & "                 , isnull((case when L.userlevel = 0 then 'WHITE'"
    sqlStr = sqlStr & " 	            when L.userlevel = 1 then 'RED'"
    sqlStr = sqlStr & " 	            when L.userlevel = 2 then 'VIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 3 then 'VIP GOLD'"
    sqlStr = sqlStr & " 	            when L.userlevel = 4 then 'VVIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 7 then 'STAFF'"
    sqlStr = sqlStr & " 	            when L.userlevel = 8 then 'FAMILY'"
    sqlStr = sqlStr & " 	            when L.userlevel = 9 then 'BIZ'"
    sqlStr = sqlStr & "                 end),'비회원')"
    sqlStr = sqlStr & "             )" & vbcrlf
    sqlStr = sqlStr & "         ,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
    sqlStr = sqlStr & "     ,'${MILEAGE}','TEST마일리지')" & vbcrlf
    sqlStr = sqlStr & " ) as MSG" & vbcrlf
    sqlStr = sqlStr & " ,N'"& templateCode &"' as TEMPLATE_CODE" & vbcrlf	' 알림톡 템플릿 번호
    sqlStr = sqlStr & " ,N'"& failed_type &"' as FAILED_TYPE" & vbcrlf		' 알림톡 실패시 문자 형식 > SMS / LMS
    sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
    sqlStr = sqlStr & "     replace(" & vbcrlf
    sqlStr = sqlStr & "         replace(" & vbcrlf
    sqlStr = sqlStr & "             replace(" & vbcrlf
    sqlStr = sqlStr & "                 replace(" & vbcrlf
    'sqlStr = sqlStr & " 			    	replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
    sqlStr = sqlStr & " 			    	replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
    sqlStr = sqlStr & "                 ,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
    sqlStr = sqlStr & "             ,'${CUSTOMERLEVELNAME}'" & vbcrlf
    sqlStr = sqlStr & "                 , isnull((case when L.userlevel = 0 then 'WHITE'"
    sqlStr = sqlStr & " 	            when L.userlevel = 1 then 'RED'"
    sqlStr = sqlStr & " 	            when L.userlevel = 2 then 'VIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 3 then 'VIP GOLD'"
    sqlStr = sqlStr & " 	            when L.userlevel = 4 then 'VVIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 7 then 'STAFF'"
    sqlStr = sqlStr & " 	            when L.userlevel = 8 then 'FAMILY'"
    sqlStr = sqlStr & " 	            when L.userlevel = 9 then 'BIZ'"
    sqlStr = sqlStr & "                 end),'비회원')"
    sqlStr = sqlStr & "             )" & vbcrlf
    sqlStr = sqlStr & "         ,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
    sqlStr = sqlStr & "     ,'${MILEAGE}','TEST마일리지')" & vbcrlf
    sqlStr = sqlStr & " ) as FAILED_SUBJECT" & vbcrlf      ' 실패시 문자 제목 (LMS 전송시에만 필요)
    ' 실패시 문자 내용
    sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
    sqlStr = sqlStr & "     replace(" & vbcrlf
    sqlStr = sqlStr & "         replace(" & vbcrlf
    sqlStr = sqlStr & "             replace(" & vbcrlf
    sqlStr = sqlStr & "                 replace(" & vbcrlf
    'sqlStr = sqlStr & " 		    		replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
    sqlStr = sqlStr & " 		    		replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
    sqlStr = sqlStr & "                 ,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
    sqlStr = sqlStr & "             ,'${CUSTOMERLEVELNAME}'" & vbcrlf
    sqlStr = sqlStr & "                 , isnull((case when L.userlevel = 0 then 'WHITE'"
    sqlStr = sqlStr & " 	            when L.userlevel = 1 then 'RED'"
    sqlStr = sqlStr & " 	            when L.userlevel = 2 then 'VIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 3 then 'VIP GOLD'"
    sqlStr = sqlStr & " 	            when L.userlevel = 4 then 'VVIP'"
    sqlStr = sqlStr & " 	            when L.userlevel = 7 then 'STAFF'"
    sqlStr = sqlStr & " 	            when L.userlevel = 8 then 'FAMILY'"
    sqlStr = sqlStr & " 	            when L.userlevel = 9 then 'BIZ'"
    sqlStr = sqlStr & "                 end),'비회원')"
    sqlStr = sqlStr & "             )" & vbcrlf
    sqlStr = sqlStr & "         ,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
    sqlStr = sqlStr & "     ,'${MILEAGE}','TEST마일리지')" & vbcrlf
    sqlStr = sqlStr & " ) as FAILED_MSG" & vbcrlf
    ' 버튼 구성 내용 (버튼타입에만 필요 / v4 메뉴얼 참고)
    if button_name<>"" and button_url_mobile<>"" and button_name2<>"" and button_url_mobile2<>"" then
        sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""},{""name"":"""& button_name2 &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile2 &"""}]}' as BUTTON_JSON" & vbcrlf
    elseif button_name<>"" and button_url_mobile<>"" then
        sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""}]}' as BUTTON_JSON" & vbcrlf
    else
        sqlStr = sqlStr & " ,N'' as BUTTON_JSON" & vbcrlf
    end if
    sqlStr = sqlStr & " , u.userid, "& sIdx &" as sIdx" & vbcrlf
    sqlStr = sqlStr & " into #tmpuser" & vbcrlf
    sqlStr = sqlStr & " from db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
    sqlStr = sqlStr & " left join db_user.dbo.tbl_logindata l with (nolock)"
    sqlStr = sqlStr & " 	on u.userid=l.userid"
    sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

    'response.write sqlStr & "<br>"
    'response.end
    dbget.Execute sqlStr

    ' 발송 DB에 저장
    sqlStr = "INSERT INTO LOGISTICSDB.[db_kakaomsg_v4_mkt].dbo.KKO_MSG (REQDATE, STATUS, PHONE, CALLBACK, MSG, TEMPLATE_CODE, FAILED_TYPE, FAILED_SUBJECT, FAILED_MSG, BUTTON_JSON, ETC1, ETC2)" & vbcrlf
    sqlStr = sqlStr & " 	select * from #tmpuser" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

    response.write "==================================<br>"
    response.write "테스트 메시지 발송요청 되었습니다.<br>"

    '//테스트카운트 올림
    sqlStr = " update db_contents.dbo.tbl_IntegrateNotificationSchedule" & VbCrlf
    sqlStr = sqlStr & " set kakaoAlrimTestCount = kakaoAlrimTestCount + "& ubound(split(useridarr,","))+1 &" where" & VbCrlf
    sqlStr = sqlStr & " sIdx = "& sIdx &""
                
    'response.write sqlStr
    dbget.Execute sqlStr

    'Response.write "sendedMsg:"&sendedMsg
    Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='닫기'/>"
    session.codePage = 949
    dbget.close()	:	response.End
else
    response.write "<script type='text/javascript'>alert('정의되지 않았음 "&mode&"');</script>"
    session.codePage = 949
    dbget.close()	:	response.End
end if
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
