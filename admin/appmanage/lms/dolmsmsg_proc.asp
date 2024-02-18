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
' Description : LMS발송관리
' Hieditor : 2020.03.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim i, sqlStr, menupos, reservationdate, reservetime, reservemin, yyyymmdd, titleMaxValidLen, len1, len2
dim ridx,sendmethod,title,contents,state,testsend,isusing,reservedate,exception7dayyn,targetkey,targetstate
dim targetcnt,regadminid,lastadminid,regdate,lastupdate,repeatlmsyn,member_smsok_checkyn,member_pushyn_checkyn, useridarr
dim appkey, message, deviceid, addCount, iparam, iparamvalue, kk, Pretargetkey, mode, olms, clsCode, targetQuery
dim makeridarr, itemidarr, keywordarr, bonuscouponidxarr, button_name, button_url_mobile, button_name2, button_url_mobile2
dim failed_subject, failed_msg, orderitemidexceptionarr, replacetagcode, replacetagcodearray, replacetagcodetemp
dim titletemp, contentstemp, failed_subjecttemp, failed_msgtemp, exceptionlogin, exceptionuserlevelarr, eventcodearr
dim member_kakaoalrimyn_checkyn, etc_template_code, failed_type, template_code
    menupos = requestcheckvar(getNumeric(request("menupos")),10)
    ridx = requestcheckvar(getNumeric(request("ridx")),10)
    sendmethod = requestcheckvar(request("sendmethod"),16)
    title = requestcheckvar(request("title"),120)
    contents = requestcheckvar(request("contents"),1000)
    state = requestcheckvar(getNumeric(request("state")),10)
    testsend = requestcheckvar(getNumeric(request("testsend")),10)
    isusing = requestcheckvar(request("isusing"),1)
    reservationdate = RequestCheckVar(request("reservationdate"),10)
	reservetime		= RequestCheckVar(request("time1"),2)
	reservemin		= RequestCheckVar(request("time2"),2)
    exception7dayyn		= RequestCheckVar(request("exception7dayyn"),16)
    targetkey = requestcheckvar(getNumeric(request("targetkey")),10)
    targetstate = requestcheckvar(getNumeric(request("targetstate")),10)
    targetcnt = requestcheckvar(getNumeric(request("targetcnt")),10)
    repeatlmsyn		= RequestCheckVar(request("repeatlmsyn"),1)
    member_smsok_checkyn		= RequestCheckVar(request("member_smsok_checkyn"),2)
	member_pushyn_checkyn		= RequestCheckVar(request("member_pushyn_checkyn"),1)
	appkey = requestcheckvar(request("appkey"),10)
	message = requestcheckvar(request("message"),800)
	deviceid = LEFT(replace(request("deviceid"),"'",""),200)  '' -- 이 치환 되면 안됨..
    mode = RequestCheckVar(request("mode"),32)
	useridarr = requestcheckvar(request("useridarr"),800)
	makeridarr = requestcheckvar(request("makeridarr"),512)
	itemidarr = requestcheckvar(request("itemidarr"),512)
	keywordarr = requestcheckvar(request("keywordarr"),512)
	bonuscouponidxarr = requestcheckvar(request("bonuscouponidxarr"),512)
	orderitemidexceptionarr = requestcheckvar(request("orderitemidexceptionarr"),512)
	button_name = requestcheckvar(request("button_name"),64)
	button_url_mobile = requestcheckvar(request("button_url_mobile"),256)
	button_name2 = requestcheckvar(request("button_name2"),64)
	button_url_mobile2 = requestcheckvar(request("button_url_mobile2"),256)
	failed_type = requestcheckvar(request("failed_type"),3)
	failed_subject = requestcheckvar(request("failed_subject"),50)
	failed_msg = requestcheckvar(request("failed_msg"),1000)
	template_code = requestcheckvar(request("template_code"),32)
	etc_template_code = requestcheckvar(request("etc_template_code"),32)
	exceptionlogin = requestcheckvar(request("exceptionlogin"),16)
	exceptionuserlevelarr = requestcheckvar(request("exceptionuserlevelarr"),10)
	eventcodearr = requestcheckvar(request("eventcodearr"),512)
	member_kakaoalrimyn_checkyn = requestcheckvar(request("member_kakaoalrimyn_checkyn"),1)

addCount = Request.Form("params").Count
lastadminid = session("ssBctId")

reservedate		= reservationdate &" "& reservetime &":"& reservemin &":000" '예약일
'yyyymmdd = year(date()) & Format00(2,Month(date())) & Format00(2,day(date()))
yyyymmdd = replace(reservationdate,"-","")

if sendmethod="KAKAOALRIM" then
	member_smsok_checkyn=""
else
	member_smsok_checkyn="Y"
	member_kakaoalrimyn_checkyn=""
end if
if repeatlmsyn="" then repeatlmsyn="N"
if isusing="" then isusing="N"

titleMaxValidLen = 120    ' 제목 글자수
if (mode="mInsert") or (mode="mEdit") then
    sqlStr = " select len(N'"&title&"') as titleLen, len(N'"& contents &"') as contentsLen"

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		len1 = rsget("titleLen")
		len2 = rsget("contentsLen")
	end if
	rsget.close
	
	if (len1)>titleMaxValidLen then
		response.write "<script type='text/javascript'>alert('타이틀 길이가 "& titleMaxValidLen &"글자를 초과 할 수 없습니다.\n현재 글자수("& len1 &")');</script>"
		session.codePage = 949
		response.write "<script type='text/javascript'>history.back();</script>"
	    dbget.Close() : response.end
	end if
end if

Select Case mode
	Case "mInsert"
        title = trim(title)
        contents = trim(contents)
		if sendmethod="LMS" then
			if title="" or isnull(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목을 입력해 주세요.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			title = replace(title,vbcrlf,"")

			if checkNotValidHTML(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			if len(title)>120 then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목이 제한길이를 초과하였습니다. 120자 까지 작성 가능합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
        if contents="" or isnull(contents) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('내용을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
		'contents = replace(contents,vbcrlf,"\n")
		if sendmethod="KAKAOALRIM" then
			if instr(contents,"#{") then
				response.write "<script type='text/javascript'>"
				response.write "	alert('템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		if checkNotValidHTML(contents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

        button_name = trim(button_name)
        button_url_mobile = trim(button_url_mobile)
        button_name2 = trim(button_name2)
        button_url_mobile2 = trim(button_url_mobile2)
		failed_type = trim(failed_type)
		failed_subject = trim(failed_subject)
		failed_msg = trim(failed_msg)
		if sendmethod="KAKAOFRIEND" or sendmethod="KAKAOALRIM" then
			if button_name="" or isnull(button_name) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name = replace(button_name,vbcrlf,"")
			end if
			if button_url_mobile="" or isnull(button_url_mobile) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile = replace(button_url_mobile,vbcrlf,"")
			end if
			if button_name2="" or isnull(button_name2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name2 = replace(button_name2,vbcrlf,"")
			end if
			if button_url_mobile2="" or isnull(button_url_mobile2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile2 = replace(button_url_mobile2,vbcrlf,"")
			end if
			if sendmethod="KAKAOALRIM" then
				' 수기템플릿
				if template_code="etc-9999" then
					if etc_template_code="" or isnull(etc_template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('수기템플릿코드를 입력해 주세요.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					template_code = etc_template_code
				else
					etc_template_code=""
				end if
			end if
			if failed_type="LMS" then
				if failed_subject="" or isnull(failed_subject) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목를 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if len(failed_subject)>50 then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				failed_subject = replace(failed_subject,vbcrlf,"")
				if failed_msg="" or isnull(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용을 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if checkNotValidHTML(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if sendmethod="KAKAOALRIM" then
					if instr(failed_subject,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자제목에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if instr(failed_msg,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if template_code="" or isnull(template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('카카오톡 알림톡 템플릿코드가 지정되어 있지 않습니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
				end if
			end if
		end if

		replacetagcode=""
		sqlStr = "SELECT" & vbcrlf
		sqlStr = sqlStr & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
		sqlStr = sqlStr & " and q.targetkey=N'"& targetkey &"'" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			replacetagcode 	= trim(db2html(rsget("replacetagcode")))
		End IF			
		rsget.Close	

		titletemp = title
		contentstemp = contents
		failed_subjecttemp = failed_subject
		failed_msgtemp = failed_msg

		if replacetagcode<>"" and not(isnull(replacetagcode)) then
			replacetagcodearray = split(replacetagcode,",")

			if isarray(replacetagcodearray) then
				for i = 0  to ubound(replacetagcodearray)
					replacetagcodetemp = trim(replacetagcodearray(i))
					if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
						titletemp = replace(titletemp,replacetagcodetemp,"")
						contentstemp = replace(contentstemp,replacetagcodetemp,"")
						failed_subjecttemp = replace(failed_subjecttemp,replacetagcodetemp,"")
						failed_msgtemp = replace(failed_msgtemp,replacetagcodetemp,"")
					end if
				next
			end if
		end if

		if instr(titletemp,"${")>0 or instr(contentstemp,"${")>0 or instr(failed_subjecttemp,"${")>0 or instr(failed_msgtemp,"${")>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		' 이상한 데이터가 있을까바 데이터 가공
		contents = replace(contents,"'","")

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
		if orderitemidexceptionarr <> "" then
			if checkNotValidHTML(orderitemidexceptionarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('해당상품구매한사람제외에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if
		if eventcodearr <> "" then
			if checkNotValidHTML(eventcodearr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('이벤트번호에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
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
		if trim(orderitemidexceptionarr)<>"" then
			orderitemidexceptionarr = replace(replace(replace(trim(orderitemidexceptionarr),".",","),"""",""),"'","")		'오타 교정
			if left(orderitemidexceptionarr,1)="," then orderitemidexceptionarr = mid(orderitemidexceptionarr,2,len(orderitemidexceptionarr))		'오타 교정
			if right(orderitemidexceptionarr,1)="," then orderitemidexceptionarr = left(orderitemidexceptionarr,len(orderitemidexceptionarr)-1)		'오타 교정
		end if
		if trim(eventcodearr)<>"" then
			eventcodearr = replace(replace(replace(trim(eventcodearr),".",","),"""",""),"'","")		'오타 교정
			if left(eventcodearr,1)="," then eventcodearr = mid(eventcodearr,2,len(eventcodearr))		'오타 교정
			if right(eventcodearr,1)="," then eventcodearr = left(eventcodearr,len(eventcodearr)-1)		'오타 교정
		end if

        sqlStr = " insert into db_contents.dbo.tbl_lms_reserve (" & VbCrlf
        sqlStr = sqlStr & " sendmethod,title,contents,state,testsend,reservedate" & VbCrlf
        sqlStr = sqlStr & " ,exception7dayyn,targetkey,targetstate,targetcnt,regadminid,lastadminid,regdate,lastupdate" & VbCrlf
        sqlStr = sqlStr & " ,repeatlmsyn,member_smsok_checkyn,member_pushyn_checkyn, makeridarr, itemidarr, keywordarr, bonuscouponidxarr" & VbCrlf
		sqlStr = sqlStr & " , button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type, failed_subject" & VbCrlf
		sqlStr = sqlStr & " , failed_msg, orderitemidexceptionarr,template_code,exceptionlogin" & VbCrlf
		sqlStr = sqlStr & " , exceptionuserlevelarr, eventcodearr, member_kakaoalrimyn_checkyn, etc_template_code" & VbCrlf
        sqlStr = sqlStr & " ) values (" & VbCrlf
        sqlStr = sqlStr & " N'"& sendmethod &"' ,N'"& html2db(title) &"',N'"& html2db(contents) &"',"& state &",0,N'"& reservedate &"'" & VbCrlf
        sqlStr = sqlStr & " ,N'"& exception7dayyn &"',"& targetkey &",0,0, N'"& lastadminid &"', N'"& lastadminid &"',getdate(),getdate()" & VbCrlf
        sqlStr = sqlStr & " ,N'"& repeatlmsyn &"', N'"& member_smsok_checkyn &"', N'"& member_pushyn_checkyn &"'" & VbCrlf
		sqlStr = sqlStr & " , N'"& makeridarr &"', N'"& itemidarr &"', N'"& keywordarr &"', N'"& bonuscouponidxarr &"'" & VbCrlf
		sqlStr = sqlStr & " , N'"& html2db(button_name) &"', N'"& button_url_mobile &"', N'"& html2db(button_name2) &"', N'"& button_url_mobile2 &"'" & VbCrlf
		sqlStr = sqlStr & " , N'"& failed_type &"', N'"& html2db(failed_subject) &"', N'"& html2db(failed_msg) &"', N'"& orderitemidexceptionarr &"'" & VbCrlf
		sqlStr = sqlStr & " , N'"& template_code &"', N'"& exceptionlogin &"', N'"& exceptionuserlevelarr &"', N'"& eventcodearr &"'" & VbCrlf
        sqlStr = sqlStr & " , N'"& member_kakaoalrimyn_checkyn &"', N'"& etc_template_code &"'" + VbCrlf
		sqlStr = sqlStr & " )" + VbCrlf

        'response.write sqlStr & "<Br>"
		'response.end
        dbget.Execute sqlStr

        response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
        session.codePage = 949
        Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
        dbget.close()	:	response.End

	Case "mEdit"
        title = trim(title)
        contents = trim(contents)
		if sendmethod="LMS" then
			if title="" or isnull(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목을 입력해 주세요.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			title = replace(title,vbcrlf,"")

			if checkNotValidHTML(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			if len(title)>120 then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목이 제한길이를 초과하였습니다. 120자 까지 작성 가능합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
        if contents="" or isnull(contents) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('내용을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
		'contents = replace(contents,vbcrlf,"\n")
		if sendmethod="KAKAOALRIM" then
			if instr(contents,"#{") then
				response.write "<script type='text/javascript'>"
				response.write "	alert('템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
		if checkNotValidHTML(contents) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

        button_name = trim(button_name)
        button_url_mobile = trim(button_url_mobile)
        button_name2 = trim(button_name2)
        button_url_mobile2 = trim(button_url_mobile2)
		failed_type = trim(failed_type)
		failed_subject = trim(failed_subject)
		failed_msg = trim(failed_msg)
		if sendmethod="KAKAOFRIEND" or sendmethod="KAKAOALRIM" then
			if button_name="" or isnull(button_name) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name = replace(button_name,vbcrlf,"")
			end if
			if button_url_mobile="" or isnull(button_url_mobile) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile = replace(button_url_mobile,vbcrlf,"")
			end if
			if button_name2="" or isnull(button_name2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name2 = replace(button_name2,vbcrlf,"")
			end if
			if button_url_mobile2="" or isnull(button_url_mobile2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile2 = replace(button_url_mobile2,vbcrlf,"")
			end if
			if sendmethod="KAKAOALRIM" then
				' 수기템플릿
				if template_code="etc-9999" then
					if etc_template_code="" or isnull(etc_template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('수기템플릿코드를 입력해 주세요.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					template_code = etc_template_code
				else
					etc_template_code=""
				end if
			end if
			if failed_type="LMS" then
				if failed_subject="" or isnull(failed_subject) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목를 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if len(failed_subject)>50 then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				failed_subject = replace(failed_subject,vbcrlf,"")
				if failed_msg="" or isnull(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용을 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if checkNotValidHTML(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if sendmethod="KAKAOALRIM" then
					if instr(failed_subject,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자제목에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if instr(failed_msg,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if template_code="" or isnull(template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('카카오톡 알림톡 템플릿코드가 지정되어 있지 않습니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
				end if
			end if
		end if

		replacetagcode=""
		sqlStr = "SELECT" & vbcrlf
		sqlStr = sqlStr & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
		sqlStr = sqlStr & " and q.targetkey=N'"& targetkey &"'" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			replacetagcode 	= trim(db2html(rsget("replacetagcode")))
		End IF			
		rsget.Close	

		titletemp = title
		contentstemp = contents
		failed_subjecttemp = failed_subject
		failed_msgtemp = failed_msg

		if replacetagcode<>"" and not(isnull(replacetagcode)) then
			replacetagcodearray = split(replacetagcode,",")

			if isarray(replacetagcodearray) then
				for i = 0  to ubound(replacetagcodearray)
					replacetagcodetemp = trim(replacetagcodearray(i))
					if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
						titletemp = replace(titletemp,replacetagcodetemp,"")
						contentstemp = replace(contentstemp,replacetagcodetemp,"")
						failed_subjecttemp = replace(failed_subjecttemp,replacetagcodetemp,"")
						failed_msgtemp = replace(failed_msgtemp,replacetagcodetemp,"")
					end if
				next
			end if
		end if

		if instr(titletemp,"${")>0 or instr(contentstemp,"${")>0 or instr(failed_subjecttemp,"${")>0 or instr(failed_msgtemp,"${")>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		' 이상한 데이터가 있을까바 데이터 가공
		contents = replace(contents,"'","")

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
		if orderitemidexceptionarr <> "" then
			if checkNotValidHTML(orderitemidexceptionarr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('해당상품구매한사람제외에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
			end if
		end if
		if eventcodearr <> "" then
			if checkNotValidHTML(eventcodearr) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('이벤트번호에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
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
		if trim(orderitemidexceptionarr)<>"" then
			orderitemidexceptionarr = replace(replace(replace(trim(orderitemidexceptionarr),".",","),"""",""),"'","")		'오타 교정
			if left(orderitemidexceptionarr,1)="," then orderitemidexceptionarr = mid(orderitemidexceptionarr,2,len(orderitemidexceptionarr))		'오타 교정
			if right(orderitemidexceptionarr,1)="," then orderitemidexceptionarr = left(orderitemidexceptionarr,len(orderitemidexceptionarr)-1)		'오타 교정
		end if
		if trim(eventcodearr)<>"" then
			eventcodearr = replace(replace(replace(trim(eventcodearr),".",","),"""",""),"'","")		'오타 교정
			if left(eventcodearr,1)="," then eventcodearr = mid(eventcodearr,2,len(eventcodearr))		'오타 교정
			if right(eventcodearr,1)="," then eventcodearr = left(eventcodearr,len(eventcodearr)-1)		'오타 교정
		end if

       '' 타겟 대상이 바뀌면 재타겟을 해야함..
        sqlStr = "select targetkey from db_contents.dbo.tbl_lms_reserve with (readuncommitted)"
        sqlStr = sqlStr + " where ridx = "& ridx

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			Pretargetkey = rsget("targetkey")
			if isNULL(Pretargetkey) then Pretargetkey=""
		end if
		rsget.close
		
		if (CStr(Pretargetkey)<>CStr(targetkey)) then
		    sqlStr = "delete from db_contents.dbo.tbl_lms_TargetTemp where ridx="& ridx

            'response.write sqlStr & "<Br>"
            dbget.Execute sqlStr
		end if

		sqlStr = " update db_contents.dbo.tbl_lms_reserve" & VbCrlf
		sqlStr = sqlStr & " set sendmethod = N'"& sendmethod &"'" & VbCrlf
		sqlStr = sqlStr &"  , title = N'"& html2db(title) &"'" & VbCrlf
		sqlStr = sqlStr & " , contents = N'"& html2db(contents) &"'" & VbCrlf
		sqlStr = sqlStr & " , state = "& state &"" & VbCrlf
		sqlStr = sqlStr & " , reservedate = '"& reservedate &"'" & VbCrlf
		sqlStr = sqlStr & " , exception7dayyn = N'"& exception7dayyn &"'" & VbCrlf
		sqlStr = sqlStr & " , targetkey= "& targetkey &"" & VbCrlf

		if (CStr(Pretargetkey)<>CStr(targetkey)) then
		    sqlStr = sqlStr + " , targetcnt=0" + VbCrlf
		    sqlStr = sqlStr + " , targetstate=0" + VbCrlf
		end if

		sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
		sqlStr = sqlStr + " , lastupdate=getdate()"  + VbCrlf
		sqlStr = sqlStr & " , repeatlmsyn = '"& repeatlmsyn &"'" & VbCrlf
		sqlStr = sqlStr & " , member_smsok_checkyn = '"& member_smsok_checkyn &"'" & VbCrlf
		sqlStr = sqlStr & " , member_pushyn_checkyn = '"& member_pushyn_checkyn &"'" & VbCrlf
		sqlStr = sqlStr & " , makeridarr=N'"& makeridarr &"'" & VbCrlf
		sqlStr = sqlStr & " , itemidarr=N'"& itemidarr &"'" & VbCrlf
		sqlStr = sqlStr & " , keywordarr=N'"& keywordarr &"'" & VbCrlf
		sqlStr = sqlStr & " , bonuscouponidxarr=N'"& bonuscouponidxarr &"'" & VbCrlf
		sqlStr = sqlStr & " , button_name=N'"& html2db(button_name) &"'" & VbCrlf
		sqlStr = sqlStr & " , button_url_mobile=N'"& button_url_mobile &"'" & VbCrlf
		sqlStr = sqlStr & " , button_name2=N'"& html2db(button_name2) &"'" & VbCrlf
		sqlStr = sqlStr & " , button_url_mobile2=N'"& button_url_mobile2 &"'" & VbCrlf
		sqlStr = sqlStr & " , failed_type=N'"& failed_type &"'" & VbCrlf
		sqlStr = sqlStr & " , failed_subject=N'"& html2db(failed_subject) &"'" & VbCrlf
		sqlStr = sqlStr & " , failed_msg=N'"& html2db(failed_msg) &"'" & VbCrlf
		sqlStr = sqlStr & " , orderitemidexceptionarr=N'"& orderitemidexceptionarr &"'" & VbCrlf
		sqlStr = sqlStr & " , template_code=N'"& template_code &"'" & VbCrlf
		sqlStr = sqlStr & " , exceptionlogin=N'"& exceptionlogin &"'" & VbCrlf
		sqlStr = sqlStr & " , exceptionuserlevelarr=N'"& exceptionuserlevelarr &"'" & VbCrlf
		sqlStr = sqlStr & " , eventcodearr=N'"& eventcodearr &"'" & VbCrlf
		sqlStr = sqlStr & " , member_kakaoalrimyn_checkyn=N'"& member_kakaoalrimyn_checkyn &"'" & VbCrlf
		sqlStr = sqlStr & " , etc_template_code= N'"& etc_template_code &"' where" & VbCrlf
		sqlStr = sqlStr + " ridx = "& ridx

		'response.write sqlStr & "<br>"
		'response.end
		dbget.Execute sqlStr

		response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
		dbget.close()	:	response.End

	Case "state"
		if ridx = "" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		sqlStr = " update db_contents.dbo.tbl_lms_reserve" & VbCrlf
		sqlStr = sqlStr & " set state = '"& state &"' where" & VbCrlf
		sqlStr = sqlStr & " ridx = "& ridx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('발송상태가 변경 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	Case "del"
		sqlStr = " update db_contents.dbo.tbl_lms_reserve set " & VbCrlf
		sqlStr = sqlStr & " isusing = 'N'" & VbCrlf
		sqlStr = sqlStr & " , lastadminid='"& lastadminid &"'"  & VbCrlf
		sqlStr = sqlStr & " , lastupdate=getdate() where"  & VbCrlf
		sqlStr = sqlStr & " ridx = "& ridx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('사용여부가 변경 되었습니다.');</script>"
		session.codePage = 949
		'Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.self.close();</script>"
		dbget.close()	:	response.End

	' 테스트발송
	Case "test_lmsinsert"
		if (useridarr="" or isnull(useridarr)) then
			response.write "필수 값 체크 오류. 발송 아이디가 없습니다."
			session.codePage = 949
			dbget.close():response.end
		end if
		if (ridx="" or isnull(ridx)) then
			response.write "필수 값 체크 오류. 발송키가 없습니다."
			session.codePage = 949
			dbget.close():response.end
		end if

		title = trim(title)
		contents = trim(contents)

		if sendmethod="LMS" then
			if title="" or isnull(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목을 입력해 주세요.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			title = replace(title,vbcrlf,"")

			if checkNotValidHTML(title) then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
			if len(title)>120 then
				response.write "<script type='text/javascript'>"
				response.write "	alert('제목이 제한길이를 초과하였습니다. 120자 까지 작성 가능합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		end if
        if contents="" or isnull(contents) then
            response.write "<script type='text/javascript'>"
            response.write "	alert('내용을 입력해 주세요.');"
            response.write "</script>"
            session.codePage = 949
            dbget.close()	:	response.End
        end if
		'contents = replace(contents,vbcrlf,"\n")
		if sendmethod="KAKAOALRIM" then
			if instr(contents,"#{") then
				response.write "<script type='text/javascript'>"
				response.write "	alert('템플릿 내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
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
        button_name = trim(button_name)
        button_url_mobile = trim(button_url_mobile)
        button_name2 = trim(button_name2)
        button_url_mobile2 = trim(button_url_mobile2)
		failed_type = trim(failed_type)
		failed_subject = trim(failed_subject)
		failed_msg = trim(failed_msg)
		if sendmethod="KAKAOFRIEND" or sendmethod="KAKAOALRIM" then
			if button_name="" or isnull(button_name) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name = replace(button_name,vbcrlf,"")
			end if
			if button_url_mobile="" or isnull(button_url_mobile) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile = replace(button_url_mobile,vbcrlf,"")
			end if
			if button_name2="" or isnull(button_name2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_name2 = replace(button_name2,vbcrlf,"")
			end if
			if button_url_mobile2="" or isnull(button_url_mobile2) then
				' response.write "<script type='text/javascript'>"
				' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
				' response.write "</script>"
				' session.codePage = 949
				' dbget.close()	:	response.End
			else
				button_url_mobile2 = replace(button_url_mobile2,vbcrlf,"")
			end if
			if sendmethod="KAKAOALRIM" then
				' 수기템플릿
				if template_code="etc-9999" then
					if etc_template_code="" or isnull(etc_template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('수기템플릿코드를 입력해 주세요.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					template_code = etc_template_code
				else
					etc_template_code=""
				end if
			end if
			if failed_type="LMS" then
				if failed_subject="" or isnull(failed_subject) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목를 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if len(failed_subject)>50 then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				failed_subject = replace(failed_subject,vbcrlf,"")
				if failed_msg="" or isnull(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용을 입력해 주세요.');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if checkNotValidHTML(failed_msg) then
					response.write "<script type='text/javascript'>"
					response.write "	alert('카카오톡 실패시 문자내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
					response.write "</script>"
					session.codePage = 949
					dbget.close()	:	response.End
				end if
				if sendmethod="KAKAOALRIM" then
					if instr(failed_subject,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자제목에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if instr(failed_msg,"#{") then
						response.write "<script type='text/javascript'>"
						response.write "	alert('템플릿 카카오톡 실패시 문자내용에 #{XXX} 으로 지정되어 있는 부분은 직접 입력해 주셔야 합니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
					if template_code="" or isnull(template_code) then
						response.write "<script type='text/javascript'>"
						response.write "	alert('카카오톡 알림톡 템플릿코드가 지정되어 있지 않습니다.');"
						response.write "</script>"
						session.codePage = 949
						dbget.close()	:	response.End
					end if
				end if
			end if
		end if

		replacetagcode=""
		sqlStr = "SELECT" & vbcrlf
		sqlStr = sqlStr & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
		sqlStr = sqlStr & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " WHERE q.isusing=N'Y'" & vbcrlf
		sqlStr = sqlStr & " and q.targetkey=N'"& targetkey &"'" & vbcrlf

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			replacetagcode 	= trim(db2html(rsget("replacetagcode")))
		End IF			
		rsget.Close	

		titletemp = title
		contentstemp = contents
		failed_subjecttemp = failed_subject
		failed_msgtemp = failed_msg

		if replacetagcode<>"" and not(isnull(replacetagcode)) then
			replacetagcodearray = split(replacetagcode,",")

			if isarray(replacetagcodearray) then
				for i = 0  to ubound(replacetagcodearray)
					replacetagcodetemp = trim(replacetagcodearray(i))
					if replacetagcodetemp<>"" and not(isnull(replacetagcodetemp)) then
						titletemp = replace(titletemp,replacetagcodetemp,"")
						contentstemp = replace(contentstemp,replacetagcodetemp,"")
						failed_subjecttemp = replace(failed_subjecttemp,replacetagcodetemp,"")
						failed_msgtemp = replace(failed_msgtemp,replacetagcodetemp,"")
					end if
				next
			end if
		end if

		if instr(titletemp,"${")>0 or instr(contentstemp,"${")>0 or instr(failed_subjecttemp,"${")>0 or instr(failed_msgtemp,"${")>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('제목이나 내용에 사용이 불가능한 치환코드가 있습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		' 이상한 데이터가 있을까바 데이터 가공
		title = replace(title,"'","")
		contents = replace(contents,"'","")
		if failed_type="LMS" then
			failed_subject = replace(failed_subject,"'","")
			failed_msg = replace(failed_msg,"'","")
		end if

		useridarr = replace(useridarr,"'","")     
		useridarr = replace(useridarr,"""","")
		useridarr = replace(useridarr,".",",")
		useridarr = replace(useridarr,",,",",")
		useridarr = "'" & replace(useridarr,",","','") & "'"

		if sendmethod="LMS" then
			response.write "LMS 테스트 발송" & "<br>"

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
			sqlStr = sqlStr & " , replace(isnull(u.usercell,''),'-','') as usercell"
			sqlStr = sqlStr & " into #userInfo"
			sqlStr = sqlStr & " from db_user.dbo.tbl_user_n u with (nolock)"
			sqlStr = sqlStr & " left join db_user.dbo.tbl_logindata l with (nolock)"
			sqlStr = sqlStr & " 	on u.userid=l.userid"
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

			sqlStr = " select distinct" & vbcrlf
			sqlStr = sqlStr & " convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& title &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& title &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as SUBJECT" & vbcrlf
			sqlStr = sqlStr & " , usercell as PHONE, '1644-6030' as CALLBACK,'0' as STATUS,getdate() as REQDATE" & vbcrlf
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as MSG" & vbcrlf
			sqlStr = sqlStr & " ,'0' as FILE_CNT,'43200' as EXPIRETIME" & vbcrlf
			sqlStr = sqlStr & " into #tmpuser" & vbcrlf
			sqlStr = sqlStr & " from #userInfo as u" & vbcrlf
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

			sqlStr = "INSERT INTO LOGISTICSDB.db_LgSMS.dbo.MMS_MSG (SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME)" & vbcrlf
			sqlStr = sqlStr & " 	select * from #tmpuser" & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

		' 친구톡
		elseif sendmethod="KAKAOFRIEND" then
			response.write "카카오톡 친구톡 테스트 발송" & "<br>"

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
			sqlStr = sqlStr & " , replace(isnull(u.usercell,''),'-','') as usercell"
			sqlStr = sqlStr & " into #userInfo"
			sqlStr = sqlStr & " from db_user.dbo.tbl_user_n u with (nolock)"
			sqlStr = sqlStr & " left join db_user.dbo.tbl_logindata l with (nolock)"
			sqlStr = sqlStr & " 	on u.userid=l.userid"
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

			sqlStr = "select distinct" & vbcrlf
			sqlStr = sqlStr & " getdate() as REQDATE,N'1' as STATUS," & vbcrlf
			sqlStr = sqlStr & " usercell as PHONE," & vbcrlf		' 수신자 휴대폰 번호
			sqlStr = sqlStr & " N'1644-6030' as CALLBACK" & vbcrlf		' 발신자 번호
			' 알림톡 내용
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as MSG" & vbcrlf
			sqlStr = sqlStr & " ,N'0000000' as TEMPLATE_CODE" & vbcrlf		' 친구톡 템플릿 번호
			sqlStr = sqlStr & " ,N'"& failed_type &"' as FAILED_TYPE" & vbcrlf		' 알림톡 실패시 문자 형식 > SMS / LMS
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as FAILED_SUBJECT" & vbcrlf      ' 실패시 문자 제목 (LMS 전송시에만 필요)
			' 실패시 문자 내용
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as FAILED_MSG" & vbcrlf
			' 버튼 구성 내용 (버튼타입에만 필요 / v4 메뉴얼 참고)
			if button_name<>"" and button_url_mobile<>"" and button_name2<>"" and button_url_mobile2<>"" then
				sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""},{""name"":"""& button_name2 &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile2 &"""}]}' as BUTTON_JSON" & vbcrlf
			elseif button_name<>"" and button_url_mobile<>"" then
				sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""}]}' as BUTTON_JSON" & vbcrlf
			else
				sqlStr = sqlStr & " ,N'' as BUTTON_JSON" & vbcrlf
			end if
			sqlStr = sqlStr & " , userid, "& ridx &" as ridx" & vbcrlf
			sqlStr = sqlStr & " into #tmpuser" & vbcrlf
			sqlStr = sqlStr & " from #userInfo as u" & vbcrlf
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			'response.end
			dbget.Execute sqlStr

			' 발송 DB에 저장
			sqlStr = "INSERT INTO LOGISTICSDB.[db_kakaoMsg_v4_ft].dbo.[KKF_MSG] (REQDATE, STATUS, PHONE, CALLBACK, MSG, TEMPLATE_CODE, FAILED_TYPE,  FAILED_SUBJECT, FAILED_MSG, BUTTON_JSON, ETC1, ETC2)" & vbcrlf
			sqlStr = sqlStr & " 	select * from #tmpuser" & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

		' 알림톡
		elseif sendmethod="KAKAOALRIM" then
			response.write "카카오톡 알림톡 테스트 발송" & "<br>"

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
			sqlStr = sqlStr & " , replace(isnull(u.usercell,''),'-','') as usercell"
			sqlStr = sqlStr & " into #userInfo"
			sqlStr = sqlStr & " from db_user.dbo.tbl_user_n u with (nolock)"
			sqlStr = sqlStr & " left join db_user.dbo.tbl_logindata l with (nolock)"
			sqlStr = sqlStr & " 	on u.userid=l.userid"
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

			sqlStr = "SELECT distinct" & vbcrlf
			sqlStr = sqlStr & " getdate() as REQDATE,N'1' as STATUS," & vbcrlf
			sqlStr = sqlStr & " usercell as PHONE," & vbcrlf		' 수신자 휴대폰 번호
			sqlStr = sqlStr & " N'1644-6030' as CALLBACK" & vbcrlf	' 발신자 번호
			' 알림톡 내용
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& contents &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as MSG" & vbcrlf
			sqlStr = sqlStr & " ,N'"& template_code &"' as TEMPLATE_CODE" & vbcrlf	' 알림톡 템플릿 번호
			sqlStr = sqlStr & " ,N'"& failed_type &"' as FAILED_TYPE" & vbcrlf		' 알림톡 실패시 문자 형식 > SMS / LMS
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& failed_subject &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as FAILED_SUBJECT" & vbcrlf      ' 실패시 문자 제목 (LMS 전송시에만 필요)
			' 실패시 문자 내용
			sqlStr = sqlStr & " ,convert(nvarchar(max)," & vbcrlf
			sqlStr = sqlStr & "		replace(" & vbcrlf
			sqlStr = sqlStr & " 		replace(" & vbcrlf
			sqlStr = sqlStr & " 			replace(" & vbcrlf
			sqlStr = sqlStr & " 				replace(" & vbcrlf
			'sqlStr = sqlStr & " 					replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE WHEN LEN(userid)>1 THEN LEFT(userid,LEN(userid)-1)+N'*' ELSE userid END))" & vbcrlf
			sqlStr = sqlStr & " 					replace(N'"& failed_msg &"','${CUSTOMERID}',(CASE when isnull(u.userid,'')='' then '고객' ELSE isnull(u.userid,'') END))" & vbcrlf
			sqlStr = sqlStr & " 				,'${CUSTOMERNAME}',(CASE when isnull(u.username,'')='' then '고객' ELSE isnull(u.username,'') END))" & vbcrlf
			sqlStr = sqlStr & " 			,'${CUSTOMERLEVELNAME}',userlevelname)" & vbcrlf
			sqlStr = sqlStr & " 		,'${PRODUCTNAME}','TEST상품명')" & vbcrlf
			sqlStr = sqlStr & " 	,'${MILEAGE}','TEST마일리지')" & vbcrlf
			sqlStr = sqlStr & " ) as FAILED_MSG" & vbcrlf
			' 버튼 구성 내용 (버튼타입에만 필요 / v4 메뉴얼 참고)
			if button_name<>"" and button_url_mobile<>"" and button_name2<>"" and button_url_mobile2<>"" then
				sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""},{""name"":"""& button_name2 &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile2 &"""}]}' as BUTTON_JSON" & vbcrlf
			elseif button_name<>"" and button_url_mobile<>"" then
				sqlStr = sqlStr & " ,N'{""button"":[{""name"":"""& button_name &""",""type"":""WL"", ""url_mobile"":"""& button_url_mobile &"""}]}' as BUTTON_JSON" & vbcrlf
			else
				sqlStr = sqlStr & " ,N'' as BUTTON_JSON" & vbcrlf
			end if
			sqlStr = sqlStr & " , userid, "& ridx &" as ridx" & vbcrlf
			sqlStr = sqlStr & " into #tmpuser" & vbcrlf
			sqlStr = sqlStr & " from #userInfo as u" & vbcrlf
			sqlStr = sqlStr & " where u.userid in ("& useridarr &") " & vbcrlf

			'response.write sqlStr & "<br>"
			'response.end
			dbget.Execute sqlStr

			' 발송 DB에 저장
			sqlStr = "INSERT INTO LOGISTICSDB.[db_kakaomsg_v4_mkt].dbo.KKO_MSG (REQDATE, STATUS, PHONE, CALLBACK, MSG, TEMPLATE_CODE, FAILED_TYPE, FAILED_SUBJECT, FAILED_MSG, BUTTON_JSON, ETC1, ETC2)" & vbcrlf
			sqlStr = sqlStr & " 	select * from #tmpuser" & vbcrlf

			'response.write sqlStr & "<br>"
			dbget.Execute sqlStr

		end if

		response.write "==================================<br>"
		response.write "테스트 메시지 발송요청 되었습니다.<br>"

		'//테스트카운트 올림
		sqlStr = " update db_contents.dbo.tbl_lms_reserve set " + VbCrlf
		sqlStr = sqlStr + " testsend = testsend + "& ubound(split(useridarr,","))+1 &" " + VbCrlf
		sqlStr = sqlStr + " where ridx = "& ridx &""

		'response.write sqlStr
		dbget.Execute sqlStr

		'Response.write "sendedMsg:"&sendedMsg
		Response.write "<br/><input type='button' onclick='opener.location.reload();self.close();' value='닫기'/>"
		session.codePage = 949
		dbget.close()	:	response.End

	'/타게팅
	Case "target"	
		if ridx="" or isnull(ridx) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('정상적인 경로가 아닙니다. 발송번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		set olms = new clms_msg_list
			olms.FRectrIdx = ridx
			olms.Frectisusing = "Y"
			olms.lmsmsg_getrow()

			if olms.FResultCount > 0 then			
				targetkey			= olms.FOneItem.ftargetkey
			else
				response.write "<script type='text/javascript'>"
				response.write "	alert('발송번호에 해당되는 내역이 없습니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		set olms = Nothing

		if targetkey="" or isnull(targetkey) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('지정된 타켓이 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		Set clsCode = new ClmstargetCommonCode  	
		clsCode.frecttargetkey  = targetkey
		clsCode.Frectisusing = "Y"
		clsCode.GetlmstargetCont

		if clsCode.FTotalCount>0 THEN
			targetQuery = clsCode.ftargetQuery
		end if
		Set clsCode = nothing 

		if targetQuery="" or isnull(targetQuery) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('지정된 타켓쿼리가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		sqlStr = replace(replace(targetQuery,"${RIDX}",ridx),"${ADMINID}",lastadminid)

		'response.write sqlStr & "<Br>"
		dbget.CommandTimeout = 60*10   ' 10분
		dbget.Execute sqlStr

		response.write "<script type='text/javascript'>alert('타겟 설정 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	'/관리자 리타게팅
	Case "retarget"	
	    sqlStr = "update db_contents.dbo.tbl_lms_reserve set targetState=0 where ridx="&ridx

		'response.write sqlStr & "<Br>"
	    dbget.Execute sqlStr

		if ridx="" or isnull(ridx) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('정상적인 경로가 아닙니다. 발송번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		set olms = new clms_msg_list
			olms.FRectrIdx = ridx
			olms.Frectisusing = "Y"
			olms.lmsmsg_getrow()

			if olms.FResultCount > 0 then			
				targetkey			= olms.FOneItem.ftargetkey
			else
				response.write "<script type='text/javascript'>"
				response.write "	alert('발송번호에 해당되는 내역이 없습니다.');"
				response.write "</script>"
				session.codePage = 949
				dbget.close()	:	response.End
			end if
		set olms = Nothing

		if targetkey="" or isnull(targetkey) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('지정된 타켓이 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		Set clsCode = new ClmstargetCommonCode  	
		clsCode.frecttargetkey  = targetkey
		clsCode.Frectisusing = "Y"
		clsCode.GetlmstargetCont

		if clsCode.FTotalCount>0 THEN
			targetQuery = clsCode.ftargetQuery
		end if
		Set clsCode = nothing 

		if targetQuery="" or isnull(targetQuery) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('지정된 타켓쿼리가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		sqlStr = replace(replace(targetQuery,"${RIDX}",ridx),"${ADMINID}",lastadminid)

		response.write sqlStr & "<Br>"
		dbget.CommandTimeout = 60*10   ' 10분
		dbget.Execute sqlStr

		response.write "<script type='text/javascript'>alert('재타겟 설정 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	'/타게팅삭제
	Case "targetdel"	
		if ridx="" or isnull(ridx) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('정상적인 경로가 아닙니다. 발송번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		sqlStr = "delete from db_contents.dbo.tbl_lms_TargetTemp where ridx="& ridx

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr
		
	    sqlStr = "update db_contents.dbo.tbl_lms_reserve set targetState=0, targetcnt=0 where ridx="&ridx

		'response.write sqlStr & "<Br>"
	    dbget.Execute sqlStr

		response.write "<script type='text/javascript'>alert('타겟이 리셋 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
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