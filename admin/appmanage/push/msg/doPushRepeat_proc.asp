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
%>
<%
'###########################################################
' Description : 푸시 반복 관리
' Hieditor : 2019.05.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim repeatidx, reservationdate , state , mode , reservetime , reservemin , lastadminid, pushimg, pushimg2, pushimg3, pushimg4, pushimg5
Dim reservedate, gaparam, ppidx, makeridarr, itemidarr, keywordarr, bonuscouponidxarr, notclickyn, pushcontents
dim i, userid, utmparam, sendedMsg, useridarr, isMsgArr, ParamMsg, objItem, len1, len2, len3, MaxValidLen, stitle , subtitle, sqlStr
dim appkey, message, deviceid, addCount, addParamMsg,admcomment, kk, iparam, iparamvalue, imgtype, sendranking
dim noduppDate,targetKey, iADODBcmd, intResult, daterepeatgubun, countrepeatgubun, yyyy, mm, dd, time1, time2
dim daterepeatgubunarr, countrepeatgubunarr, yyyyarr, mmarr, ddarr, reservetimearr, reserveminarr, reservecountrepeatgubun
dim replacetagcode, titletemp,contentstemp
dim CUSTOMERIDYN, CUSTOMERNAMEYN, CUSTOMERID,CUSTOMERNAME, replacetagcodearray, replacetagcodetemp, privateYN
	useridarr = requestcheckvar(request("useridarr"),256)
	repeatidx = requestcheckvar(request("repeatidx"),10)
	makeridarr = request("makeridarr")
	itemidarr = request("itemidarr")
	keywordarr = request("keywordarr")
	bonuscouponidxarr = request("bonuscouponidxarr")
	notclickyn = requestcheckvar(request("notclickyn"),10)
	appkey = requestcheckvar(request("appkey"),10)
	message = requestcheckvar(request("message"),800)
	deviceid = LEFT(replace(request("deviceid"),"'",""),200)  '' -- 이 치환 되면 안됨..
	stitle			= RequestCheckVar(request("stitle"),800)
	pushcontents			= RequestCheckVar(request("pushcontents"),3000)
	subtitle		= RequestCheckVar(request("subtitle"),500)
	reservetime		= RequestCheckVar(request("time1"),800)
	reservemin		= RequestCheckVar(request("time2"),800)
	state			= RequestCheckVar(request("state"),2)
	admcomment      = RequestCheckVar(request("admcomment"),200)
	noduppDate      = RequestCheckVar(request("noduppDate"),10)
	targetKey       = RequestCheckVar(request("targetKey"),10)
	sendranking       = RequestCheckVar(getNumeric(request("sendranking")),10)
	daterepeatgubun       = RequestCheckVar(request("daterepeatgubun"),800)
	countrepeatgubun       = RequestCheckVar(request("countrepeatgubun"),800)
	yyyy       = RequestCheckVar(request("yyyy"),800)
	mm       = RequestCheckVar(request("mm"),800)
	dd       = RequestCheckVar(request("dd"),800)

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

if (addParamMsg<>"") then
	addParamMsg = Left(addParamMsg,Len(addParamMsg)-1)
end if
if sendranking="" then sendranking="6"
if noduppDate="" then noduppDate="0"
if notclickyn="on" then
	notclickyn="Y"
else
	notclickyn="N"
end if

pushimg = RequestCheckVar(request("pushimg"),200)
pushimg2 = RequestCheckVar(request("pushimg2"),200)
pushimg3 = RequestCheckVar(request("pushimg3"),200)
pushimg4 = RequestCheckVar(request("pushimg4"),200)
pushimg5 = RequestCheckVar(request("pushimg5"),200)
mode = RequestCheckVar(request("mode"),32)
imgtype = RequestCheckVar(request("imgtype"),10)

' PUSH메시지 크기
''' ios 메세지(json) 의 총길이는 256 바이트를 넘을 수 없음
'' pushtitle + pushurl 길이를 제한 <= 160 (169 까지는 나갔음..)
'MaxValidLen = 186 ''160  2016/11/09

' iOS 8 이상, Android 4 이상에서 전체 전송 제한 크기는 4Kb로 확인됨
'   다만 UTF8 한글 특성상 한글 한글자는 3byte가 할당되므로 1300자 정도 넣을 있으며
'   통신에 기타 해더및 추가 정보를 빼면 약 800자정도 넣을 수 있는것으로 판단됨
MaxValidLen = 800		' 2018.08.21
if (mode="repeatInsert") or (mode="repeatmEdit") then
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
	' end If
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
	Case "repeatInsert"
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

		sqlStr = " insert into db_contents.dbo.tbl_app_push_repeat (" + VbCrlf
		sqlStr = sqlStr & " pushtitle, pushcontents, pushurl, pushimg, imgtype, state, testpush, isusing, noduppDate, targetKey, admcomment, targetstate, mayTargetCnt" + VbCrlf
		sqlStr = sqlStr & " , makeridarr, itemidarr, keywordarr, bonuscouponidxarr, notclickyn, regadminid, lastadminid, regdate, lastupdate, pushimg2, pushimg3, pushimg4" + VbCrlf
		sqlStr = sqlStr & " , pushimg5, sendranking, privateYN) values (" & VbCrlf
		sqlStr = sqlStr + " N'" + stitle + "' ,N'" + pushcontents + "', N'" + subtitle + "' , N'"& pushimg &"', "& imgtype &"," + state + ", 0, N'Y',"& noduppDate & "" + VbCrlf
		sqlStr = sqlStr + " ,"&CHKIIF(targetKey="","NULL",targetKey)&", N'"&html2db(admcomment)&"', 0, 0, N'"& makeridarr &"', N'"& itemidarr &"', N'"& keywordarr &"'" + VbCrlf
		sqlStr = sqlStr & " , N'"& bonuscouponidxarr &"', '"& notclickyn &"', N'"& lastadminid &"', N'"& lastadminid &"', getdate(), getdate(), N'"& pushimg2 &"'" & VbCrlf
		sqlStr = sqlStr & " , N'"& pushimg3 &"', N'"& pushimg4 &"', N'"& pushimg5 &"', "& sendranking &",'"& privateYN &"'" & VbCrlf
		sqlStr = sqlStr + " )" + VbCrlf

		'response.write sqlStr & "<Br>"
		dbget.Execute sqlStr

		'' gaparam 추가 2016/11/09 ------------------------------------------------------------------
		if InStr(subtitle,"gaparam")<1 then
			gaparam = ""
			
			sqlStr = "select top 1 repeatidx, isNULL(targetKey,0) as targetKey from db_contents.dbo.tbl_app_push_repeat where pushurl='"&subtitle&"' order by repeatidx desc"

			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				ppidx = rsget("repeatidx")
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
				
				sqlStr = " update db_contents.dbo.tbl_app_push_repeat" + VbCrlf
				sqlStr = sqlStr + " set pushurl=N'"&subtitle&gaparam&"'"  + VbCrlf
				sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
				sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
				sqlStr = sqlStr + " repeatidx="&ppidx
				
				dbget.Execute sqlStr
			end if
		end if

		daterepeatgubun = trim(daterepeatgubun)
		if daterepeatgubun <> "" then
			daterepeatgubunarr = split(daterepeatgubun,",")
			countrepeatgubunarr = split(countrepeatgubun,",")
			yyyyarr = split(yyyy,",")
			mmarr = split(mm,",")
			ddarr = split(dd,",")
			reservetimearr = split(reservetime,",")
			reserveminarr = split(reservemin,",")
			
			if isarray(daterepeatgubunarr) then
				for i = 0 to ubound(daterepeatgubunarr)
					' 수행구분 : 일별
					if trim(daterepeatgubunarr(i))="1" then
						reservationdate = dateserial( year(date()), month(date()), day(date()) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 월별
					elseif trim(daterepeatgubunarr(i))="2" then
						reservationdate = dateserial( year(date()), month(date()), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 년별
					elseif trim(daterepeatgubunarr(i))="3" then
						reservationdate = dateserial( year(date()), trim(mmarr(i)), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 상시
					elseif trim(daterepeatgubunarr(i))="4" then
						reservationdate = date()
						reservedate		= ""
						reservecountrepeatgubun = ""
					else
						reservationdate = dateserial( year(date()), trim(mmarr(i)), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))
					end if

					sqlStr = " insert into db_contents.dbo.tbl_app_push_repeat_schedule"&VbCRLF
					sqlStr = sqlStr & " (repeatidx, daterepeatgubun, countrepeatgubun, repeatdate) values"&VbCRLF
					sqlStr = sqlStr & " ("& ppidx &","& trim(daterepeatgubunarr(i)) &",'"& reservecountrepeatgubun &"', '"& reservedate &"')"

					'response.write sqlStr &"<br>"
					'response.end
					dbget.execute sqlStr
				next
			end if
		end if

	'	dim referer
	'	referer = request.ServerVariables("HTTP_REFERER")
		response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
		dbget.close()	:	response.End

	Case "repeatmEdit"
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

		sqlStr = " update db_contents.dbo.tbl_app_push_repeat set " + VbCrlf
		sqlStr = sqlStr + " pushtitle = N'"& stitle &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushurl = N'"& subtitle &"'" + VbCrlf
		sqlStr = sqlStr + " ,state = "& state &"" + VbCrlf
		sqlStr = sqlStr + " ,pushimg = N'"& pushimg &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg2 = N'"& pushimg2 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg3 = N'"& pushimg3 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg4 = N'"& pushimg4 &"'" + VbCrlf
		sqlStr = sqlStr + " ,pushimg5 = N'"& pushimg5 &"'" + VbCrlf
		sqlStr = sqlStr + " ,imgtype = "& imgtype &"" + VbCrlf
		sqlStr = sqlStr + " ,admcomment= N'"& html2db(admcomment) &"'" + VbCrlf
		sqlStr = sqlStr + " ,noduppDate= "& noduppDate &"" + VbCrlf
		sqlStr = sqlStr + " ,targetKey= "& CHKIIF(targetKey="","NULL",targetKey) &"" + VbCrlf
		sqlStr = sqlStr + " ,notclickyn= '"& notclickyn &"'" + VbCrlf
		sqlStr = sqlStr & " , makeridarr=N'"& makeridarr &"'" & VbCrlf
		sqlStr = sqlStr & " , itemidarr=N'"& itemidarr &"'" & VbCrlf
		sqlStr = sqlStr & " , keywordarr=N'"& keywordarr &"'" & VbCrlf
		sqlStr = sqlStr & " , bonuscouponidxarr=N'"& bonuscouponidxarr &"'" & VbCrlf
		sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
		sqlStr = sqlStr + " , lastupdate=getdate()"  + VbCrlf
		sqlStr = sqlStr + " , sendranking = "& sendranking &"" + VbCrlf
		sqlStr = sqlStr + " , pushcontents = N'"& pushcontents &"'" + VbCrlf
		sqlStr = sqlStr + " , privateYN = '"& privateYN &"' where" + VbCrlf
		sqlStr = sqlStr + " repeatidx = "& repeatidx

		'response.write sqlStr & "<br>"
		'response.end
		dbget.Execute sqlStr

		daterepeatgubun = trim(daterepeatgubun)
		if daterepeatgubun <> "" then
			daterepeatgubunarr = split(daterepeatgubun,",")
			countrepeatgubunarr = split(countrepeatgubun,",")
			yyyyarr = split(yyyy,",")
			mmarr = split(mm,",")
			ddarr = split(dd,",")
			reservetimearr = split(reservetime,",")
			reserveminarr = split(reservemin,",")
			
			if isarray(daterepeatgubunarr) then
				sqlStr = "delete from db_contents.dbo.tbl_app_push_repeat_schedule where repeatidx="& repeatidx &""&VbCRLF

				'response.write sqlStr &"<br>"
				dbget.execute sqlStr

				for i = 0 to ubound(daterepeatgubunarr)
					' 수행구분 : 일별
					if trim(daterepeatgubunarr(i))="1" then
						reservationdate = dateserial( year(date()), month(date()), day(date()) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 월별
					elseif trim(daterepeatgubunarr(i))="2" then
						reservationdate = dateserial( year(date()), month(date()), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 년별
					elseif trim(daterepeatgubunarr(i))="3" then
						reservationdate = dateserial( year(date()), trim(mmarr(i)), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))

					' 수행구분 : 상시
					elseif trim(daterepeatgubunarr(i))="4" then
						reservationdate = date()
						reservedate		= ""
						reservecountrepeatgubun = ""
					else
						reservationdate = dateserial( year(date()), trim(mmarr(i)), trim(ddarr(i)) )
						reservedate		= reservationdate &" "& reservetimearr(i) &":"& reserveminarr(i) &":000" '예약일
						reservecountrepeatgubun = trim(countrepeatgubunarr(i))
					end if

					sqlStr = " insert into db_contents.dbo.tbl_app_push_repeat_schedule"&VbCRLF
					sqlStr = sqlStr & " (repeatidx, daterepeatgubun, countrepeatgubun, repeatdate) values"&VbCRLF
					sqlStr = sqlStr & " ("& repeatidx &","& trim(daterepeatgubunarr(i)) &",'"& reservecountrepeatgubun &"', '"& reservedate &"')"

					'response.write sqlStr &"<br>"
					dbget.execute sqlStr
				next
			end if
		end if

		'' gaparam 추가 2016/11/09 ------------------------------------------------------------------
		if InStr(subtitle,"gaparam")<1 then
			gaparam = ""
			
			sqlStr = "select top 1 repeatidx, isNULL(targetKey,0) as targetKey from db_contents.dbo.tbl_app_push_repeat where repeatidx="&repeatidx&" order by repeatidx desc"

			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				ppidx = rsget("repeatidx")
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
				
				sqlStr = " update db_contents.dbo.tbl_app_push_repeat" + VbCrlf
				sqlStr = sqlStr + " set pushurl=N'"&subtitle&gaparam&"'"  + VbCrlf
				sqlStr = sqlStr + " , lastadminid=N'"& lastadminid &"'"  + VbCrlf
				sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
				sqlStr = sqlStr + " repeatidx="&repeatidx
				
				dbget.Execute sqlStr
			end if
		end if

		response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
		dbget.close()	:	response.End
	
	Case "state"
		if repeatidx = "" then
			response.write "<script type='text/javascript'>"
			response.write "	alert('푸시번호가 없습니다.');"
			response.write "</script>"
			session.codePage = 949
			dbget.close()	:	response.End
		end if

		'' 조건 추가 state0=>1 로 변경시(발송예약), 타겟 상태 check
		'' 타겟 발송인경우 istargetMsg=1 & targetstate=0 & targetKey<>9999(수기) & 발송예약일이 6시간 이후인경우 타겟 예약상태로 변경함.
		'' 발송 엔진에서 타겟 상태가 타겟 예약인경우 발송 30분 전에 타게팅함.
		
		sqlStr = " update db_contents.dbo.tbl_app_push_repeat set " + VbCrlf
		sqlStr = sqlStr + " state = '"& state &"' where" + VbCrlf
		sqlStr = sqlStr + " repeatidx = "& repeatidx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('발송상태가 변경 되었습니다.');</script>"
		session.codePage = 949
		Response.write "<script type='text/javascript'>parent.opener.location.reload();parent.location.reload();</script>"
		dbget.close()	:	response.End

	Case "del"
		sqlStr = " update db_contents.dbo.tbl_app_push_repeat set " + VbCrlf
		sqlStr = sqlStr + " isusing = 'N'" + VbCrlf
		sqlStr = sqlStr + " , lastadminid='"& lastadminid &"'"  + VbCrlf
		sqlStr = sqlStr + " , lastupdate=getdate() where"  + VbCrlf
		sqlStr = sqlStr + " repeatidx = "& repeatidx

		dbget.Execute sqlStr
		
		response.write "<script type='text/javascript'>alert('사용여부가 변경 되었습니다.');</script>"
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