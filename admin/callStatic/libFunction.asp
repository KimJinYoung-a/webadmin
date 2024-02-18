<%

'내선번호 옵션박스
Sub DrawInlinePhoneBox(byval extension)
	dim buf,i

	buf = "<select class='select' name='extension'>"
	if (""=CStr(extension)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
    for i=801 to 807
		if (CStr(i)=CStr(extension)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    for i=901 to 911
		if (CStr(i)=CStr(extension)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 시간대
Sub DrawCallcenterHourBox(byval hour_from, hour_to)
	dim buf,i

	buf = "<select class='select' name='hour_from'>"
	if (""=CStr(hour_from)) then
		buf = buf + "<option value='' selected>--</option>"
	else
		buf = buf + "<option value='' >--</option>"
    end if
    for i=9 to 18
		if (CStr(i)=CStr(hour_from)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + " - "

	buf = buf + "<select class='select' name='hour_to'>"
	if (""=CStr(hour_to)) then
		buf = buf + "<option value='' selected>--</option>"
	else
		buf = buf + "<option value='' >--</option>"
    end if
    for i=9 to 18
		if (CStr(i)=CStr(hour_to)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 수발신 옵션박스
Sub DrawCallcenterInOutStateBox(byval dcontext)
	dim buf,i

	buf = "<select class='select' name='dcontext'>"
	if (""=CStr(dcontext)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("inbound"=CStr(dcontext)) then
		buf = buf + "<option value='inbound' selected>수신전화</option>"
	else
		buf = buf + "<option value='inbound' >수신전화</option>"
    end if
	if ("outbound"=CStr(dcontext)) then
		buf = buf + "<option value='outbound' selected>발신전화</option>"
	else
		buf = buf + "<option value='outbound' >발신전화</option>"
    end if
	if ("hunt_context"=CStr(dcontext)) then
		buf = buf + "<option value='hunt_context' selected>헌트연결</option>"
	else
		buf = buf + "<option value='hunt_context' >헌트연결</option>"
    end if
	if ("pers_context"=CStr(dcontext)) then
		buf = buf + "<option value='pers_context' selected>개인직통</option>"
	else
		buf = buf + "<option value='pers_context' >개인직통</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 상태 옵션박스
Sub DrawCallcenterModeBox(byval mode)
	dim buf,i

	buf = "<select class='select' name='mode'>"
	if ("all"=CStr(mode)) then
		buf = buf + "<option value='all' selected>전부보기</option>"
	else
		buf = buf + "<option value='all' >전부보기</option>"
    end if
	
	if ("try"=CStr(mode)) then
		buf = buf + "<option value='try' selected>시도전화(전체)</option>"
	else
		buf = buf + "<option value='try' >시도전화(전체)</option>"
    end if
    
	if ("trycall"=CStr(mode)) then
		buf = buf + "<option value='trycall' selected>시도전화(콜전체)</option>"
	else
		buf = buf + "<option value='trycall' >시도전화(콜전체)</option>"
    end if
    
	if ("trycallnotplay"=CStr(mode)) then
		buf = buf + "<option value='trycallnotplay' selected>시도전화(콜,근무시간)</option>"
	else
		buf = buf + "<option value='trycallnotplay' >시도전화(콜,근무시간)</option>"
    end if
    
	if ("trycallonlyplay"=CStr(mode)) then
		buf = buf + "<option value='trycallonlyplay' selected>시도전화(콜,근무시간외)</option>"
	else
		buf = buf + "<option value='trycallonlyplay' >시도전화(콜,근무시간외)</option>"
    end if
    
	if ("successall"=CStr(mode)) then
		buf = buf + "<option value='successall' selected>성공전화(전체)</option>"
	else
		buf = buf + "<option value='successall' >성공전화(전체)</option>"
    end if
    
	if ("successcall"=CStr(mode)) then
		buf = buf + "<option value='successcall' selected>성공전화(콜센터)</option>"
	else
		buf = buf + "<option value='successcall' >성공전화(콜센터)</option>"
    end if
    
	if ("successnotcall"=CStr(mode)) then
		buf = buf + "<option value='successnotcall' selected>성공전화(콜센터제외)</option>"
	else
		buf = buf + "<option value='successnotcall' >성공전화(콜센터제외)</option>"
    end if
    
	if ("success2"=CStr(mode)) then
		buf = buf + "<option value='success2' selected>성공전화(제공받음)</option>"
	else
		buf = buf + "<option value='success2' >성공전화(제공받음)</option>"
    end if
    
	if ("outcall"=CStr(mode)) then
		buf = buf + "<option value='outcall' selected>콜센터발신전화</option>"
	else
		buf = buf + "<option value='outcall' >콜센터발신전화</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 답변여부 옵션박스
Sub DrawCallcenterAnswerStateBox(byval dispositiono)
	dim buf,i

	buf = "<select class='select' name=dispositiono'>"
	if (""=CStr(dispositiono)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("ANSWERED9"=CStr(dispositiono)) then
		buf = buf + "<option value='ANSWERED' selected>ANSWERED</option>"
	else
		buf = buf + "<option value='ANSWERED' >ANSWERED</option>"
    end if
	if ("BUSY"=CStr(disposition)) then
		buf = buf + "<option value='BUSY' selected>BUSY</option>"
	else
		buf = buf + "<option value='BUSY' >BUSY</option>"
    end if
	if ("FAILED"=CStr(disposition)) then
		buf = buf + "<option value='FAILED' selected>FAILED</option>"
	else
		buf = buf + "<option value='FAILED' >FAILED</option>"
    end if
	if ("NO ANSWER"=CStr(disposition)) then
		buf = buf + "<option value='NO ANSWER' selected>NO ANSWER</option>"
	else
		buf = buf + "<option value='FAILED' >NO ANSWER</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 전화번호 옵션박스
Sub DrawCallcenterPhoneNameBox(byval phoneno)
	dim buf,i

	buf = "<select class='select' name='phoneno'>"
	if (""=CStr(phoneno)) then
		buf = buf + "<option value='' selected>ALL</option>"
	else
		buf = buf + "<option value='' >ALL</option>"
    end if
	if ("07075490429"=CStr(phoneno)) then
		buf = buf + "<option value='07075490429' selected>콜센터헌트</option>"
	else
		buf = buf + "<option value='07075490429' >콜센터헌트</option>"
    end if
	if ("07075490556"=CStr(phoneno)) then
		buf = buf + "<option value='07075490556' selected>사무실헌트</option>"
	else
		buf = buf + "<option value='07075490556' >사무실헌트</option>"
    end if
	if ("07075490449"=CStr(phoneno)) then
		buf = buf + "<option value='07075490449' selected>대표번호2</option>"
	else
		buf = buf + "<option value='07075490449' >대표번호2</option>"
    end if
	if ("07075490448"=CStr(phoneno)) then
		buf = buf + "<option value='07075490448' selected>대표번호1</option>"
	else
		buf = buf + "<option value='07075490448' >대표번호1</option>"
    end if
    buf = buf + "</select>"

    response.write buf
end Sub

'콜센터 수발신 문자열
Sub PrintCallcenterInOutState(byval dcontext)
	dim buf

	if ("inbound"=CStr(dcontext)) then
		buf = "수신"
	elseif ("outbound"=CStr(dcontext)) then
    	buf = "발신"
	elseif ("toexten"=CStr(dcontext)) then
		'buf = "내선"
		buf = CStr(dcontext)
    elseif ("hunt_context"=CStr(dcontext)) then
		buf = "헌트"
    elseif ("pers_context"=CStr(dcontext)) then
		buf = "개인"
    else
    	buf = CStr(dcontext)
    end if

    response.write buf
end Sub

'콜센터 마지막 상태 문자열
Sub PrintCallcenterLastState(byval lastapp)
	dim buf

	if ("Playback"=CStr(lastapp)) then
		buf = "안내멘트"
	elseif ("Hangup"=CStr(lastapp)) then
    	buf = "통화종료"
	elseif ("Dial"=CStr(lastapp)) then
		buf = "통화연결"
	elseif ("BackGround"=CStr(lastapp)) then
		buf = "대기멘트"
	elseif ("WaitExten"=CStr(lastapp)) then
		buf = "내선대기"
	elseif ("Busy"=CStr(lastapp)) then
		buf = "연결대기"
	else
		buf = CStr(lastapp)
    end if

    response.write buf
end Sub

'콜센터 전화번호 문자열
Sub PrintCallcenterPhoneNumberString(byval dst)
	dim buf

	if ("07075490429"=CStr(dst)) then
		buf = "콜센터헌트"
	elseif ("07075490556"=CStr(dst)) then
    	buf = "사무실헌트"
	elseif ("Dial"=CStr(dst)) then
		buf = "통화연결"
	elseif ("07075490449"=CStr(dst)) then
		buf = "대표번호2"
	elseif ("07075490448"=CStr(dst)) then
		buf = "대표번호1"
	else
		buf = CStr(dst)
    end if

    response.write buf
end Sub

function SectoTime(v)
	dim temp, h, m, s
	if v > (60*60*24) then
		v = v mod 24
	end if
	h = int(v/3600)
	temp = v mod 3600
	m = int(temp/60)
	s = temp mod 60

	if (h < 10) then h = "0" & h
	if (m < 10) then m = "0" & m
	if (s < 10) then s = "0" & s

	sectotime = h &":"& m &":"& s
end function

%>