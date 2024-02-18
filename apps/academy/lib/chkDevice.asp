<%
	Dim nowip : nowip =  Request.ServerVariables("REMOTE_HOST")

	Dim uAgent, flgDevice, flgDevicePC
	'///// 접속기종 및 브라우져 종류 가져오기(모바일에서만 ipad 추가;2012.05.23) /////
	uAgent = Lcase(Request.ServerVariables("HTTP_USER_AGENT"))
	if instr(uAgent,"windows ce")>0 or instr(uAgent,"lgtelecom")>0 or instr(uAgent,"midp")>0 or instr(uAgent,"wipi")>0 or instr(uAgent,"android")>0 or instr(uAgent,"ipod")>0 or instr(uAgent,"iphone")>0 or instr(uAgent,"ipad")>0 or instr(uAgent,"playstation") or instr(uAgent,"blackberry") then
		'휴대기기
		if instr(uAgent,"ppc")>0 or instr(uAgent,"iemobile")>0 then
			flgDevice = "P"	'PDA
		elseif instr(uAgent,"ipod")>0 or instr(uAgent,"iphone")>0 or instr(uAgent,"ipad")>0 then
			flgDevice = "I" 'iPhone,iPod,iPad
		elseif instr(uAgent,"android")>0 then
			flgDevice = "A" 'Android
		else
			flgDevice = "M"	'Mobile
		end if
	else
		'일반
		flgDevice = "W"
	end If

	if instr(uAgent,"windows nt")>0 or instr(uAgent,"windows xp")>0 or instr(uAgent,"mac os x")>0 then
		if instr(uAgent,"windows nt")>0 or instr(uAgent,"windows xp")>0 then
			flgDevicePC = "W"
		elseif instr(uAgent,"mac os x")>0 then
			flgDevicePC = "M"
		end if
	end if
%>
