<footer>
	<div class="btnGroup">
		<%
		If instr(nowViewPage,"login")<=0 Then
			if Not(IsUserLoginOK) and Not(IsGuestLoginOK) then
				Response.Write "<a href=""" & SSLUrl & "/login/login.asp?backpath=" & Server.URLEncode(CurrURLQ()) & """ class=""btn btnWht"">로그인</a>"
			Else
				Response.Write "<a href=""javascript:TnLogOut();"" class=""btn btnWht"">로그아웃</a>"
			End If
		End If
		%>
		<a href="/partnership/partnership_writer.asp" class="btn btnWht">작가/강사 신청 문의</a>
		<a href="http://www.thefingers.co.kr/?mfg=pc" class="btn btnWht" target="_blank">PC 버전</a>
	</div>
	<div class="footNav">
		<a href="/cscenter/pop_NoticeList.asp">공지사항</a> <span class="bar">ㅣ</span> 
		<a href="/member/private.asp">개인정보 처리방침</a> <span class="bar">ㅣ</span> 
		<a href="/member/viewUsageTerms.asp">이용약관</a> <span class="bar">ㅣ</span> 
		<a href="http://m.10x10.co.kr/" target="_blank">텐바이텐</a>
	</div>
	<h1>THE FINGERS</h1>
	<div class="footInfo">
		<p><a href="tel:16441557" class="cRed1"><strong>TEL 1644-1557</strong></a> <span class="bar">ㅣ</span> <a href="mailto:customer@thefingers.co.kr"><strong>customer@thefingers.co.kr</strong></a></p>
		<p>(평일 09:00-18:00 / 점심시간 12:00 ~ 13:00)</p>
		<p>주말 및 공휴일은 1:1 상담을 이용해주세요.</p>
		<p>이용 시간 외 강좌 관련 상담 : 02-741-9070</p>
		<address class="tMar1r">(03086) 서울 종로구 대학로12길31 자유빌딩 2층 (주)텐바이텐</address>
		<p>대표이사 : 최은희 <span class="bar">ㅣ</span> 개인정보보호 및 청소년 보호 책임자 : 이문재</p>
		<p>사업자등록번호 : 211-87-00620 <a href="http://www.ftc.go.kr/info/bizinfo/communicationView.jsp?apv_perm_no=2004300010130201968&amp;area1=&amp;area2=&amp;currpage=1&amp;searchKey=01&amp;searchVal=텐바이텐&amp;stdate=&amp;enddate=" target="_blank">[사업자 정보확인]</a></p>
		<p>통신판매업신고 : 제 01-1968호 <span class="bar">ㅣ</span> 호스팅서비스 : (주)텐바이텐</p>
	</div>
	<p id="btnGotop" class="btnGotop"><span>TOP</span></p>
</footer>
