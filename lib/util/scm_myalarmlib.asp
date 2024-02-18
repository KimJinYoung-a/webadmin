<%

'// 웹, 앱, 모바일에 공통적으로 적용하려면 /imgstatic/lib/badgelib.asp 참조
'// 어드민에 파일 생성(로그 등록 + 쿠키X = 재로그인시 알림 받음)

'' /lib/util/myalarmlib.asp
'' /lib/util/scm_myalarmlib.asp

''  msgdiv 	구분					입력 시점
'' ===========================================================================
''  000		단체알림
'' 	001		신규가입쿠폰			회원가입시
'' 	002		쿠폰만료				일 1회(새벽)
'' 	003		장바구니 상품 이벤트	MyAlarm_CheckNewMyAlarm 실행시
'' 	004		위시 상품 이벤트		MyAlarm_CheckNewMyAlarm 실행시
'' 	005		1:1 상담				일 1회(새벽,어제날짜로)
'' 	006		상품 QnA				일 1회(새벽,어제날짜로)
'' 	007		이벤트 당첨				일 1회
''  901		관심상품 없음
''  902		관련이벤트 없음

Function MyAlarm_InsertMyAlarm_SCM(userid, msgdiv, title, subtitle, contents, wwwTargetURL)
	dim strSql, i

	'// 중복입력 안함(1:!상담, 상품QNA는 허용)
	strSql = " [db_my10x10].[dbo].[usp_Ten_MyAlarm_ProcInsertLOG] ('" + CStr(userid) + "', '" + CStr(msgdiv) + "', '" + CStr(html2db(title)) + "', '" + CStr(html2db(subtitle)) + "', '" + CStr(html2db(contents)) + "', '" + CStr(wwwTargetURL) + "') "
	dbget.Execute strSql
End Function

%>
