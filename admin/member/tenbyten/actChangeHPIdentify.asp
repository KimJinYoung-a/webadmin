<%@ language=vbscript %>
<%
	Option Explicit
	Response.Expires = -1440
	
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  본인확인 결과값 처리
' History : 2011.05.31 허진원 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/common/member/oneself/nice.nuguya.oivs.asp"-->
<%
	dim KeyString, strRecvData, ssResidNo, ssOrderNo
	dim sConnectIP,sUserName,sUserPW
	Dim clsConnDB, clsUILog,clsSearchPW	
	dim iUserSeq, sEmail,dRegdate
		
	 sConnectIP = Left(request.ServerVariables("REMOTE_ADDR"),32)
	 sUserName 	=  requestCheckVar(Request.Form("sHUN"),20)	

	'=================================================================================================
	'=====	▣ 계약시에 발급 받은 데이터 설정 : 계약시에 발급된 KeyString값을 설정하십시오. ▣
	'=================================================================================================
	KeyString = "r6cA3YS9s8WTktrzgfNSOqQXsKf6GnNNpVEnn4DeDuwzgXhICcDpFhTefoTvFUbsux9EvPsbadplISwb"  '//키스트링(80자리)를 넣어주세요.
	
	oivsObject.AthKeyStr = KeyString
	
	strRecvData = Request.Form( "SendInfo" )
	oivsObject.resolveDatas(strRecvData)
	
	'// 해킹방지를 위해 세션에 저장된 값과 비교 .. 
	ssResidNo = session("niceRsdNo")
	ssOrderNo = session("niceOrderNo")
	
	If  ssResidNo <> oivsObject.residNo or ssOrderNo <> oivsObject.ordNo then
		response.write("세션정보가 존재하지 않습니다.")
	End if
	
'	response.write("<BR>주문번호 : " + oivsObject.ordNo)
'	response.write("<BR>본인인증 성공여부 : " + oivsObject.retCd + "(1:성공 / 0:실패)")
'	response.write("<BR>인증코드 : " + oivsObject.resCd)
'	response.write("<BR>응답 메시지 : " + oivsObject.message)
'	response.write("<BR>회원사 ID : " + oivsObject.niceId)
'	response.write("<BR>주민번호 : " + oivsObject.residNo)
'	response.write("<BR>휴대폰번호 : " + oivsObject.phoneNo)
	
	
	IF  oivsObject.retCd = "1" THEN
		'// 본인확인 성공
	%>
		<script language="javascript">
		opener.parent.actChgHP();
		self.close();
		</script>
	<%
	ELSE
	    Call Alert_close("본인인증에 실패하였습니다.\n입력한 정보를 확인해주세요.")
	    response.end
	END IF	
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->