<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/util/base64unicode.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<%
	dim brandid, qstring, email, hp
	brandid = requestCheckVar(Request("brandid"),32)
    hp = requestCheckVar(Request("hp"),32)
    email = requestCheckVar(Request("email"),128)
    qstring = Request("qs")

    dim title, lmscontents, kakaocontents, btnJson
    title = "[10x10]입점 관련 안내입니다."
    lmscontents = "안녕하세요. 텐바이텐입니다." & vbcrlf
    lmscontents = lmscontents & "입점을 환영합니다."  & vbcrlf& vbcrlf
    lmscontents = lmscontents & "아래 링크로 이동 하신 후 업체정보를 입력해 주시면,"  & vbcrlf
    lmscontents = lmscontents & "어드민(SCM) 로그인이 가능합니다."  & vbcrlf & vbcrlf
    lmscontents = lmscontents & "[사업자등록페이지 이동]"  & vbcrlf
    lmscontents = lmscontents & "https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs="+qstring & vbcrlf& vbcrlf
    lmscontents = lmscontents & "감사합니다."

    kakaocontents = "텐바이텐 입점을 환영합니다." & vbcrlf & vbcrlf
    kakaocontents = kakaocontents & "아래 링크로 이동 하신 후 업체 정보를 입력해 주시면, "
    kakaocontents = kakaocontents & "어드민(SCM) 로그인이 가능합니다." & vbcrlf & vbcrlf
    kakaocontents = kakaocontents & "감사합니다." & vbcrlf
    btnJson = "{""button"":[{""name"":""업체정보 입력 바로가기"",""type"":""WL"", ""url_mobile"":""https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs=" & trim(qstring) &"""}]}"
    'LMS발송
    'call SendNormalLMS(hp,title,"",lmscontents)
    'On Error Resume Next
    'KAKAO발송
    call SendKakaoMsg_LINK(hp,"","A-0006",kakaocontents,"LMS",title,lmscontents,btnJson)
    'Email발송
	call sendmailPartnerJoin(email,"https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs="+qstring)

    If Err.Number = 0 Then
        response.write "OK"
        response.end
    end if
    On Error Goto 0

function sendmailPartnerJoin(mailto, contents)
    dim mailfrom, mailtitle, mailcontent,dirPath,fileName
    dim fs,objFile

    mailfrom = "customer@10x10.co.kr"
    mailtitle = "[10x10] 입점 관련 안내 메일입니다."

    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    dirPath = server.mappath("/lib/email/mailtemplate")
    fileName = dirPath&"\\mail_partner_join.htm"
    Set objFile = fs.OpenTextFile(fileName,1)
    mailcontent = objFile.readall
    mailcontent = replace(mailcontent,":CONTENTSHTML:",contents)
    call sendmail(mailfrom, mailto, mailtitle, mailcontent)
end Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->