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
    title = "[10x10]���� ���� �ȳ��Դϴ�."
    lmscontents = "�ȳ��ϼ���. �ٹ������Դϴ�." & vbcrlf
    lmscontents = lmscontents & "������ ȯ���մϴ�."  & vbcrlf& vbcrlf
    lmscontents = lmscontents & "�Ʒ� ��ũ�� �̵� �Ͻ� �� ��ü������ �Է��� �ֽø�,"  & vbcrlf
    lmscontents = lmscontents & "����(SCM) �α����� �����մϴ�."  & vbcrlf & vbcrlf
    lmscontents = lmscontents & "[����ڵ�������� �̵�]"  & vbcrlf
    lmscontents = lmscontents & "https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs="+qstring & vbcrlf& vbcrlf
    lmscontents = lmscontents & "�����մϴ�."

    kakaocontents = "�ٹ����� ������ ȯ���մϴ�." & vbcrlf & vbcrlf
    kakaocontents = kakaocontents & "�Ʒ� ��ũ�� �̵� �Ͻ� �� ��ü ������ �Է��� �ֽø�, "
    kakaocontents = kakaocontents & "����(SCM) �α����� �����մϴ�." & vbcrlf & vbcrlf
    kakaocontents = kakaocontents & "�����մϴ�." & vbcrlf
    btnJson = "{""button"":[{""name"":""��ü���� �Է� �ٷΰ���"",""type"":""WL"", ""url_mobile"":""https://scm.10x10.co.kr/common/partner/companyinfo.asp?qs=" & trim(qstring) &"""}]}"
    'LMS�߼�
    'call SendNormalLMS(hp,title,"",lmscontents)
    'On Error Resume Next
    'KAKAO�߼�
    call SendKakaoMsg_LINK(hp,"","A-0006",kakaocontents,"LMS",title,lmscontents,btnJson)
    'Email�߼�
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
    mailtitle = "[10x10] ���� ���� �ȳ� �����Դϴ�."

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