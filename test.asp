<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->

<%
dim fullText, failText, btnJson, requiremakerid, itemName, orderserial, refundstr, buyhp, failtitle, ipgodate
buyhp="010-9177-8708"
requiremakerid="temp"
orderserial="21100198279"
itemName="2014����ͽ�ƼĿ(3��)"
refundstr="1,000��"
ipgodate="2021-10-18"
failtitle = "[�ٹ�����]��ǰ��� �ȳ�"
fullText = "[10x10] ��ǰ��� �ȳ�" & vbCrLf & vbCrLf
fullText = fullText & "ǰ����� �ȳ���ȴ� ��ǰ�� ��� Ȯ���Ǿ� �߼� ��������, �Ʒ��� �����ϱ��� ����� �� �ֵ��� �ּ��� ����� ���ϰڽ��ϴ�." & vbCrLf & vbCrLf & vbCrLf
fullText = fullText & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
fullText = fullText & "�� ��ǰ�� : "& Itemname &"" & vbCrLf
fullText = fullText & "�� ��߿����� : "& ipgodate &"" & vbCrLf & vbCrLf
fullText = fullText & "�����մϴ�."
failText = fullText
btnJson = "{""button"":[{""name"":""�ֹ����� �ٷΰ���"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/L1izHiDBdjb""}]}"

'call SendKakaoCSMsg_LINK("", "01091778708","1644-6030","KC-0024",fullText,"LMS",failtitle,failText,btnJson,"","")
'call SendKakaoMsg_LINK(buyhp,"1644-6030","A-001","����","LMS","���н� ����","���н� ����","")
'call SendNormalSMS_LINK(buyhp,"1644-6030","����2")
%>