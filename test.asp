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
itemName="2014쿨워터스티커(3종)"
refundstr="1,000원"
ipgodate="2021-10-18"
failtitle = "[텐바이텐]상품출고 안내"
fullText = "[10x10] 상품출고 안내" & vbCrLf & vbCrLf
fullText = fullText & "품절취소 안내드렸던 상품의 재고가 확보되어 발송 예정으로, 아래의 예정일까지 출발할 수 있도록 최선의 노력을 다하겠습니다." & vbCrLf & vbCrLf & vbCrLf
fullText = fullText & "■ 주문번호 : "& orderserial &"" & vbCrLf
fullText = fullText & "■ 상품명 : "& Itemname &"" & vbCrLf
fullText = fullText & "■ 출발예정일 : "& ipgodate &"" & vbCrLf & vbCrLf
fullText = fullText & "감사합니다."
failText = fullText
btnJson = "{""button"":[{""name"":""주문내역 바로가기"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/L1izHiDBdjb""}]}"

'call SendKakaoCSMsg_LINK("", "01091778708","1644-6030","KC-0024",fullText,"LMS",failtitle,failText,btnJson,"","")
'call SendKakaoMsg_LINK(buyhp,"1644-6030","A-001","내용","LMS","실패시 제목","실패시 내용","")
'call SendNormalSMS_LINK(buyhp,"1644-6030","내용2")
%>