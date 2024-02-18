<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 수기 등록 품절 이메일 발송
' Hieditor : 2013.06.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
Dim iCurrentpage, oxsite, i, iTotCnt, vBody	, xsite, webImgUrl
	xsite = "interparkPTM"
	webImgUrl = "http://webimage.10x10.co.kr"	
%>
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%

	Set oxsite = new CxSiteTempOrder
	oxsite.FCurrPage = 1
	oxsite.FPageSize = 100
	oxsite.frectmallid = xsite
	oxsite.frectitemid = ""
	oxsite.FrectItemName = ""
	oxsite.frectoutmallSellYn = "Y"
	oxsite.FrectSoldOUT = "Y"
	oxsite.getxsitesoldout_scheduler
	
	iTotCnt = oxsite.ftotalcount
	
	If iTotCnt <> 0 Then
		vBody = vBody & "<table cellpadding=""3"" cellspacing=""1"" border=""1"">" & vbCrLf
		vBody = vBody & "<tr bgcolor=""#E6E6E6"">" & vbCrLf
		vBody = vBody & "	<td align=""center"">번호</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">제휴몰</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">상품</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">상품코드</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">상품명</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">제휴몰판매가</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">10x10<br>품절여부</td>" & vbCrLf
		vBody = vBody & "	<td align=""center"">사용여부</td>" & vbCrLf
		vBody = vBody & "</tr>" & vbCrLf
			
		For i = 0 To oxsite.FResultCount -1

			vBody = vBody & "<tr bgcolor=""FFFFFF"">" & vbCrLf
			vBody = vBody & "	<td width=""70"" align=""center"">" & i+1 & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""70"" align=""center"">" & oxsite.FItemList(i).fmallID & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center""><a href=""http://www.10x10.co.kr/" & oxsite.FItemList(i).fitemid & """ target=""_blank""><img src=""" & oxsite.FItemList(i).fsmallimage & """ border=""0""></a></td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & oxsite.FItemList(i).fitemid & "</td>" & vbCrLf
			vBody = vBody & "	<td>" & oxsite.FItemList(i).fitemname & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & FormatNumber(oxsite.FItemList(i).foutmallPrice,0) & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & CHKIIF(oxsite.FItemList(i).fsoldout="True","<b><font color=red>품절</font></b>","판매중") & "</td>" & vbCrLf
			vBody = vBody & "	<td width=""60"" align=""center"">" & oxsite.FItemList(i).foutmallSellYn & "</td>" & vbCrLf
			vBody = vBody & "</tr>" & vbCrLf
		
		Next

		vBody = vBody & "</table>" & vbCrLf
		
		response.write vBody
		
		if oxsite.FResultCount > 0 then
			Call SendMail("admin@10x10.co.kr", "kjy8517@10x10.co.kr", "제휴몰 품절상품", vBody)
		end if
		
	End If
	set oxsite = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->