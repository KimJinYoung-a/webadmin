<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Session.codepage="65001"
Response.codepage="65001"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")

if InStr(refer,"webadmin.10x10.co.kr")<1 then
	Call Alert_Return("잘못된 접속입니다.")
	response.end
end if

dim sqlStr,DesignerID
dim isellvat, ibuyvat, imargin
dim cd1,cd2,cd3, waititemid, makerid
dim cd2slice,cd3slice,vAddImgNameTotCnt

waititemid = Request.Form("waititemid")
cd1 = Request.Form("cd1")
cd2 = Request.Form("cd2")
cd3 = Request.Form("cd3")
imgmain = Request.Form("imgmain")
imgbasic = Request.Form("imgbasic")
imgadd1 = Request.Form("imgadd1")
imgadd2 = Request.Form("imgadd2")
imgadd3 = Request.Form("imgadd3")
DesignerID = Request.Form("designerid")
vAddImgNameTotCnt = Request.Form("addimgname").Count
makerid = request.cookies("partner")("userid")

If (WaitItemCheckMyItemYN(makerid,waititemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If
Dim vAddImgMobileNameReal

'// 한정판매 수량 계산 (옵션 포함) //
dim limitno
limitno = Request.Form("limitno")
if (limitno="") or (Not IsNumeric(limitno)) then limitno=0
    
'#########// 상품 옵션 넣기 //##############################
'옵션이 있을경우 옵션수에 따른 수량 계산
dim iErrMsg
dim itemoptionlist, itemoptionText, itemoptioncount, itemoptionPrice, itemoptionBuyprice, itemoptionLimitNo, itemoptionLimitYn
dim sellcash,buycash
sellcash=Request.Form("sellcash")
buycash=Request.Form("buycash")
If sellcash="" Then sellcash=0
If buycash="" Then buycash=0

Dim requireimgchk, requireMakeEmail
requireimgchk = Request.Form("requireimgchk")
requireMakeEmail = Request.Form("requireMakeEmail")
If requireimgchk="" Then requireimgchk="N"
If requireimgchk="N" Then requireMakeEmail=""
'###########################################################################
'상품 데이터 입력
'###########################################################################
sqlStr = "update db_academy.dbo.tbl_diy_wait_item" + vbCrlf
sqlStr = sqlStr & " set cate_large='" & cd1 & "'" + vbCrlf
sqlStr = sqlStr & " ,cate_mid='" & cd2 & "'" + vbCrlf
sqlStr = sqlStr & " ,cate_small='" & cd3 & "'" + vbCrlf
sqlStr = sqlStr & " ,itemdiv='" & Cstr(Request.Form("itemdiv")) & "'" + vbCrlf
sqlStr = sqlStr & " ,itemname=convert(varchar(64),'" & chrbyte(html2db(Request.Form("itemname")),64,"") & "')" + vbCrlf
sqlStr = sqlStr & " ,itemsource='" & chrbyte(html2db(Request.Form("itemsource")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & " ,itemsize='" & chrbyte(html2db(Request.Form("itemsize")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & " ,itemWeight='" & chrbyte(html2db(Request.Form("itemWeight")),12,"") & "'" + vbCrlf
sqlStr = sqlStr & " ,buycash=" & buycash & "" + vbCrlf
sqlStr = sqlStr & " ,sellcash=" & sellcash & "" + vbCrlf
sqlStr = sqlStr & " ,mileage=" & CLng(CLng(sellcash)*0.01) & "" + vbCrlf
sqlStr = sqlStr & " ,deliverytype='" & Request.Form("deliverytype") & "'" + vbCrlf
sqlStr = sqlStr & " ,sourcearea='" & chrbyte(html2db(Request.Form("sourcearea")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & " ,makername='" & chrbyte(html2db(Request.Form("makername")),64,"") & "'" + vbCrlf
sqlStr = sqlStr & " ,limityn='" & Request.Form("limityn") & "'" + vbCrlf
sqlStr = sqlStr & " ,limitno="  & limitno & "" + vbCrlf
sqlStr = sqlStr & " ,currstate='"  & Request.Form("currstate") & "'" + vbCrlf
sqlStr = sqlStr & " ,reRegDate=getdate()" + vbCrlf
sqlStr = sqlStr & " ,keywords=convert(varchar(128),'" & chrbyte(html2db(Request.Form("keywords")),128,"") & "')" + vbCrlf
sqlStr = sqlStr & " ,mwdiv='" & Request.Form("mwdiv") & "'" + vbCrlf
sqlStr = sqlStr & " ,vatYn='" & Request.Form("vatYn") & "'" + vbCrlf
sqlStr = sqlStr & " ,ordercomment='" & html2db(Request.Form("ordercomment")) & "'" + vbCrlf
sqlStr = sqlStr & " ,refundpolicy='" & html2db(Request.Form("refundpolicy")) & "'" + vbCrlf
sqlStr = sqlStr & " ,cstodr='" & html2db(Request.Form("cstodr")) & "'" & vbCrlf
sqlStr = sqlStr & " ,requireMakeDay='" & html2db(Request.Form("requireMakeDay")) & "'" & vbCrlf
sqlStr = sqlStr & " ,requirecontents='" & html2db(Request.Form("requirecontents")) & "'" & vbCrlf
sqlStr = sqlStr & " ,requireMakeEmail='" & html2db(requireMakeEmail) & "'" & vbCrlf
sqlStr = sqlStr & " ,requireimgchk='" + requireimgchk + "'" + vbCrlf
sqlStr = sqlStr & " ,infoDiv='" & Request.Form("infoDiv") & "'" & vbCrlf
sqlStr = sqlStr & " ,safetyYn='" & Request.Form("safetyYn") & "'" & vbCrlf
sqlStr = sqlStr & " ,safetyDiv='" & Request.Form("safetyDiv") & "'" & vbCrlf
sqlStr = sqlStr & " ,safetyNum='" & chrbyte(html2db(Request.Form("safetyNum")),24,"") & "'" & vbCrlf
sqlStr = sqlStr & " where itemid=" + CStr(waititemid) + vbCrlf
'Response.write sqlStr
'Response.end
dbACADEMYget.Execute sqlStr

dim addimageComma,imgmain,imgbasic,imgadd1,imgadd2,imgadd3
'' 임시 등록시에는 아이콘 만들지 않음. 
imgmain = Request.Form("imgmain")
imgbasic = Request.Form("imgbasic")
imgadd1 = Request.Form("imgadd1")
imgadd2 = Request.Form("imgadd2")
imgadd3 = Request.Form("imgadd3")
addimageComma = imgadd1 & "," & imgadd2 & "," & imgadd3 & ",,"

'###########################################################################
'이미지 데이터 넣기
'###########################################################################
sqlStr = "update db_academy.dbo.tbl_diy_wait_item set "
sqlStr = sqlStr & "  mainimage='" & imgmain & "'"
sqlStr = sqlStr & ", basicimage='" & imgbasic & "'"
sqlStr = sqlStr & ", imgadd='" & addimageComma & "'"
sqlStr = sqlStr & " where itemid=" & waititemid
rsACADEMYget.Open sqlStr,dbACADEMYget,1

Dim oldimageMobileName(15), addgubun, vMobileRealNum
sqlStr = " select gubun, addimage"
sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item_addimage"
sqlStr = sqlStr & " where itemid = '" & waititemid & "' "
sqlStr = sqlStr & " and IMGTYPE=2"
sqlStr = sqlStr & " order by gubun"
rsACADEMYget.Open sqlStr, dbACADEMYget, 1 
if Not rsACADEMYget.Eof then
	do until rsACADEMYget.Eof
		addgubun = rsACADEMYget("gubun")
		oldimageMobileName(addgubun-1) = rsACADEMYget("addimage")
		rsACADEMYget.MoveNext
	loop
end if
rsACADEMYget.close

Dim vTemptext , vText, vAddimg, vTempAddimg
Dim vCnt : vCnt = Request.Form("addimgtext").count
If vCnt > 1 Then
	For k=1 To vCnt
		vText = vText + Request.Form("addimgtext")(k)
		vAddimg = vAddimg + Request.Form("addimgname")(k)
		If k < vCnt Then
			vText = vText & "|"
			vAddimg = vAddimg & "|"
		End If 
	Next
Else
	vText = Request.Form("addimgtext")
	vAddimg = Request.Form("addimgname")
End If 

vMobileRealNum = 0
If Request.Form("addimgtext")<>"" Or Request.Form("addimgname")<>"" Then

	For k = 0 To vAddImgNameTotCnt-1
			If vAddImgNameTotCnt>1 Then
				vTemptext = Split(vText,"|")(k)
				vTempAddimg = Split(vAddimg,"|")(k)
			Else
				vTemptext=vText
				vTempAddimg = vAddimg
			End If

			If(vTemptext<>"" Or vTempAddimg<>"") Then
			vMobileRealNum = vMobileRealNum + 1

			sqlStr = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_wait_item_addimage WHERE ITEMID='" & waititemid & "' and IMGTYPE=2 and GUBUN="&vMobileRealNum&") "			
			sqlStr = sqlStr + "	BEGIN "
			sqlStr = sqlStr+ " 		INSERT INTO db_academy.dbo.tbl_diy_wait_item_addimage (ITEMID,IMGTYPE,GUBUN,ADDIMAGE,addimgtext)"
			sqlStr = sqlStr + "     	VALUES ('" & waititemid & "',2,"&vMobileRealNum&",'" & vTempAddimg & "' ,'"& html2db(vTemptext) &"' ) "
			sqlStr = sqlStr + " 	END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " 	BEGIN "			
			sqlStr = sqlStr + "		UPDATE db_academy.dbo.tbl_diy_wait_item_addimage "
			sqlStr = sqlStr + " 		SET ADDIMAGE ='" & vTempAddimg & "'"
			sqlStr = sqlStr + " 		, addimgtext ='" & html2db(vTemptext) & "'"
			sqlStr = sqlStr + " 		WHERE ITEMID = '" & waititemid & "' "
			sqlStr = sqlStr + " 		and IMGTYPE=2"
			sqlStr = sqlStr + " 		and GUBUN ="&vMobileRealNum&""
			sqlStr = sqlStr + " 	END "

			dbACADEMYget.execute sqlStr
			End If

	Next
End If

'###########################################################################
'아이템 동영상
'###########################################################################
'// 2016.2.16 신규추가 상품상세설명 동영상 추가 - 원승현
'// 2016.6.24 핑거스 diy샵 추가  - 이종화
'// 아이템 동영상 값 정규식으로 src, width, height값 뽑아냄
If Trim(Request.Form("itemvideo")) <> "" Then
	Dim itemvideo, RetSrc, RetWidth, RetHeight, regEx, Matches, Match, VideoTempSrc, VideoTempWidth, VideoTempHeight, videoType, dbsql
	itemvideo = Request.Form("itemvideo")
	itemvideo = Trim(Replace(itemvideo,"""","'"))
	itemvideo = replace(itemvideo,"BUFiframe","iframe")
	
	'// iframe 이외의 코드는 잘라버림
	itemvideo = Left(itemvideo, InStrRev(itemvideo, "</iframe>")+9)

	'// 비디오 타입지정(유투브인지 비메오인지)
	If InStr(itemvideo, "youtube")>0 Then
		videoType = "youtube"
	ElseIf InStr(itemvideo, "vimeo")>0 Then
		videoType = "vimeo"
	Else
		videoType = "etc"
	End If

	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True

	regEx.pattern = "<iframe [^<>]*>"
	Set Matches = regEx.execute(itemvideo)
	For Each Match In Matches
		VideoTempSrc =  Mid(Match.Value, InStrRev(Match.Value,"src='")+5)
		RetSrc = Left(VideoTempSrc, InStr(VideoTempSrc, "'")-1)
		
		If InStrRev(Match.Value,"width='") > 0 then
		VideoTempWidth =  Mid(Match.Value, InStrRev(Match.Value,"width='")+7)
		RetWidth = Left(VideoTempWidth, InStr(VideoTempWidth, "'")-1)
		End If 
		
		If InStrRev(Match.Value,"height='") > 0 then
		VideoTempHeight =  Mid(Match.Value, InStrRev(Match.Value,"height='")+8)
		RetHeight = Left(VideoTempHeight, InStr(VideoTempHeight, "'")-1)
		End If 
	Next
	Set regEx = Nothing
	Set Matches = Nothing

	If Not(videoType="etc") Then
		dbsql = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_wait_item_videos WHERE ITEMID='" & waititemid & "')" + vbCrlf
		dbsql = dbsql + "	BEGIN " + vbCrlf
		dbsql = dbsql+ " 		INSERT INTO db_academy.dbo.tbl_diy_wait_item_videos (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate)" + vbCrlf
		dbsql = dbsql + "     	VALUES ('"&CStr(waititemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
		dbsql = dbsql + " 	END " + vbCrlf
		dbsql = dbsql + " ELSE " + vbCrlf
		dbsql = dbsql + " 	BEGIN " + vbCrlf
		dbsql = dbsql + "		UPDATE db_academy.dbo.tbl_diy_wait_item_videos " + vbCrlf
		dbsql = dbsql + " 		SET videotype ='" & videoType & "'" + vbCrlf
		dbsql = dbsql + " 		, videourl ='" & RetSrc & "'" + vbCrlf
		dbsql = dbsql + " 		, videowidth ='" & RetWidth & "'" + vbCrlf
		dbsql = dbsql + " 		, videoheight ='" & RetHeight & "'" + vbCrlf
		dbsql = dbsql + " 		, videofullurl ='" & chrbyte(html2db(itemvideo),255,"") & "'" + vbCrlf
		dbsql = dbsql + " 		WHERE ITEMID = '" & waititemid & "' " + vbCrlf
		dbsql = dbsql + " 	END "
'		Response.write dbsql & "<br>"
'		Response.end
		dbACADEMYget.execute(dbsql)
	End If
'		Response.end
End If
%>
<script>
<!--
	parent.fntempSaveEnd('<%=waititemid%>');
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->