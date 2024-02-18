<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
dim i,k

dim refer
''// 유효 접근 주소 검사 //
refer = request.ServerVariables("HTTP_REFERER")

if InStr(refer,"webadmin.10x10.co.kr")<1 then
	Call Alert_Return("잘못된 접속입니다.")
	response.end
end if

dim sqlStr,DesignerID, makerid
dim isellvat, ibuyvat, imargin
dim cd1,cd2,cd3, itemid
dim cd2slice,cd3slice,vAddImgNameTotCnt
itemid = requestCheckvar(request("itemid"),10)
makerid = requestCheckvar(request("designerid"),32)

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If
vAddImgNameTotCnt = Request.Form("addimgname").Count

Dim vAddImgMobileNameReal

'// 한정판매 수량 계산 (옵션 포함) //
dim limitno
limitno = Request.Form("limitno")
if (limitno="") or (Not IsNumeric(limitno)) then limitno=0
    
'#########// 상품 옵션 넣기 //##############################
'옵션이 있을경우 옵션수에 따른 수량 계산
dim iErrMsg
dim totalLimitno,sellcash,buycash
sellcash=Request.Form("sellcash")
buycash=Request.Form("buycash")
If sellcash="" Then sellcash=0
If buycash="" Then buycash=0

Dim keywords, cstodr, requireMakeDay, requirecontents, refundpolicy
Dim infoDiv, safetyYn, safetyDiv, safetyNum, requirechk, requireemail, imgadd1, imgadd2
Dim itemdiv, deliverytype, sourcearea, makername, itemsource, itemsize, itemWeight, imgbasic


keywords		= html2db(requestCheckvar(request("keywords"),512))
cstodr			= Request("cstodr")
requireMakeDay	= Request("requireMakeDay")
requirecontents	= html2db(Request("requirecontents"))
refundpolicy	= html2db(Request("refundpolicy"))
infoDiv			= Request("infoDiv")
safetyYn		= Request("safetyYn")
safetyDiv		= Request("safetyDiv")

itemdiv			= Request("itemdiv")
deliverytype	= Request("deliverytype")
sourcearea		= requestCheckvar(Request("sourcearea"),64)
makername		= requestCheckvar(Request("makername"),64)
itemsource		= requestCheckvar(Request("itemsource"),128)
itemsize		= requestCheckvar(Request("itemsize"),32)
itemWeight		= Request("itemWeight")

safetyNum		= chrbyte(html2db(Request("safetyNum")),24,"")
requirechk		= html2db(Request("requireimgchk"))
If requirechk="" Then requirechk="N"
requireemail	= html2db(requestCheckvar(Request("requireMakeEmail"),128))
imgbasic = Request("imgbasic")
imgadd1 = Request("imgadd1")
imgadd2 = Request("imgadd2")

'###########################################################################
'상품 데이터 입력
'###########################################################################
sqlStr = "update db_academy.dbo.tbl_diy_item" + vbCrlf
sqlStr = sqlStr & " set limityn='" & Request.Form("limityn") & "'" + vbCrlf
sqlStr = sqlStr & " ,limitno="  & limitno & "" + vbCrlf
sqlStr = sqlStr & " ,lastupdate=getdate()" + vbCrlf
sqlStr = sqlStr + " , requireimgchk = '" + html2db(requirechk) + "'" + VbCrlf
sqlStr = sqlStr + " , itemname = '" + html2db(Request("itemname")) + "'" + VbCrlf
If imgbasic <> "" Then
sqlStr = sqlStr + " , basicimage = '" + html2db(imgbasic) + "'" + VbCrlf
End If
sqlStr = sqlStr + " , deliverytype = '" + html2db(deliverytype) + "'" + VbCrlf
sqlStr = sqlStr + " , itemdiv = '" + html2db(itemdiv) + "'" + VbCrlf
sqlStr = sqlStr & " where itemid=" + CStr(itemid) + vbCrlf
dbACADEMYget.Execute sqlStr

'// 상품상세 설명 유무에 따른 처리
sqlStr = "select count(*) from db_academy.dbo.tbl_diy_item_Contents where itemid=" & itemid
rsACADEMYget.Open sqlStr ,dbACADEMYget,1
if rsACADEMYget(0)<1 then
	sqlStr = "insert into db_academy.dbo.tbl_diy_item_Contents (itemid, keywords, cstodr , requireMakeDay , requirecontents,infoDiv,safetyYn,safetyDiv,safetyNum, requiremakeemail) values " + VbCrlf
	sqlStr = sqlStr + " (" + CStr(itemid) + VbCrlf
	sqlStr = sqlStr + " , '" + keywords + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + cstodr + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + requireMakeDay + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + requirecontents + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + infoDiv + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + safetyYn + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + safetyDiv + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + safetyNum + "'" + VbCrlf
	sqlStr = sqlStr + " , '" + requireemail + "')" + VbCrlf
	dbACADEMYget.Execute sqlStr
Else
	sqlStr = "update db_academy.dbo.tbl_diy_item_Contents" + VbCrlf
	sqlStr = sqlStr + " set cstodr='" & cstodr & "'" + vbCrlf
	sqlStr = sqlStr & " ,sourcearea='" & sourcearea & "'" + vbCrlf
	sqlStr = sqlStr & " ,makername='" & makername & "'" + vbCrlf
	sqlStr = sqlStr & " ,itemsource='" & itemsource & "'" + vbCrlf
	sqlStr = sqlStr & " ,itemsize='" & itemsize & "'" + vbCrlf
	sqlStr = sqlStr & " ,itemWeight='" & itemWeight & "'" + vbCrlf
	sqlStr = sqlStr & " ,requireMakeDay='" & requireMakeDay & "'" + vbCrlf
	sqlStr = sqlStr & " ,infoDiv='" & infoDiv & "'" + vbCrlf
	sqlStr = sqlStr & " ,requiremakeemail='" & requireemail & "'" + vbCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "
	dbACADEMYget.Execute sqlStr
end if
rsACADEMYget.Close

'###########################################################################
'이미지 데이터 넣기
'###########################################################################
if (imgadd1<>"") then
	sqlStr = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_item_addimage WHERE ITEMID=" & itemid & " and IMGTYPE=0 and GUBUN=1) "			
	sqlStr = sqlStr + "	BEGIN "
	sqlStr = sqlStr+ " 		INSERT INTO db_academy.dbo.tbl_diy_item_addimage (ITEMID,IMGTYPE,GUBUN,ADDIMAGE)"
	sqlStr = sqlStr + "     	VALUES (" & itemid & ",0,1,'" & imgadd1 & "') "
	sqlStr = sqlStr + " 	END "
	sqlStr = sqlStr + " ELSE "
	sqlStr = sqlStr + " 	BEGIN "			
	sqlStr = sqlStr + "		UPDATE db_academy.dbo.tbl_diy_item_addimage "
	sqlStr = sqlStr + " 		SET ADDIMAGE ='" & imgadd1 & "'"
	sqlStr = sqlStr + " 		WHERE ITEMID =" & itemid 
	sqlStr = sqlStr + " 		and IMGTYPE=0"
	sqlStr = sqlStr + " 		and GUBUN =1"
	sqlStr = sqlStr + " 	END "		
	
	dbACADEMYget.execute sqlStr
end if

if (imgadd2<>"") then
	sqlStr = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_item_addimage WHERE ITEMID=" & itemid & " and IMGTYPE=0 and GUBUN=2) "			
	sqlStr = sqlStr + "	BEGIN "
	sqlStr = sqlStr+ " 		INSERT INTO db_academy.dbo.tbl_diy_item_addimage (ITEMID,IMGTYPE,GUBUN,ADDIMAGE)"
	sqlStr = sqlStr + "     	VALUES (" & itemid & ",0,2,'" & imgadd2 & "') "
	sqlStr = sqlStr + " 	END "
	sqlStr = sqlStr + " ELSE "
	sqlStr = sqlStr + " 	BEGIN "			
	sqlStr = sqlStr + "		UPDATE db_academy.dbo.tbl_diy_item_addimage "
	sqlStr = sqlStr + " 		SET ADDIMAGE ='" & imgadd2 & "'"
	sqlStr = sqlStr + " 		WHERE ITEMID =" & itemid 
	sqlStr = sqlStr + " 		and IMGTYPE=0"
	sqlStr = sqlStr + " 		and GUBUN =2"
	sqlStr = sqlStr + " 	END "		
	
	dbACADEMYget.execute sqlStr
end if

Dim oldimageMobileName(15), addgubun, vMobileRealNum
sqlStr = " select gubun, addimage"
sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item_addimage"
sqlStr = sqlStr & " where itemid = '" & itemid & "' "
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

			sqlStr = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_item_addimage WHERE ITEMID='" & itemid & "' and IMGTYPE=2 and GUBUN="&vMobileRealNum&") "			
			sqlStr = sqlStr + "	BEGIN "
			sqlStr = sqlStr+ " 		INSERT INTO db_academy.dbo.tbl_diy_item_addimage (ITEMID,IMGTYPE,GUBUN,ADDIMAGE,addimgtext)"
			sqlStr = sqlStr + "     	VALUES ('" & itemid & "',2,"&vMobileRealNum&",'" & vTempAddimg & "' ,'"& html2db(vTemptext) &"' ) "
			sqlStr = sqlStr + " 	END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " 	BEGIN "			
			sqlStr = sqlStr + "		UPDATE db_academy.dbo.tbl_diy_item_addimage "
			sqlStr = sqlStr + " 		SET ADDIMAGE ='" & vTempAddimg & "'"
			sqlStr = sqlStr + " 		, addimgtext ='" & html2db(vTemptext) & "'"
			sqlStr = sqlStr + " 		WHERE ITEMID = '" & itemid & "' "
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
		dbsql = " IF Not Exists(SELECT IDX FROM db_academy.dbo.tbl_diy_item_videos WHERE ITEMID='" & itemid & "')" + vbCrlf
		dbsql = dbsql + "	BEGIN " + vbCrlf
		dbsql = dbsql+ " 		INSERT INTO db_academy.dbo.tbl_diy_item_videos (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate)" + vbCrlf
		dbsql = dbsql + "     	VALUES ('"&CStr(itemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
		dbsql = dbsql + " 	END " + vbCrlf
		dbsql = dbsql + " ELSE " + vbCrlf
		dbsql = dbsql + " 	BEGIN " + vbCrlf
		dbsql = dbsql + "		UPDATE db_academy.dbo.tbl_diy_item_videos " + vbCrlf
		dbsql = dbsql + " 		SET videotype ='" & videoType & "'" + vbCrlf
		dbsql = dbsql + " 		, videourl ='" & RetSrc & "'" + vbCrlf
		dbsql = dbsql + " 		, videowidth ='" & RetWidth & "'" + vbCrlf
		dbsql = dbsql + " 		, videoheight ='" & RetHeight & "'" + vbCrlf
		dbsql = dbsql + " 		, videofullurl ='" & chrbyte(html2db(itemvideo),255,"") & "'" + vbCrlf
		dbsql = dbsql + " 		WHERE ITEMID = '" & itemid & "' " + vbCrlf
		dbsql = dbsql + " 	END "
		dbACADEMYget.execute(dbsql)
	End If
End If
%>
<script>
<!--
//alert("ok");
	parent.fnEditItemSaveEnd('<%=FormatDate(now(),"0000.00.00-00:00")%>');
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->