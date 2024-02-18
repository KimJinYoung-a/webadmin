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

dim sqlStr,DesignerID, waititemid
dim isellvat, ibuyvat, imargin
dim cd1,cd2,cd3,OptionSaveYN, makerid
dim cd2slice,cd3slice,vAddImgNameTotCnt
OptionSaveYN = "N"
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

Dim vAddImgMobileNameReal

'// 한정판매 수량 계산 (옵션 포함) //
dim limitno
limitno = Request.Form("limitno")
if (limitno="") or (Not IsNumeric(limitno)) then limitno=0
    
'#########// 상품 옵션 넣기 //##############################
'옵션이 있을경우 옵션수에 따른 수량 계산
dim iErrMsg
dim itemoptionlist, itemoptionText, itemoptioncount, itemoptionPrice, itemoptionBuyprice, itemoptionLimitNo
dim sellcash,buycash
sellcash=Request.Form("sellcash")
buycash=Request.Form("buycash")
If sellcash="" Then sellcash=0
If buycash="" Then buycash=0

itemoptionlist = split(Request.Form("itemoptioncode2"),"|")
itemoptionText = split(Request.Form("itemoptioncode3"),"|")
itemoptionPrice = split(Request.Form("itemoptioncode4"),"|")
itemoptionBuyprice = split(Request.Form("itemoptioncode5"),"|")
itemoptionLimitNo = split(Request.Form("itemoptioncode6"),"|")
itemoptioncount = ubound(itemoptionlist)

Dim requireimgchk, requireMakeEmail
requireimgchk = Request.Form("requireimgchk")
requireMakeEmail = Request.Form("requireMakeEmail")
If requireimgchk="" Then requireimgchk="N"
If requireimgchk="N" Then requireMakeEmail=""
'###########################################################################
'상품 데이터 입력
'###########################################################################
sqlStr = "insert into db_academy.dbo.tbl_diy_wait_item" + vbCrlf
sqlStr = sqlStr & " (cate_large,cate_mid,cate_small," + vbCrlf
sqlStr = sqlStr & " itemdiv,makerid,itemname,itemcontent," + vbCrlf
sqlStr = sqlStr & " regdate,designercomment,itemsource,itemsize,itemWeight," + vbCrlf
sqlStr = sqlStr & " buycash, sellcash," + vbCrlf
sqlStr = sqlStr & " mileage, sellyn, deliverytype," + vbCrlf
sqlStr = sqlStr & " sourcearea, makername, limityn,limitno,limitsold, oregdate,reRegDate," + vbCrlf
sqlStr = sqlStr & " currstate, keywords,usinghtml, mwdiv, vatYn, ordercomment," + vbCrlf
sqlStr = sqlStr & " upchemanagecode ,cstodr,requireMakeDay,requirecontents,requireimgchk,refundpolicy,infoDiv,safetyYn,safetyDiv,safetyNum,freight_min,freight_max, requireMakeEmail)" + vbCrlf
sqlStr = sqlStr & " values(" + vbCrlf
sqlStr = sqlStr & "'" & cd1 & "'" + vbCrlf
sqlStr = sqlStr & ",'" & cd2 & "'" + vbCrlf
sqlStr = sqlStr & ",'" & cd3 & "'" + vbCrlf
sqlStr = sqlStr & ",'" & Cstr(Request.Form("itemdiv")) & "'" + vbCrlf
sqlStr = sqlStr & ",'" & DesignerID & "'" + vbCrlf
sqlStr = sqlStr & ",convert(varchar(64),'" & chrbyte(html2db(Request.Form("itemname")),64,"") & "')" + vbCrlf
sqlStr = sqlStr & ",'" & html2db(Request.Form("itemcontent")) & "'" + vbCrlf
sqlStr = sqlStr & ",getdate()" + vbCrlf
sqlStr = sqlStr & ",'" & html2db(Request.Form("designercomment")) & "'" + vbCrlf
sqlStr = sqlStr & ",'" & chrbyte(html2db(Request.Form("itemsource")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & chrbyte(html2db(Request.Form("itemsize")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & chrbyte(html2db(Request.Form("itemWeight")),12,"") & "'" + vbCrlf
sqlStr = sqlStr & "," & buycash & "" + vbCrlf
sqlStr = sqlStr & "," & sellcash & "" + vbCrlf
sqlStr = sqlStr & "," & CLng(CLng(sellcash)*0.01) & "" + vbCrlf
sqlStr = sqlStr & ",'N'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("deliverytype") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & chrbyte(html2db(Request.Form("sourcearea")),128,"") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & chrbyte(html2db(Request.Form("makername")),64,"") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("limityn") & "'" + vbCrlf
sqlStr = sqlStr & ","  & limitno & "" + vbCrlf
sqlStr = sqlStr & ",'0'" + vbCrlf
sqlStr = sqlStr & ",getdate()" + vbCrlf
sqlStr = sqlStr & ",getdate()" + vbCrlf
sqlStr = sqlStr & ",'8'" + vbCrlf
sqlStr = sqlStr & ",convert(varchar(128),'" & chrbyte(html2db(Request.Form("keywords")),128,"") & "')" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("usinghtml") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("mwdiv") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & Request.Form("vatYn") & "'" + vbCrlf
sqlStr = sqlStr & ",'" & html2db(Request.Form("ordercomment")) & "'" + vbCrlf
sqlStr = sqlStr & ",convert(varchar(32),'" & chrbyte(html2db(Request.Form("upchemanagecode")),32,"") & "')" + vbCrlf
sqlStr = sqlStr & ", '" & html2db(Request.Form("cstodr")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" & html2db(Request.Form("requireMakeDay")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" & html2db(Request.Form("requirecontents")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" + requireimgchk + "'" + vbCrlf
sqlStr = sqlStr & ", '" & html2db(Request.Form("refundpolicy")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" & Request.Form("infoDiv") & "'" & vbCrlf
sqlStr = sqlStr & ", '" & Request.Form("safetyYn") & "'" & vbCrlf
sqlStr = sqlStr & ", '" & Request.Form("safetyDiv") & "'" & vbCrlf
sqlStr = sqlStr & ", '" & chrbyte(html2db(Request.Form("safetyNum")),24,"") & "'" & vbCrlf
sqlStr = sqlStr & ", '" & getNumeric(Request.Form("freight_min")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" & getNumeric(Request.Form("freight_max")) & "'" & vbCrlf
sqlStr = sqlStr & ", '" & html2db(requireMakeEmail) & "'" & vbCrlf
sqlStr = sqlStr & ")" + vbCrlf
dbACADEMYget.Execute sqlStr
'###########################################################################
'상품 아이디 가져오기
'###########################################################################
sqlStr = "Select IDENT_CURRENT('db_academy.dbo.tbl_diy_wait_item') as maxitemid "
rsACADEMYget.Open sqlStr,dbACADEMYget,1
	waititemid = rsACADEMYget("maxitemid")
rsACADEMYget.close

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
'Response.write Request.Form("itemvideo") & "<br>"
'Response.write itemvideo & "<br>"
'Response.write videoType & "<br>"
'Response.write VideoTempSrc & "<br>"
'Response.write RetSrc & "<br>"
'Response.write VideoTempWidth & "<br>"
'Response.write RetWidth & "<br>"
'Response.write VideoTempHeight & "<br>"
'Response.write RetHeight & "<br>"
	If Not(videoType="etc") Then
		dbsql = " insert into [db_academy].[dbo].[tbl_diy_wait_item_videos] (itemid, videogubun, videotype, videourl, videowidth, videoheight, videofullurl, regdate) values " + vbCrlf
		dbsql = dbsql & " ('"&CStr(waititemid)&"', 'video1', '"&videoType&"', '"&RetSrc&"', '"&RetWidth&"', '"&RetHeight&"','"&chrbyte(html2db(itemvideo),255,"")&"', getdate()) " + vbCrlf
'		Response.write dbsql & "<br>"
'		Response.end
		dbACADEMYget.execute(dbsql)
	End If
'Response.end
End If

'###########################################################################
'전시카테고리 넣기
'###########################################################################
Dim  chkCateDef,sqlStrChk, catecode, catedepth, isDefault
chkCateDef = 0
catecode = Request.Form("catecode")
catedepth = Request.Form("catedepth")
isDefault = Request.Form("isDefault")

If (catecode<>"") Then 
	catecode = Split(catecode,",")
	catedepth = Split(catedepth,",")
	isDefault = Split(isDefault,",")
	sqlStr = "delete from db_academy.dbo.tbl_display_cate_waitItem_Academy Where itemid='" & CStr(waititemid) & "';" & vbCrLf
	for i=0 to ubound(catecode)
		'2015.06.18 수정 (기본카테고리는 하나만 설정되게)
		if  UCase(isDefault(i)) ="Y" and chkCateDef = 1 then
			isDefault(i)  ="N"
		end if
	
		sqlStr = sqlStr & "Insert into db_academy.dbo.tbl_display_cate_waitItem_Academy (catecode, itemid, depth, sortNo, isDefault) values "
		sqlStr = sqlStr & "('" & Cstr(catecode(i)) & "'"
		sqlStr = sqlStr & ",'" & CStr(waititemid) & "'"
		sqlStr = sqlStr & ",'" & CStr(catedepth(i)) & "',9999"
		sqlStr = sqlStr & ",'" & isDefault(i) & "');" & vbCrLf 
		 
		IF UCase(isDefault(i)) ="Y" THEN '기본 카테고리 설정되어 있는지 확인
			chkCateDef = 1 
		END IF
	next
	dbACADEMYget.execute(sqlStr)
end if
'###########################################################################
%>
<script>
<!--
	parent.fntempSaveEnd('<%=waititemid%>','<%=OptionSaveYN%>');
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->