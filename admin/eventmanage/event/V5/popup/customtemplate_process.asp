<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : grouptemplate_process.asp
' Discription : I형 기차 템플릿 수정
' History : 2019.02.13 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , device , idx, saveafter
	Dim slideimg , sorting , isusing
	Dim menuidx
	Dim sqlStr, GroupItemCheck, GroupItemType

	Dim sIdx, sSortNo, sImgURL, i, sItemID, sTitle, sGroupcode
	Dim sBrandid, sItemname, sIconnew, sIconbest, GroupItemPriceView
	dim refer, BGImage, BGImagePC, BGColorLeft, BGColorRight, contentsAlign, textColor
	dim GroupItemTitleName , GroupItemViewType , GroupItemBrandName , saleColor , priceColor , orgpriceColor
	dim smobileimageurl , sPCImageurl
	dim sXposition , sYposition
	dim Margin, MarginBottom, MarginColor, MarginBottomColor

	refer = request.ServerVariables("HTTP_REFERER")
	idx = requestCheckVar(Request.form("idx"),10)
	eventid = requestCheckVar(Request.form("evt_code"),10)
	mode = requestCheckVar(Request.form("mode"),6)
	device = requestCheckVar(Request.form("device"),1)
	slideimg = requestCheckVar(Request.form("slideimg"),200)
	GroupItemCheck = requestCheckVar(Request.form("GroupItemCheck"),1)
	GroupItemType = requestCheckVar(Request.form("GroupItemType"),1)
	GroupItemPriceView = requestCheckVar(Request.form("GroupItemPriceView"),1)
	menuidx = requestCheckvar(request("menuidx"),10)
	BGImage	= requestCheckVar(Request.form("BGImage"),128)
	BGImagePC	= requestCheckVar(Request.form("BGImagePC"),128)
	BGColorLeft	= requestCheckVar(Request.form("BGColorLeft"),8)
	BGColorRight	= requestCheckVar(Request.form("BGColorRight"),8)
	contentsAlign	= requestCheckVar(Request.form("contentsAlign"),1)
	textColor	= requestCheckVar(Request.form("textColor"),8)
	Margin	= requestCheckVar(Request.form("Margin"),10)
	MarginBottom	= requestCheckVar(Request.form("MarginBottom"),10)
	MarginColor = requestCheckVar(Request.form("MarginColor"), 8)
	MarginBottomColor = requestCheckVar(Request.form("MarginBottomColor"), 8)

	GroupItemTitleName = requestCheckVar(Request.form("GroupItemTitleName"),1)
	GroupItemViewType = requestCheckVar(Request.form("GroupItemViewType"),1)
	GroupItemBrandName = requestCheckVar(Request.form("GroupItemBrandName"),1)
	saleColor = requestCheckVar(Request.form("saleColor"), 8)
	priceColor = requestCheckVar(Request.form("priceColor"), 8)
	orgpriceColor = requestCheckVar(Request.form("orgpriceColor"), 8)

	if BGColorLeft="" then BGColorLeft="#FFFFFF"
'Response.write mode & "<br>"
	if BGImage <> "" then
        if checkNotValidHTML(BGImage) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('배경화면이미지에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
        response.write "</script>"
        response.End
        end if
    end If
	if BGImagePC <> "" then
        if checkNotValidHTML(BGImagePC) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('배경화면이미지에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
        response.write "</script>"
        response.End
        end if
    end If
	if not isNumeric(Margin) then Margin=0

Select Case mode
	 Case "TI"
		'slide이미지 신규 등록
		sqlStr = "Insert Into [db_event].[dbo].[tbl_event_multi_contents] " &_
					" (menuidx, imgurl,grouptype) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & slideimg &"','B')"
		dbget.Execute(sqlStr)
		saveafter="TI"
	Case "TU"
		'//리스트에서수정
		for i=1 to request.form("bidx").count
			sIdx = request.form("bidx")(i)
			sSortNo = request.form("viewidx")(i)
			sImgURL = request.form("imgurl")(i)
			sItemID = request.form("itemid")(i)
			sTitle = request.form("title")(i)
			sGroupcode = request.form("groupcode")(i)
			sBrandid = request.form("brandid")(i)
			sItemname = request.form("itemname")(i)
			sIconnew = request.form("iconnew")(i)
			sIconbest = request.form("iconbest")(i)
			smobileimageurl = request.form("mobileImage"&i-1)
			sPCImageurl = request.form("pcImage"&i-1)

			if device = "W" then
				sXposition = request.form("xPosition"&i-1)
				sYposition = request.form("yPosition"&i-1)
			end if 

			if sSortNo="" then sSortNo="0"

			sqlStr = sqlStr & " Update [db_event].[dbo].[tbl_event_multi_contents]" & vbCrLf
			sqlStr = sqlStr & " Set viewidx=" & sSortNo & ""& vbCrLf
			sqlStr = sqlStr & " ,imgurl='" & sImgURL & "'"& vbCrLf
			sqlStr = sqlStr & " ,itemid='" & sItemID & "'"& vbCrLf
			if sTitle <>"" then
			sqlStr = sqlStr & " ,title='" & sTitle & "'"& vbCrLf
			end if
			if sGroupcode <>"" then
			sqlStr = sqlStr & " ,groupcode='" & sGroupcode & "'"& vbCrLf
			end if
			if sBrandid <>"" then
			sqlStr = sqlStr & " ,makerid='" & sBrandid & "'"& vbCrLf
			end if
			if sItemname <>"" then
			sqlStr = sqlStr & " ,itemname='" & sItemname & "'"& vbCrLf
			end if
			if sIconnew <>"" then
			sqlStr = sqlStr & " ,iconnew='" & sIconnew & "'"& vbCrLf
			end if
			if sIconbest <>"" then
			sqlStr = sqlStr & " ,iconbest='" & sIconbest & "'"& vbCrLf
			end if

			if device = "M" then 
				sqlStr = sqlStr & " , mobileimageurl='" & smobileimageurl & "'"& vbCrLf
			else
				sqlStr = sqlStr & " , pcimageurl='" & sPCImageurl & "'"& vbCrLf
				sqlStr = sqlStr & " , xposition='" & sXposition & "'"& vbCrLf
				sqlStr = sqlStr & " , yposition='" & sYposition & "'"& vbCrLf
			end if 

			sqlStr = sqlStr & " Where idx='" & sIdx & "';"
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		End If 
		saveafter="TU"
	Case "TD" '삭제
		sIdx = request.form("bidx")

		sqlStr = "delete from db_event.dbo.tbl_event_multi_contents Where idx='"& sIdx &"'"
		dbget.Execute sqlStr
End Select

	sqlStr = " Update [db_event].[dbo].[tbl_event_multi_contents_master]"
	sqlStr = sqlStr & " Set GroupItemType='" + GroupItemType + "'"
	sqlStr = sqlStr & " , GroupItemCheck='" + Cstr(GroupItemCheck) + "'"
	sqlStr = sqlStr & " , GroupItemPriceView='" + Cstr(GroupItemPriceView) + "'"
	If device = "M" Then
		sqlStr = sqlStr & " , BGImage='" & BGImage & "'" & vbCrLf
		sqlStr = sqlStr & " , Margin='" & Margin & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginColor='" & MarginColor & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginBottom='" & MarginBottom & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginBottomColor='" & MarginBottomColor & "'" & vbCrLf
    Else
		sqlStr = sqlStr & " , BGImagePC='" & BGImagePC & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginPC='" & Margin & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginColorPC='" & MarginColor & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginBottomPC='" & MarginBottom & "'" & vbCrLf
		sqlStr = sqlStr & " , MarginBottomColorPC='" & MarginBottomColor & "'" & vbCrLf
	End If
	sqlStr = sqlStr & " , BGColorLeft='" & BGColorLeft & "'" & vbCrLf
	sqlStr = sqlStr & " , BGColorRight='" & BGColorRight & "'" & vbCrLf
    sqlStr = sqlStr & " , contentsAlign='" & contentsAlign & "'" & vbCrLf
	sqlStr = sqlStr & " , textColor='" & trim(textColor) & "'" & vbCrLf
	sqlStr = sqlStr & " , GroupItemTitleName='" & trim(GroupItemTitleName) & "'" & vbCrLf
	sqlStr = sqlStr & " , GroupItemViewType='" & trim(GroupItemViewType) & "'" & vbCrLf
	sqlStr = sqlStr & " , GroupItemBrandName='" & trim(GroupItemBrandName) & "'" & vbCrLf
	sqlStr = sqlStr & " , saleColor='" & saleColor & "'" & vbCrLf
	sqlStr = sqlStr & " , priceColor='" & priceColor & "'" & vbCrLf
	sqlStr = sqlStr & " , orgpriceColor='" & orgpriceColor & "'" & vbCrLf
	sqlStr = sqlStr & " Where idx=" & menuidx
	dbget.Execute sqlStr

    '--3.theme 수정
    sqlStr = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
    sqlStr = sqlStr + " SET contentsAlign='" & contentsAlign & "'" & vbCrlf
    sqlStr = sqlStr + " where evt_code=" & eventid
    dbget.execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eventid) + "&togglediv=5&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
	response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->