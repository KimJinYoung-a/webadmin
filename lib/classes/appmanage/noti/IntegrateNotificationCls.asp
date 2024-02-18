<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.14 한용민 생성
'###########################################################

Class cNotiItem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FsIdx
	public FnotiType
	public FlinkCode
	public FstartDate
	public FendDate
	public FreserveTime
    public frectisusing
	public FpushIsusing
	public FkakaoAlrimIsusing
	public Fpushtitle
	public Fpushcontents
	public Fpushurl
	public FtemplateCode
	public Fcontents
	public Fbutton_name
	public Fbutton_url_mobile
	public Fbutton_name2
	public Fbutton_url_mobile2
	public Ffailed_type
	public Ffailed_subject
	public Ffailed_msg
	public Fetc_template_code
	public Fmember_smsok_checkyn
	public Fmember_kakaoalrimyn_checkyn
	public FregDate
	public FlastUpdate
	public FadminUserid
	public FlastUserid
	public Fisusing
	public FpushTestCount
	public FkakaoAlrimTestCount
	public fnIdx
	public fsendType
	public fuserId
	public fdevice
	public freplaceItemId
	public freplaceMileage
end class

class cNotiList
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	Public FOneItem

	public farrList
    public frectreservationdate
    public frectlinkCode
    public frectisusing
    public frectpushIsusing
    public frectkakaoAlrimIsusing
    public FrectsIdx
	public frectnotiType
	public frectsendType
	public FrectnIdx
	public frectuserId

	' /admin/appmanage/noti/popIntegrateNotificationEdit.asp
	public sub fIntegrateNotificationOne()
		dim sqlStr, i, sqlsearch

		if FrectnIdx <> "" then
			sqlsearch = sqlsearch & " and n.nIdx="&FrectnIdx&""
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " nIdx, notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by nIdx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new cNotiItem
		if Not rsget.Eof then
			FOneItem.fnIdx = rsget("nIdx")
			FOneItem.fnotiType = rsget("notiType")
			FOneItem.flinkCode = rsget("linkCode")
			FOneItem.fsendType = rsget("sendType")
			FOneItem.fuserId = rsget("userId")
			FOneItem.fdevice = rsget("device")
			FOneItem.fregDate = rsget("regDate")
			FOneItem.flastUpdate = rsget("lastUpdate")
			FOneItem.freplaceItemId = rsget("replaceItemId")
			FOneItem.freplaceMileage = rsget("replaceMileage")
			FOneItem.fisusing = rsget("isusing")

		end if

		rsget.Close
	end Sub

	' /admin/appmanage/noti/IntegrateNotification.asp
	' 밑에 함수를 수정할경우 fIntegrateNotificationListNotPaging 함수도 똑같이 수정해야 한다.
	public sub fIntegrateNotificationList()
		dim sqlStr, i, sqlsearch

		if frectsendType <> "" then
			sqlsearch = sqlsearch & " and n.sendType = '"& frectsendType &"'"
		end If
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and n.isusing = '"& frectisusing &"'"
		end If
		if frectlinkCode <> "" then
			sqlsearch = sqlsearch & " and n.linkCode = "& frectlinkCode &""
		end If
		if frectnotiType <> "" then
			sqlsearch = sqlsearch & " and n.notiType = '"& frectnotiType &"'"
		end If
		if frectuserId <> "" then
			sqlsearch = sqlsearch & " and n.userId = '"& frectuserId &"'"
		end If

		sqlStr = "select count(nIdx) as cnt, CEILING(CAST(Count(nIdx) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " nIdx, notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by n.nIdx desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cNotiItem

				FItemList(i).fnIdx = rsget("nIdx")
				FItemList(i).fnotiType = rsget("notiType")
				FItemList(i).flinkCode = rsget("linkCode")
				FItemList(i).fsendType = rsget("sendType")
				FItemList(i).fuserId = rsget("userId")
				FItemList(i).fdevice = rsget("device")
				FItemList(i).fregDate = rsget("regDate")
				FItemList(i).flastUpdate = rsget("lastUpdate")
				FItemList(i).freplaceItemId = rsget("replaceItemId")
				FItemList(i).freplaceMileage = rsget("replaceMileage")
				FItemList(i).fisusing = rsget("isusing")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/noti/IntegrateNotificationExcel.asp
	' 밑에 함수를 수정할경우 fIntegrateNotificationList 함수도 똑같이 수정해야 한다.
	public sub fIntegrateNotificationListNotPaging()
		dim sqlStr, i, sqlsearch

		if frectsendType <> "" then
			sqlsearch = sqlsearch & " and n.sendType = '"& frectsendType &"'"
		end If
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and n.isusing = '"& frectisusing &"'"
		end If
		if frectlinkCode <> "" then
			sqlsearch = sqlsearch & " and n.linkCode = "& frectlinkCode &""
		end If
		if frectnotiType <> "" then
			sqlsearch = sqlsearch & " and n.notiType = '"& frectnotiType &"'"
		end If
		if frectuserId <> "" then
			sqlsearch = sqlsearch & " and n.userId = '"& frectuserId &"'"
		end If

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " nIdx, notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by n.nIdx desc"
		
		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly  ''2016/04/06

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount

		i=0
		if  not rsget.EOF  then
			fArrLIst = rsget.getrows()
		end if
	end Sub

	' admin/appmanage/noti/IntegrateNotificationSchedule.asp
	public sub fIntegrateNotificationScheduleList()
		dim sqlStr, i, sqlsearch

		if frectreservationdate <> "" then
			sqlsearch = sqlsearch & " and '"& frectreservationdate &" 00:00:00' between s.startDate and s.endDate"
		end If
		if frectlinkCode <> "" then
			sqlsearch = sqlsearch & " and s.linkCode = '"& frectlinkCode &"'"
		end If
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and s.isusing = '"& frectisusing &"'"
		end If
		if frectpushIsusing <> "" then
			sqlsearch = sqlsearch & " and s.pushIsusing = '"& frectpushIsusing &"'"
		end If
		if frectkakaoAlrimIsusing <> "" then
			sqlsearch = sqlsearch & " and s.kakaoAlrimIsusing = '"& frectkakaoAlrimIsusing &"'"
		end If
		if frectnotiType <> "" then
			sqlsearch = sqlsearch & " and s.notiType = '"& frectnotiType &"'"
		end If

		sqlStr = "select count(sIdx) as cnt, CEILING(CAST(Count(sIdx) AS FLOAT)/'"&FPageSize&"' ) as totPg"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotificationSchedule as s with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		

		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		if FTotalCount < 1 then exit sub
		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " sIdx, notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents"
        sqlStr = sqlStr & " , pushurl, templateCode, contents, button_name, button_url_mobile, button_name2"
        sqlStr = sqlStr & " , button_url_mobile2, failed_type, failed_subject, failed_msg, etc_template_code"
        sqlStr = sqlStr & " , member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate, lastUpdate, adminUserid"
        sqlStr = sqlStr & " , lastUserid, isusing, pushTestCount, kakaoAlrimTestCount"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotificationSchedule as s with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch		
		sqlStr = sqlStr & " order by s.sidx desc"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cNotiItem

				FItemList(i).fsIdx = rsget("sIdx")
				FItemList(i).fnotiType = rsget("notiType")
				FItemList(i).flinkCode = rsget("linkCode")
				FItemList(i).fstartDate = rsget("startDate")
				FItemList(i).fendDate = rsget("endDate")
				FItemList(i).freserveTime = rsget("reserveTime")
                FItemList(i).fpushIsusing = rsget("pushIsusing")
                FItemList(i).fkakaoAlrimIsusing = rsget("kakaoAlrimIsusing")
				FItemList(i).fpushtitle = db2html(rsget("pushtitle"))
				FItemList(i).fpushcontents = db2html(rsget("pushcontents"))
				FItemList(i).fpushurl = db2html(rsget("pushurl"))
				FItemList(i).ftemplateCode = rsget("templateCode")
				FItemList(i).fcontents = db2html(rsget("contents"))
				FItemList(i).fbutton_name = db2html(rsget("button_name"))
				FItemList(i).fbutton_url_mobile = db2html(rsget("button_url_mobile"))
				FItemList(i).fbutton_name2 = db2html(rsget("button_name2"))
				FItemList(i).fbutton_url_mobile2 = db2html(rsget("button_url_mobile2"))
				FItemList(i).ffailed_type = rsget("failed_type")
				FItemList(i).ffailed_subject = db2html(rsget("failed_subject"))
				FItemList(i).ffailed_msg = db2html(rsget("failed_msg"))
				FItemList(i).fetc_template_code = rsget("etc_template_code")
				FItemList(i).fmember_smsok_checkyn = rsget("member_smsok_checkyn")
				FItemList(i).fmember_kakaoalrimyn_checkyn = rsget("member_kakaoalrimyn_checkyn")
				FItemList(i).fregDate = rsget("regDate")
				FItemList(i).flastUpdate = rsget("lastUpdate")
				FItemList(i).fadminUserid = rsget("adminUserid")
				FItemList(i).flastUserid = rsget("lastUserid")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fpushTestCount = rsget("pushTestCount")
				FItemList(i).fkakaoAlrimTestCount = rsget("kakaoAlrimTestCount")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Sub

	' /admin/appmanage/push/msg/popIntegrateNotificationScheduleEdit.asp
	public sub fIntegrateNotificationScheduleOne()
		dim sqlStr, i, sqlsearch

		if FrectsIdx <> "" then
			sqlsearch = sqlsearch & " and s.sIdx="&FrectsIdx&""
		end if

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " sIdx, notiType, linkCode, startDate, endDate, reserveTime, pushIsusing, kakaoAlrimIsusing, pushtitle, pushcontents"
        sqlStr = sqlStr & " , pushurl, templateCode, contents, button_name, button_url_mobile, button_name2"
        sqlStr = sqlStr & " , button_url_mobile2, failed_type, failed_subject, failed_msg, etc_template_code"
        sqlStr = sqlStr & " , member_smsok_checkyn, member_kakaoalrimyn_checkyn, regDate, lastUpdate, adminUserid"
        sqlStr = sqlStr & " , lastUserid, isusing, pushTestCount, kakaoAlrimTestCount"
		sqlStr = sqlStr & " from db_contents.dbo.tbl_IntegrateNotificationSchedule as s with (nolock)"
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by sIdx Desc"
		
		'response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		set FOneItem = new cNotiItem
		if Not rsget.Eof then
            FOneItem.fsIdx = rsget("sIdx")
            FOneItem.fnotiType = rsget("notiType")
            FOneItem.flinkCode = rsget("linkCode")
            FOneItem.fstartDate = rsget("startDate")
            FOneItem.fendDate = rsget("endDate")
            FOneItem.freserveTime = rsget("reserveTime")
            FOneItem.fpushIsusing = rsget("pushIsusing")
            FOneItem.fkakaoAlrimIsusing = rsget("kakaoAlrimIsusing")
            FOneItem.fpushtitle = db2html(rsget("pushtitle"))
            FOneItem.fpushcontents = db2html(rsget("pushcontents"))
            FOneItem.fpushurl = db2html(rsget("pushurl"))
            FOneItem.ftemplateCode = rsget("templateCode")
            FOneItem.fcontents = db2html(rsget("contents"))
            FOneItem.fbutton_name = db2html(rsget("button_name"))
            FOneItem.fbutton_url_mobile = db2html(rsget("button_url_mobile"))
            FOneItem.fbutton_name2 = db2html(rsget("button_name2"))
            FOneItem.fbutton_url_mobile2 = db2html(rsget("button_url_mobile2"))
            FOneItem.ffailed_type = rsget("failed_type")
            FOneItem.ffailed_subject = db2html(rsget("failed_subject"))
            FOneItem.ffailed_msg = db2html(rsget("failed_msg"))
            FOneItem.fetc_template_code = rsget("etc_template_code")
            FOneItem.fmember_smsok_checkyn = rsget("member_smsok_checkyn")
            FOneItem.fmember_kakaoalrimyn_checkyn = rsget("member_kakaoalrimyn_checkyn")
            FOneItem.fregDate = rsget("regDate")
            FOneItem.flastUpdate = rsget("lastUpdate")
            FOneItem.fadminUserid = rsget("adminUserid")
            FOneItem.flastUserid = rsget("lastUserid")
            FOneItem.fisusing = rsget("isusing")
			FOneItem.fpushTestCount = rsget("pushTestCount")
			FOneItem.fkakaoAlrimTestCount = rsget("kakaoAlrimTestCount")

		end if

		rsget.Close
	end Sub

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

function DrawNotiType(selBoxName,selVal,chplg)
%>
<select name="<%= selBoxName %>" <%= chplg %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='EVENT' <% if cstr(selVal)=cstr("EVENT") then response.write " selected" %> >이벤트</option>
	<option value='EXHIBITION' <% if cstr(selVal)=cstr("EXHIBITION") then response.write " selected" %> >통합기획전</option>
</select>
<%
end Function

function getNotiType(notiType)
    dim resultNotiType

    if notiType="EVENT" then
        resultNotiType="이벤트"
    elseif notiType="EXHIBITION" then
        resultNotiType="통합기획전"
    else
        resultNotiType=""
    end if

    getNotiType=resultNotiType
End function

function DrawsendType(selBoxName,selVal,chplg)
%>
<select name="<%= selBoxName %>" <%= chplg %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='KAKAOALRIM' <% if cstr(selVal)=cstr("KAKAOALRIM") then response.write " selected" %> >카카오알림톡</option>
	<option value='PUSH' <% if cstr(selVal)=cstr("PUSH") then response.write " selected" %> >푸시</option>
</select>
<%
end Function

function getSendType(sendType)
    dim resultSendType

    if sendType="KAKAOALRIM" then
        resultSendType="카카오알림톡"
    elseif sendType="PUSH" then
        resultSendType="푸시"
    else
        resultSendType=""
    end if

    getSendType=resultSendType
End function

function DrawIntegrateNotificationDevice(selBoxName,selVal,chplg)
%>
<select name="<%= selBoxName %>" <%= chplg %>>
	<option value='' <% if selVal="" then response.write " selected" %> >전체</option>
	<option value='W' <% if cstr(selVal)=cstr("W") then response.write " selected" %> >PC웹</option>
	<option value='M' <% if cstr(selVal)=cstr("M") then response.write " selected" %> >모바일웹</option>
	<option value='A' <% if cstr(selVal)=cstr("A") then response.write " selected" %> >앱</option>
</select>
<%
end Function

function getIntegrateNotificationDevice(sendType)
    dim resultDevice

    if sendType="W" then
        resultDevice="PC웹"
    elseif sendType="M" then
        resultDevice="모바일웹"
    elseif sendType="A" then
        resultDevice="앱"
    else
        resultDevice=""
    end if

    getIntegrateNotificationDevice=resultDevice
End function
%>

