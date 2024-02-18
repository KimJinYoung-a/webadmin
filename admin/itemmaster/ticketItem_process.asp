<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/ticketItemCls.asp"-->

<%
Dim mode : mode = requestCheckvar(request("mode"),32)
Dim ticketPlaceIdx : ticketPlaceIdx = requestCheckvar(request("ticketPlaceIdx"),10)
Dim ticketPlaceName: ticketPlaceName = requestCheckvar(request("ticketPlaceName"),100)
Dim tPAddress      : tPAddress = requestCheckvar(request("tPAddress"),120)
Dim tPTel          : tPTel = requestCheckvar(request("tPTel"),20)
Dim tPHomeURL      : tPHomeURL = requestCheckvar(request("tPHomeURL"),120)
Dim placeImg       : placeImg = requestCheckvar(request("placeImg"),100)
Dim contentsImage1 : contentsImage1 = requestCheckvar(request("contentsImage1"),100)
Dim contentsImage2 : contentsImage2 = requestCheckvar(request("contentsImage2"),100)
Dim contentsImage3 : contentsImage3 = requestCheckvar(request("contentsImage3"),100)
Dim brd_content    : brd_content = request("brd_content")
Dim parkingGuide    : parkingGuide = request("parkingguide")

Dim itemid : itemid = requestCheckvar(request("itemid"),10)
Dim txGenre : txGenre  = requestCheckvar(request("txGenre"),30)
Dim txGrade : txGrade  = requestCheckvar(request("txGrade"),64)
Dim txRunTime : txRunTime  = requestCheckvar(request("txRunTime"),32)
Dim bookingStDt : bookingStDt = requestCheckvar(request("bookingStDt"),10) + " " + requestCheckvar(request("bookingStDtTime"),8)
Dim bookingEdDt : bookingEdDt = requestCheckvar(request("bookingEdDt"),10) + " " + requestCheckvar(request("bookingEdDtTime"),8)
Dim stDt : stDt  = requestCheckvar(request("stDt"),10)
Dim edDt : edDt  = requestCheckvar(request("edDt"),10)
Dim txplayTimInfo : txplayTimInfo  = requestCheckvar(request("txplayTimInfo"),250)
Dim bookingCharge : bookingCharge  = requestCheckvar(request("bookingCharge"),10)
Dim ticketDlvType : ticketDlvType  = requestCheckvar(request("ticketDlvType"),10)
Dim refundInfoType : refundInfoType = requestCheckvar(request("refundInfoType"),10)

Dim Tk_itemoption : Tk_itemoption = request("Tk_itemoption") ''Array
Dim Tk_StSchedule : Tk_StSchedule = request("Tk_StSchedule") ''Array
Dim Tk_StScheduleTime : Tk_StScheduleTime = request("Tk_StScheduleTime") ''Array
Dim Tk_EdSchedule : Tk_EdSchedule = request("Tk_EdSchedule") ''Array
Dim Tk_EdScheduleTime : Tk_EdScheduleTime = request("Tk_EdScheduleTime") ''Array

Dim itemdiv : itemdiv = requestCheckvar(request("itemdiv"),10)

Dim optArr : optArr = split(Tk_itemoption,",")
Dim i

Function chkPlayDate(ADate,BDate, alertMsg)
    On Error Resume Next
    IF (CDate(ADate)>CDate(BDate)) then
        IF Err Then
            response.write "<script>alert('날짜 형식 오류');history.back();</script>"
            dbget.Close() : response.end
            On Error Goto 0
        End if
        response.write "<script>alert('"&alertMsg&"');history.back();</script>"
        dbget.Close() : response.end
    end If
    
    On Error Goto 0
end Function

function getReturnExpireDate(stDateTime,itemdiv)
    Dim sqlStr, retDateTime
    
    if (itemdiv="18") then
        retDateTime = Left(DateAdd("d",-7,stDateTime),10) + " 23:59:59"
        getReturnExpireDate = retDateTime
        exit function
    end if
    
    retDateTime = Left(DateAdd("d",-1,stDateTime),10) + " 18:00:00"
    
    sqlStr = " select top 1 (convert(varchar(10),solar_date,21) + ' ' + CASE WHEN holiday=1 THEN '11:00:00' ELSE '18:00:00' END) as expiredDate"
    sqlStr = sqlStr & " from db_sitemaster.dbo.LunarToSolar"
    sqlStr = sqlStr & " where solar_date<'"&stDateTime&"'"
    sqlStr = sqlStr & " and holiday<1"                              '' 토요일 1, 공휴일2   토요일 도 가능하게 할경우  holiday<2로 세팅
    sqlStr = sqlStr & " order by solar_date desc"
    
    rsget.CursorLocation = adUseClient                            
	rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if Not (Rsget.Eof) then
	    retDateTime = rsget("expiredDate")
	end if
	Rsget.Close  
	        
	getReturnExpireDate = retDateTime
end Function


IF (mode="ticketInfo") then
    ''Check DateTime
    Call chkPlayDate(stDt,edDt,"총공연기간- 종료일이 시작일보다 이전 날짜 일 수 없습니다.")
    
    Call chkPlayDate(bookingStDt,bookingEdDt,"예약일정- 종료일이 시작일보다 이전 날짜 일 수 없습니다.")
    
    
    If IsArray(optArr) THEN
        Tk_StSchedule       = split(Tk_StSchedule,",")
        Tk_StScheduleTime   = split(Tk_StScheduleTime,",")
        Tk_EdSchedule       = split(Tk_EdSchedule,",")
        Tk_EdScheduleTime   = split(Tk_EdScheduleTime,",")
        
        for i=LBound(optArr) to UBound(optArr)
            Call chkPlayDate(Tk_StSchedule(i) + " " + Tk_StScheduleTime(i),Tk_EdSchedule(i) + " " + Tk_EdScheduleTime(i),"관람일정- 종료일이 시작일보다 이전 날짜 일 수 없습니다.")
        
            Call chkPlayDate(stDt,Tk_StSchedule(i) + " " + Tk_StScheduleTime(i),"관람일정은 총 공연기간 내에 만 가능합니다.")
            
            Call chkPlayDate(Tk_StSchedule(i) + " " + Tk_StScheduleTime(i),edDt + " 23:59:59","관람일정은 총 공연기간 내에 만 가능합니다.")
            
            Call chkPlayDate(Tk_EdSchedule(i) + " " + Tk_EdScheduleTime(i),edDt + " 23:59:59","관람일정은 총 공연기간 내에 만 가능합니다.")
            
            Call chkPlayDate(stDt,Tk_EdSchedule(i) + " " + Tk_EdScheduleTime(i),"관람일정은 총 공연기간 내에 만 가능합니다.")
        next
    ELSE
        optArr = Trim(optArr)
        IF (optArr<>"0000") Then
            response.write "<script>alert('옵션코드오류');history.back();</script>"
            response.end
        end if
        Call chkPlayDate(Tk_StSchedule + " " + Tk_StScheduleTime,Tk_EdSchedule + " " + Tk_EdScheduleTime,"관람일정- 종료일이 시작일보다 이전 날짜 일 수 없습니다.")
        
        Call chkPlayDate(stDt,Tk_StSchedule + " " + Tk_StScheduleTime,"관람일정은 총 공연기간 내에 만 가능합니다.")
        
        Call chkPlayDate(Tk_StSchedule + " " + Tk_StScheduleTime,edDt + " 23:59:59","관람일정은 총 공연기간 내에 만 가능합니다.")
        
        Call chkPlayDate(Tk_EdSchedule + " " + Tk_EdScheduleTime,edDt + " 23:59:59","관람일정은 총 공연기간 내에 만 가능합니다.")
        
        Call chkPlayDate(stDt,Tk_EdSchedule + " " + Tk_EdScheduleTime,"관람일정은 총 공연기간 내에 만 가능합니다.")
        
    end if
end if

IF (bookingCharge="") then bookingCharge=0
''IF (ticketDlvType="") then ticketDlvType=0
IF (refundInfoType="") then refundInfoType=0

Dim sqlStr, AssignedRow

IF (mode="ticketInfo") then
    sqlStr = "update db_item.dbo.tbl_ticket_itemInfo" & VbCrlf
    sqlStr = sqlStr & " set stDt='"&stDt&"'"& VbCrlf
    sqlStr = sqlStr & " , edDt='"&edDt&"'"& VbCrlf
    sqlStr = sqlStr & " , bookingStDt='"&bookingStDt&"'"& VbCrlf
    sqlStr = sqlStr & " , bookingEdDt='"&bookingEdDt&"'"& VbCrlf
    sqlStr = sqlStr & " , bookingCharge="&bookingCharge& VbCrlf
    sqlStr = sqlStr & " , ticketDlvType="&ticketDlvType& VbCrlf
    sqlStr = sqlStr & " , refundInfoType="&refundInfoType& VbCrlf
    sqlStr = sqlStr & " , ticketPlaceIdx="&ticketPlaceIdx& VbCrlf
    sqlStr = sqlStr & " , txplayTimInfo='"&html2DB(txplayTimInfo)&"'"&VbCrlf
    sqlStr = sqlStr & " , txGenre='"&html2DB(txGenre)&"'"&VbCrlf
    sqlStr = sqlStr & " , txGrade='"&html2DB(txGrade)&"'"&VbCrlf
    sqlStr = sqlStr & " , txRunTime='"&html2DB(txRunTime)&"'"&VbCrlf
    sqlStr = sqlStr & " where itemid="& itemid
    
    dbget.Execute sqlStr,AssignedRow
    
    if (AssignedRow<1) then
        sqlStr = " insert into db_item.dbo.tbl_ticket_itemInfo" & VbCrlf
        sqlStr = sqlStr & " (itemid,stDt,edDt,bookingStDt,bookingEdDt" & VbCrlf
        sqlStr = sqlStr & " ,bookingCharge, ticketDlvType, refundInfoType" & VbCrlf
        sqlStr = sqlStr & " ,ticketPlaceIdx, txplayTimInfo" & VbCrlf
        sqlStr = sqlStr & " ,txGenre, txGrade, txRunTime" & VbCrlf
        sqlStr = sqlStr & " )"& VbCrlf
        sqlStr = sqlStr & " values("& VbCrlf
        sqlStr = sqlStr & " "&itemid& VbCrlf
        sqlStr = sqlStr & " ,'"&stDt&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&edDt&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&bookingStDt&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&bookingEdDt&"'"& VbCrlf
        sqlStr = sqlStr & " ,"&bookingCharge&""& VbCrlf
        sqlStr = sqlStr & " ,"&ticketDlvType&""& VbCrlf
        sqlStr = sqlStr & " ,"&refundInfoType&""& VbCrlf
        sqlStr = sqlStr & " ,"&ticketPlaceIdx&""& VbCrlf
        sqlStr = sqlStr & " ,'"&html2DB(txplayTimInfo)&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&html2DB(txGenre)&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&html2DB(txGrade)&"'"& VbCrlf
        sqlStr = sqlStr & " ,'"&html2DB(txRunTime)&"'"& VbCrlf
        sqlStr = sqlStr & ")"
        
        dbget.Execute sqlStr
    end if
    
    If (IsArray(optArr)) then
        for i=LBound(optArr) to UBound(optArr)
            optArr(i) = trim(optArr(i))
            sqlStr = "update db_item.dbo.tbl_ticket_Schedule"& VbCrlf
            sqlStr = sqlStr & " set Tk_StSchedule='"& Trim(Tk_StSchedule(i)) + " " + Trim(Tk_StScheduleTime(i)) & "'"& VbCrlf
            sqlStr = sqlStr & " , Tk_EdSchedule='"& Trim(Tk_EdSchedule(i)) + " " + Trim(Tk_EdScheduleTime(i)) & "'"& VbCrlf
            if (itemdiv="18") then
                sqlStr = sqlStr & " , returnExpireDate='"& getReturnExpireDate(Trim(Tk_StSchedule(i)),itemdiv)&"'"&VbCrlf
            else
                sqlStr = sqlStr & " , returnExpireDate='"& getReturnExpireDate(Trim(Tk_EdSchedule(i)),itemdiv)&"'"&VbCrlf  '' Tk_StSchedule(i) => Tk_EdSchedule(i) 2017/08/09
            end if
            sqlStr = sqlStr & " where Tk_itemid="&itemid& VbCrlf
            sqlStr = sqlStr & " and Tk_itemoption='"&optArr(i)&"'"
            dbget.Execute sqlStr,AssignedRow
            
            if (AssignedRow<1) then
                sqlStr = "insert into db_item.dbo.tbl_ticket_Schedule"& VbCrlf
                sqlStr = sqlStr & " (Tk_itemid,Tk_itemoption,Tk_StSchedule,Tk_EdSchedule,returnExpireDate)"& VbCrlf
                sqlStr = sqlStr & " values("
                sqlStr = sqlStr & " "&itemid& VbCrlf
                sqlStr = sqlStr & " ,'"&optArr(i)&"'"& VbCrlf
                sqlStr = sqlStr & " ,'"& Trim(Tk_StSchedule(i)) + " " + Trim(Tk_StScheduleTime(i)) &"'"& VbCrlf
                sqlStr = sqlStr & " ,'"& Trim(Tk_EdSchedule(i)) + " " + Trim(Tk_EdScheduleTime(i)) &"'"& VbCrlf
                if (itemdiv="18") then
                    sqlStr = sqlStr & " ,'"&getReturnExpireDate(Trim(Tk_StSchedule(i)),itemdiv)&"'"& VbCrlf
                else
                    sqlStr = sqlStr & " ,'"&getReturnExpireDate(Trim(Tk_EdSchedule(i)),itemdiv)&"'"& VbCrlf               '' Tk_StSchedule(i) => Tk_EdSchedule(i) 2017/08/09
                end if
                sqlStr = sqlStr & " )"
                dbget.Execute sqlStr
            end if
        Next
    ELSE
        optArr = trim(optArr)
        sqlStr = "update db_item.dbo.tbl_ticket_Schedule"& VbCrlf
        sqlStr = sqlStr & " set Tk_StSchedule='"& Trim(Tk_StSchedule) + " " + Trim(Tk_StScheduleTime) & "'"& VbCrlf
        sqlStr = sqlStr & " , Tk_EdSchedule='"& Trim(Tk_EdSchedule) + " " + Trim(Tk_EdScheduleTime) & "'"& VbCrlf
        if (itemdiv="18") then
            sqlStr = sqlStr & " , returnExpireDate='"& getReturnExpireDate(Trim(Tk_StSchedule),itemdiv)&"'"&VbCrlf
        else
            sqlStr = sqlStr & " , returnExpireDate='"& getReturnExpireDate(Trim(Tk_EdSchedule),itemdiv)&"'"&VbCrlf      '' Tk_StSchedule(i) => Tk_EdSchedule(i) 2017/08/09
        end if
        sqlStr = sqlStr & " where Tk_itemid="&itemid& VbCrlf
        sqlStr = sqlStr & " and Tk_itemoption='"&optArr&"'"
        dbget.Execute sqlStr,AssignedRow
        
        if (AssignedRow<1) then
            sqlStr = "insert into db_item.dbo.tbl_ticket_Schedule"& VbCrlf
            sqlStr = sqlStr & " (Tk_itemid,Tk_itemoption,Tk_StSchedule,Tk_EdSchedule,returnExpireDate)"& VbCrlf
            sqlStr = sqlStr & " values("
            sqlStr = sqlStr & " "&itemid& VbCrlf
            sqlStr = sqlStr & " ,'"&optArr&"'"& VbCrlf
            sqlStr = sqlStr & " ,'"& Trim(Tk_StSchedule) + " " + Trim(Tk_StScheduleTime) &"'"& VbCrlf
            sqlStr = sqlStr & " ,'"& Trim(Tk_EdSchedule) + " " + Trim(Tk_EdScheduleTime) &"'"& VbCrlf
            if (itemdiv="18") then
                sqlStr = sqlStr & " ,'"&getReturnExpireDate(Trim(Tk_StSchedule),itemdiv)&"'"& VbCrlf
            else
                sqlStr = sqlStr & " ,'"&getReturnExpireDate(Trim(Tk_EdSchedule),itemdiv)&"'"& VbCrlf            '' Tk_StSchedule(i) => Tk_EdSchedule(i) 2017/08/09
            end if
            sqlStr = sqlStr & " )"
            dbget.Execute sqlStr
        end if
    end if

elseIF (mode="ticketPlace") then
    if (ticketPlaceIdx<>"0") then
        sqlStr = "update db_item.dbo.tbl_ticket_placeInfo" & VbCrlf
        sqlStr = sqlStr & " set ticketPlaceName='"& html2DB(ticketPlaceName) & "'" & VbCrlf
        sqlStr = sqlStr & " ,tPAddress  ='"& html2DB(tPAddress) & "'" & VbCrlf
        sqlStr = sqlStr & " ,tPTel      ='"& html2DB(tPTel) & "'" & VbCrlf
        sqlStr = sqlStr & " ,tPHomeURL  ='"& html2DB(tPHomeURL) & "'" & VbCrlf
        sqlStr = sqlStr & " ,placeImgURL   ='"& html2DB(placeImg) & "'" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage1 ='"& html2DB(contentsImage1) & "'" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage2 ='"& html2DB(contentsImage2) & "'" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage3 ='"& html2DB(contentsImage3) & "'" & VbCrlf
        sqlStr = sqlStr & " ,placeContents ='"& html2DB(brd_content) & "'" & VbCrlf
        sqlStr = sqlStr & " ,parkingGuide ='"& html2DB(parkingGuide) & "'" & VbCrlf
        sqlStr = sqlStr & " where ticketPlaceIdx="&ticketPlaceIdx
        
        dbget.Execute sqlStr
        
    else
        sqlStr = "insert into db_item.dbo.tbl_ticket_placeInfo" & VbCrlf
        sqlStr = sqlStr & " (ticketPlaceName" & VbCrlf
        sqlStr = sqlStr & " ,tPAddress" & VbCrlf
        sqlStr = sqlStr & " ,tPTel" & VbCrlf
        sqlStr = sqlStr & " ,tPHomeURL" & VbCrlf
        sqlStr = sqlStr & " ,placeImgURL" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage1" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage2" & VbCrlf
        sqlStr = sqlStr & " ,placecontentsImage3" & VbCrlf
        sqlStr = sqlStr & " ,placeContents" & VbCrlf
        sqlStr = sqlStr & " ,parkingGuide" & VbCrlf        
        sqlStr = sqlStr & " )" & VbCrlf
        sqlStr = sqlStr & " values("& VbCrlf
        sqlStr = sqlStr & " '"& html2DB(ticketPlaceName) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(tPAddress) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(tPTel) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(tPHomeURL) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(placeImg) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(contentsImage1) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(contentsImage2) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(contentsImage3) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(brd_content) & "'" & VbCrlf
        sqlStr = sqlStr & " ,'"& html2DB(parkingGuide) & "'" & VbCrlf
        sqlStr = sqlStr & " )" & VbCrlf
       
        dbget.Execute sqlStr
    end if
else
    response.write "<script>alert('정의 되지 않았습니다. - "&mode&"');history.back();</script>"
    response.end
end if

'// 요청페이지로 이동
IF (mode="ticketInfo") then
    response.redirect("/admin/itemmaster/pop_ticketIteminfo.asp?itemid=" & cStr(itemid))
elseIF (mode="ticketPlace") then
    response.redirect("/admin/itemmaster/pop_TicketPlaceList.asp")
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->