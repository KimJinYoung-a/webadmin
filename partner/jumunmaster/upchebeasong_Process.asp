<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체배송
' Hieditor : 2007.04.07 서동석 생성
'			 2021.09.17 한용민 수정(미출고 알림톡 추가. 자동취소 로직생성.)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim Makerid ,mode ,orderserialArr ,songjangnoArr ,songjangdivArr ,detailidxArr ,MisendReason, ipgodate, detailidx
dim sqlStr,i ,Overlap ,RectdetailidxArr, RectOrderSerialArr, RectSongjangnoArr, RectSongjangdivArr, OrderCount
dim TotAssignedRow, AssignedRow, FailRow, iGLBSongjangDiv, isBlankSDivForceDefaultDivBrand
dim IsMisendReasonInserted, prevStateStr, prevcode, orderserial, preisSendSMS, preisSendEmail
dim itemlackno, chulgodeleygubun, previpgodate, reqaddstr
dim fullText, failText, btnJson, failtitle, oneMisend, BuyHP, Itemname
    iGLBSongjangDiv = CStr(getDefaultSongJangDiv(session("ssBctId")))
    isBlankSDivForceDefaultDivBrand = (LCASE(session("ssBctId"))="visviva") or (LCASE(session("ssBctId"))="roomnhome") or (LCASE(session("ssBctId"))="houseinstyle") or (LCASE(session("ssBctId"))="temp")
    isBlankSDivForceDefaultDivBrand = (isBlankSDivForceDefaultDivBrand AND (iGLBSongjangDiv<>"0"))
	Makerid = session("ssBctID")
	orderserialArr = Replace(request.Form("orderserialArr"), " ", "")
	songjangnoArr  = Replace(request.Form("songjangnoArr"), " ", "")
	songjangdivArr = Replace(request.Form("songjangdivArr"), " ", "")
	detailidxArr   = Replace(request.Form("detailidxArr"), " ", "")
	mode            = requestCheckVar(request.Form("mode"), 32)
	MisendReason    = requestCheckVar(request.Form("MisendReason"), 32)
	ipgodate        = requestCheckVar(request.Form("ipgodate"), 32)
	detailidx       = Replace(request.Form("detailidx"), " ", "")
	reqaddstr       = requestCheckVar(request.Form("reqaddstr"), 2048)
	itemlackno     	= requestCheckVar(request.Form("itemlackno"), 32)

TotAssignedRow = 0
AssignedRow    = 0
FailRow        = 0
preisSendSMS="N"
preisSendEmail="N"

if (mode="SongjangInputCSV") then
    ''CSV 입력은 끝에 , 가 하나 없음. 콤마 사이에 공백 있음
    orderserialArr = Replace(orderserialArr," ","") & ","
    songjangnoArr  = Replace(songjangnoArr," ","") & ","
    songjangdivArr = Replace(songjangdivArr," ","") & ","
    detailidxArr   = Replace(detailidxArr," ","") & ","
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

dim iMailOrderserialArr : iMailOrderserialArr=""
'if (mode="SongjangInputCSV") then
'    response.write "RectdetailidxArr=" & detailidxArr & "<br>"
'    response.write "RectOrderSerialArr=" & orderserialArr & "<br>"
'    response.write "songjangnoArr=" & songjangnoArr & "<br>"
'    response.write "songjangdivArr=" & songjangdivArr & "<br>"
'
'    RectdetailidxArr   = split(detailidxArr,",")
'    RectOrderSerialArr = split(orderserialArr,",")
'    RectSongjangnoArr  = split(songjangnoArr,",")
'    RectSongjangdivArr = split(songjangdivArr,",")
'    OrderCount = Ubound(RectdetailidxArr)
'
'    response.write "OrderCount=" & OrderCount & "<br>"
'    response.write RectdetailidxArr(0)
'end if

'' SongjangInputCSV CSV로 입력 추가
dim mibeasongSoldOutExists
dim psongjangno, psongjangdiv, pcurrstate, newsongjangdiv
dim ret_songjangcheckReturn, ret_songjangcheckErr, songjangcheckErrALL

if (mode="SongjangInput") or (mode="SongjangInputCSV") then
    RectdetailidxArr   = split(detailidxArr,",")
    RectOrderSerialArr = split(orderserialArr,",")
    RectSongjangnoArr  = split(songjangnoArr,",")
    RectSongjangdivArr = split(songjangdivArr,",")

    if IsArray(RectdetailidxArr) then
        OrderCount = Ubound(RectdetailidxArr)

        ''2010-05-26 추가
        if (OrderCount<>Ubound(RectOrderSerialArr)) or (OrderCount<>Ubound(RectSongjangnoArr)) or (OrderCount<>Ubound(RectSongjangdivArr)) then
            response.write "<script type='text/javascript'>alert('전송된 데이터가 일치하지 않습니다.');</script>"
            dbget.close()	:	response.end
        end if

        ''택배사 빈값 체크 필요함..
    end if

    if Right(detailidxArr,1)="," then detailidxArr = Left(detailidxArr,Len(detailidxArr)-1)
    if (Right(orderserialArr,1)=",") then orderserialArr=Left(orderserialArr,Len(orderserialArr)-1)
    orderserialArr = replace(orderserialArr,",","','")

    ''#################################################
    ''송장번호입력 루프
    ''#################################################
    ''2009 출고 소요일 passday 추가.
    for i=0 to OrderCount - 1
        if (Trim(RectdetailidxArr(i))<>"") then
            ''품절출고 불가 등록된경우 SKIP
            mibeasongSoldOutExists = false
            sqlStr = "select count(*) as CNT from db_temp.dbo.tbl_mibeasong_list WITH(NOLOCK)" & VbCRLF
            sqlStr = sqlStr + " where detailidx=" & Trim(RectdetailidxArr(i))  & VbCRLF
            sqlStr = sqlStr + " and orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
            sqlStr = sqlStr + " and code='05'" & VbCRLF
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly
        	if Not rsget.Eof then
                mibeasongSoldOutExists = rsget("CNT")>0
            end if
        	rsget.close

            ''송장 검증 (길이, 특수문자 , DIGIT 코드 등)
            ret_songjangcheckReturn = 0
            ret_songjangcheckErr = ""
            sqlStr = "DECLARE @retERR varchar(50)"&vbCRLF
            sqlStr = sqlStr + " DECLARE @RETURNVALUE int"&vbCRLF
            sqlStr = sqlStr + " exec @RETURNVALUE=[db_order].[dbo].[usp_Ten_Delivery_Trace_VALIDCHECK_Songjang] '"&Trim(RectSongjangdivArr(i))&"', '"&Trim(RectSongjangnoArr(i))&"', '"&Makerid&"','"&Trim(RectOrderSerialArr(i))&"','"&Trim(RectdetailidxArr(i))&"', @retERR output"&vbCRLF
            sqlStr = sqlStr + " select @RETURNVALUE as returnvalue,@retERR as reterr"

            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly
            if Not rsget.Eof then
                ret_songjangcheckReturn = rsget("returnvalue")
                ret_songjangcheckErr    = rsget("reterr")
            end if
        	rsget.close


        	if (mibeasongSoldOutExists) or (ret_songjangcheckReturn<1) then
        	    FailRow = FailRow + 1

                if (ret_songjangcheckReturn<1) and (ret_songjangcheckErr<>"") then
                    songjangcheckErrALL = songjangcheckErrALL & ret_songjangcheckErr &vbCRLF
                end if
        	ELSE

				''response.write i&"="&Trim(RectOrderSerialArr(i))&"<br>"
                ''중복메일 방지용.
                psongjangno = ""
                psongjangdiv= -1
                pcurrstate  = 0

                if (mode="SongjangInputCSV") then
                    sqlStr = "select d.orderserial,isNULL(d.songjangdiv,-1) as psongjangdiv,isNULL(d.songjangno,'') as psongjangno, IsNULL(d.currstate,0) as pcurrstate from [db_order].[dbo].tbl_order_detail d"
                    sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_master m"
                    sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
                    sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
                	sqlStr = sqlStr + " and d.idx =" & Trim(RectdetailidxArr(i))  & VbCRLF
                	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
                	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
                	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
            	    '' sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"                                   ''''' 기 발송 내역 재발송 못하게.. => 아래서 체크
                else
'                    sqlStr = "select d.orderserial " & VbCRLF
'                    sqlStr = sqlStr + "  from [db_order].[dbo].tbl_order_detail d " & VbCRLF
'                    sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
'                    sqlStr = sqlStr + " and d.makerid='" & Makerid & "'" & VbCRLF
'                    sqlStr = sqlStr + " and d.itemid<>0" & VbCRLF
'                    sqlStr = sqlStr + " and d.cancelyn<>'Y'" & VbCRLF
'                    sqlStr = sqlStr + " group by d.orderserial" & VbCRLF
'                    sqlStr = sqlStr + " having count(*)<=sum(CASE WHEN d.currstate=7 then 1 else 0 END)+1" & VbCRLF  ''재입력시 발송 안되게 하려면 < 뺄것.

                    ''송장번호 다른경우만.  ''2013/01/07 수정
                    sqlStr = "select d.orderserial,isNULL(d.songjangdiv,-1) as psongjangdiv,isNULL(d.songjangno,'') as psongjangno, IsNULL(d.currstate,0) as pcurrstate from [db_order].[dbo].tbl_order_detail d"
                    sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_master m"
                    sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
                    sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
                	sqlStr = sqlStr + " and d.idx =" & Trim(RectdetailidxArr(i))  & VbCRLF
                	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
                	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
                	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
                	''sqlStr = sqlStr + " and isNULL(d.songjangno,'')<>'" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF  '' 아래서 체크로 변경.

                end if

            	rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly
            	if Not rsget.Eof then
                    psongjangno = rsget("psongjangno")
                    psongjangdiv= rsget("psongjangdiv")

                    if (Trim(RectSongjangnoArr(i))<>psongjangno) then          '' 송장번호가 변경되는경우만.
                        if ((mode<>"SongjangInputCSV") or ((mode="SongjangInputCSV") and (pcurrstate<7))) then   '' CSV 입력인경우는 최초만.
                            if Not (InStr(iMailOrderserialArr,rsget("orderserial") + ",")>0) then
                                iMailOrderserialArr = iMailOrderserialArr + rsget("orderserial") + ","
                            end if
                        end if
                    end if
            	end if
            	rsget.close

                newsongjangdiv = CHKIIF((Trim(RectSongjangdivArr(i))="" or Trim(RectSongjangdivArr(i))="0") and (isBlankSDivForceDefaultDivBrand),iGLBSongjangDiv,Trim(RectSongjangdivArr(i)))

            	sqlStr = "update D" & VbCRLF
            	sqlStr = sqlStr + " set currstate='7'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangno=convert(varchar(32), '" & Trim(RectSongjangnoArr(i)) & "') " & VbCRLF
            	sqlStr = sqlStr + " ,songjangdiv='" &newsongjangdiv& "'" & VbCRLF  ''2013/10/16 플레이오토 제대로 안넘기는듯 (visviva)
            	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCRLF
            	sqlStr = sqlStr + " ,passday=IsNULL(db_sitemaster.dbo.fn_Ten_NetWorkDays((select convert(varchar(10),baljudate,21) from db_order.dbo.tbl_order_master where orderserial='" & Trim(RectOrderSerialArr(i)) & "'),IsNULL(convert(varchar(10),beasongdate,21),convert(varchar(10),getdate(),21))),0)"& VbCRLF
            	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail D"
            	sqlStr = sqlStr + "     Join [db_order].[dbo].tbl_order_master m"
                sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
            	sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.idx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"   ''완료후 송장번호 변경 할 수 있음.. :: 개별입력만 가능하도록.
                end if

    			''rw sqlStr
                dbget.Execute sqlStr, AssignedRow

                TotAssignedRow = TotAssignedRow + AssignedRow

                if (AssignedRow=0) then FailRow = FailRow + 1

                if ((psongjangno<>"") and (psongjangdiv<>-1)) and ((psongjangno<>Trim(RectSongjangnoArr(i))) or (CStr(psongjangdiv)<>newsongjangdiv)) then
                    ''로그 / 추적 큐 추가 //2019/06/27 by eastone
                    sqlStr = " exec db_order.[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang_AddOnlyLog] "&Trim(RectdetailidxArr(i))&",'"&Trim(RectOrderSerialArr(i))&"','"&psongjangno&"','"&psongjangdiv&"','"&Trim(RectSongjangnoArr(i))&"','"&newsongjangdiv&"','"&session("ssBctId")&"'"
                    dbget.Execute sqlStr

                end if

            END IF
        end if
    next

''rw "iMailOrderserialArr="&iMailOrderserialArr
''rw "orderserialArr="&orderserialArr

    '' ipkumdiv 8 추가
    sqlStr = "update [db_order].[dbo].tbl_order_master" & VbCRLF
    sqlStr = sqlStr + " set  ipkumdiv='7'" & VbCRLF
    '''sqlStr = sqlStr + " , beadaldate=getdate()" & VbCRLF                                '' ipkumdiv='8' 만 beadaldate 입력 ''2013/01/07 수정
    sqlStr = sqlStr + " where orderserial in (" & VbCRLF
    sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m" & VbCRLF
    sqlStr = sqlStr + "         left join [db_order].[dbo].tbl_order_detail d" & VbCRLF
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
    sqlStr = sqlStr + "     where m.orderserial in ('" & orderserialArr & "')" & VbCRLF
    sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
    sqlStr = sqlStr + "     and jumundiv<>9" & VbCRLF
    sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
    sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )>0 and sum(case when IsNull(d.currstate,'0')='7' then 1 else 0 end )>0" & VbCRLF
    sqlStr = sqlStr + " ) "

    ''rw sqlStr
	dbget.Execute sqlStr

    sqlStr = "update [db_order].[dbo].tbl_order_master" & VbCRLF
    sqlStr = sqlStr + " set  ipkumdiv='8'" & VbCRLF
    sqlStr = sqlStr + " , beadaldate=(CASE WHEN ipkumdiv='8' and beadaldate is Not NULL then beadaldate ELSE getdate() END)" & VbCRLF  ''2013/01/07 수정
    sqlStr = sqlStr + " where orderserial in (" & VbCRLF
    sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m" & VbCRLF
    sqlStr = sqlStr + "         left join [db_order].[dbo].tbl_order_detail d" & VbCRLF
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
    sqlStr = sqlStr + "     where m.orderserial in ('" & orderserialArr & "')" & VbCRLF
    sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
    sqlStr = sqlStr + "     and m.jumundiv<>9" & VbCRLF
    sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
    sqlStr = sqlStr + "     and d.cancelyn<>'Y'" & VbCRLF
    sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0"
    sqlStr = sqlStr + " ) "

    ''rw sqlStr
    dbget.Execute sqlStr

    ''-- 미출고 마일리지 업데이트 --2015/03/06
	sqlStr = "insert into db_temp.dbo.tbl_michulgoMile_Recalcu_Que" & VbCRLF
	sqlStr = sqlStr + " (userid)" & VbCRLF
	sqlStr = sqlStr + " select m.userid" & VbCRLF
	sqlStr = sqlStr + " from db_order.dbo.tbl_order_master m" & VbCRLF
	sqlStr = sqlStr + " where m.orderserial in ('" & orderserialArr & "')" & VbCRLF
	sqlStr = sqlStr + " and m.userid<>''" & VbCRLF

	dbget.Execute sqlStr

    ''#################################################
    ''메일보내기 루프
    ''#################################################
    iMailOrderserialArr = split(iMailOrderserialArr,",")

    if IsArray(iMailOrderserialArr) then
        for i=LBound(iMailOrderserialArr) to UBound(iMailOrderserialArr)
            if Trim(iMailOrderserialArr(i))<>"" then
        		On Error resume Next
				'// 즉시 발송되는 대신 [AMAILDB].DB_AMailer.dbo.auto_Mail_Basic_QUE 에 추가된다.
                ''if (application("Svr_Info")<>"Dev") then
                    if (isDlvFinishedByBrand(iMailOrderserialArr(i),MakerID)) then                '''2014/03/31 추가
                        call fcSendMailFinish_Dlv_Designer(iMailOrderserialArr(i),MakerID)

                        '' send Push Msg 2014/03/31
                        sqlStr = "exec db_contents.[dbo].[sp_Ten_sendPushMsg_Deliver] '"&iMailOrderserialArr(i)&"','"&MakerID&"'"
                        dbget.Execute sqlStr
                    end if
                ''end if
                on Error Goto 0
            end if
        next
    end if

    dim AlertMsg
    AlertMsg = TotAssignedRow & "건 처리 되었습니다."
    if (FailRow>0) then
        AlertMsg = AlertMsg & "\n\n(" & FailRow & "건 입력 실패)"

        if (songjangcheckErrALL<>"") then
            AlertMsg = AlertMsg & "\n\n"&replace(songjangcheckErrALL,vbCRLF,"\r\n")
        end if
    end if

    response.write "<script type='text/javascript'>alert('" & AlertMsg & "')</script>"

    if (mode="SongjangInputCSV") then
        response.write "<script type='text/javascript'>opener.location.reload();</script>"
        response.write "<script type='text/javascript'>window.close();</script>"
    else
        response.write "<script type='text/javascript'>location.replace('" + CStr(referer) + "')</script>"
    end if
    dbget.close()	:	response.End

elseif (mode="misendInputOne") then
    ''출고 지연 아니면 ipgodate 널
    dim ckSendSMS, ckSendEmail, ckSendCall, sendState
    dim Sitemid, Sitemoption, itemSoldOut, optSoldOut

    sendState = "2"

    ''관리자경우
    if (C_ADMIN_USER) then
        ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
        ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")
        ckSendCall  = CHKIIF(request("ckSendCall")="on","Y","N")

        if (ckSendCall="Y") then sendState = "4"

        if ((MisendReason="05") or (MisendReason="66") or (MisendReason="06")) then
            ipgodate    = ""
            ckSendSMS   = "N"
            ckSendEmail = "N"
            ckSendCall  = "N"
        else
            sendState = "4"
        end if
    else
        ''업체인경우
        if ((MisendReason="05") or (MisendReason="66") or (MisendReason="06")) then
            ipgodate    = ""
            ckSendSMS   = "N"
            ckSendEmail = "N"
            ckSendCall  = "N"
        else
            sendState = "4"

            ckSendSMS   = "Y"
            ckSendEmail = "Y"
            ckSendCall  = "N"
        end if
    end if

	if ((MisendReason="05") or (MisendReason="66") or (MisendReason="06")) then
		'// CS미처리 전환
		sendState = "0"
		ipgodate = ""
	end if

    if ((MisendReason="05") or (MisendReason="66")) then
        Sitemid     = RequestCheckVar(request("Sitemid"),10)
        Sitemoption = RequestCheckVar(request("Sitemoption"),4)
        itemSoldOut = RequestCheckVar(request("itemSoldOut"),4)

		if Not C_ADMIN_USER then
			'// 업체는 일시품절만 등록가능
			itemSoldOut = "S"
		end if

        if (Sitemid<>"") and (Sitemoption<>"") then
            if (Sitemoption="0000") then
                sqlStr = " update db_item.dbo.tbl_item" & VbCrlf
                sqlStr = sqlStr & " set sellyn='" & itemSoldOut & "'" & VbCrlf
                sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
                sqlStr = sqlStr & " where itemid=" & Sitemid
				sqlStr = sqlStr & " and sellyn = 'Y' "

                dbget.Execute sqlStr
            else
                optSoldOut = "N"

                sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(Sitemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(Sitemoption) + "'"

				dbget.Execute sqlStr

				''옵션갯수
                sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
                sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
                sqlStr = sqlStr + " from (" + VBCrlf
                sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
                sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
                sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
                sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
                sqlStr = sqlStr + " ) T" + VBCrlf
                sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(Sitemid) + VBCrlf

                dbget.Execute sqlStr

                ''상품한정수량
            	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
            	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
            	sqlStr = sqlStr + " from (" + VBCrlf
            	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
            	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
            	sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
            	sqlStr = sqlStr + " ) T" + VBCrlf
            	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"

            	dbget.Execute sqlStr

            	'' 한정 판매 0 이면 일시 품절 처리
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='S'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " " & VbCrlf
            	sqlStr = sqlStr + " and sellyn='Y'"
            	sqlStr = sqlStr + " and limityn='Y'"
            	sqlStr = sqlStr + " and limitno-limitSold<1"

                dbget.Execute sqlStr

            	'' 판매중인 옵션이 없으면 품절처리
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='N'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and optioncnt=0"

                dbget.Execute sqlStr

            end if
        end if
    end if

	sqlStr = "select top 1 orderserial, itemname, IsNull(itemoptionname, '') as itemoptionname, code, IsNull(isSendSms, 'N') as isSendSms"
    sqlStr = sqlStr & " , IsNull(isSendEmail, 'N') as isSendEmail, IsNull(isSendCall, '') as isSendCall, isnull(convert(varchar(10),ipgodate,121),'') as ipgodate"
	sqlStr = sqlStr & " from [db_temp].dbo.tbl_mibeasong_list with (nolock) where detailidx=" & CStr(detailidx) & " "

    'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly

	IsMisendReasonInserted = False
	if Not rsget.Eof then
		IsMisendReasonInserted = True
		prevcode = rsget("code")
        previpgodate = rsget("ipgodate")
        preisSendSMS = rsget("isSendSMS")
        preisSendEmail = rsget("isSendEmail")
		orderserial = rsget("orderserial")

		prevStateStr = "기존 미출고사유" + vbCrLf
		prevStateStr = prevStateStr + "상품명 : " + CStr(rsget("itemname"))
		prevStateStr = prevStateStr + "[" + CStr(rsget("itemoptionname")) + "]" + vbCrLf
		prevStateStr = prevStateStr + "미출고사유 : " + MiSendCodeToName(rsget("code")) + vbCrLf
		prevStateStr = prevStateStr + "고객알림 : SMS(" + CStr(rsget("isSendSms")) + "), 이메일(" + CStr(rsget("isSendEmail")) + "), 통화(" + CStr(rsget("isSendCall")) + ")" + vbCrLf
		prevStateStr = prevStateStr + "처리예정일 : " + CStr(rsget("ipgodate"))
	end if
	rsget.close

	if IsMisendReasonInserted = True then

		if (prevcode <> MisendReason) then
			'// 사유변경시 CS메모 등록
			Call AddCsMemo(orderserial, "1", "", session("ssBctId"), prevStateStr)
		end if

		sqlStr = sqlStr + " update [db_temp].dbo.tbl_mibeasong_list"
		sqlStr = sqlStr + " set code='" & MisendReason & "' "

		if (prevcode <> MisendReason) then
			sqlStr = sqlStr + " , prevcode = '" + CStr(prevcode) + "' "
		end if

		sqlStr = sqlStr + " ,state='"&sendState&"'"                                         ''상태 변경 (기존 안내메일완료)
        if MisendReason = "05" or MisendReason = "06" then
            if prevcode <> MisendReason then
                '// 품절 등록되면 품절알림 문자발송, 2020-02-13, skyer9
                sqlStr = sqlStr + " ,isSendSMS='N' "
                sqlStr = sqlStr + " ,isSendEmail='N'"
                sqlStr = sqlStr + "	,sendCount=0 "
            end if
        else
			sqlStr = sqlStr + " ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS발송완료
			sqlStr = sqlStr + " ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email발송완료
			'''sqlStr = sqlStr + " ,isSendCall=(CASE WHEN isSendCall='Y' then 'Y' ELSE '"&ckSendCall&"' END)"  '' CALL완료 : 따로 처리
		    if (ckSendSMS = "Y") or (ckSendEmail = "Y") then
			    sqlStr = sqlStr + "	,sendCount=IsNull(sendCount,0) + 1 "
			    sqlStr = sqlStr + "	,lastSendUserid='" + CStr(session("ssBctId")) + "' "
			    sqlStr = sqlStr + "	,lastSendDate=getdate() "
		    end if
        end if

		if (ipgodate<>"") then
			sqlStr = sqlStr + "	,ipgodate='" & ipgodate & "'"
		else
			sqlStr = sqlStr + "	,ipgodate=NULL"
		end if
		sqlStr = sqlStr + "	,reqaddstr = '" & html2db(reqaddstr) & "' "
		sqlStr = sqlStr + "	,modiuserid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + "	,modidate = getdate() "
		sqlStr = sqlStr + " where detailidx=" & detailidx
	else
		sqlStr = sqlStr + "	    insert into [db_temp].dbo.tbl_mibeasong_list"
		sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
		sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
		if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
			sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"             ''상태 변경 (기존 안내메일완료)
			''sqlStr = sqlStr + "	isSendCall,"

			if (ckSendSMS = "Y") or (ckSendEmail = "Y") then
				sqlStr = sqlStr + "	sendCount, lastSendUserid, lastSendDate, "
			end if
		end if
		sqlStr = sqlStr + "	    itemname, itemoptionname,reqaddstr, reguserid)"
		sqlStr = sqlStr + "	    select idx, orderserial, itemid,itemoption,"
		sqlStr = sqlStr + "	    itemno, itemno, '" & MisendReason & "',"

		if (ipgodate<>"") then
			sqlStr = sqlStr + "	'" & ipgodate & "','',"
		else
			sqlStr = sqlStr + "	NULL,'',"
		end if
		if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
			sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
			''sqlStr = sqlStr + "	 '"&ckSendCall&"',"

			if (ckSendSMS = "Y") or (ckSendEmail = "Y") then
				sqlStr = sqlStr + "	1, '" + CStr(session("ssBctId")) + "', getdate(), "
			end if
		end if
		sqlStr = sqlStr + "	    itemname, itemoptionname, '" & html2db(reqaddstr) & "', '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + "	    from [db_order].[dbo].tbl_order_detail"
		sqlStr = sqlStr + "	    where idx=" & detailidx
	end if

	''rw   sqlStr
	dbget.Execute sqlStr

	if ((MisendReason="05") or (MisendReason="66") or (MisendReason="06")) and itemlackno <> "" then
		if Not IsNumeric(itemlackno) then
			itemlackno = "0"
		end if

		sqlStr = " update [db_temp].dbo.tbl_mibeasong_list "
		sqlStr = sqlStr + " set itemlackno = " & itemlackno
		sqlStr = sqlStr + " where detailidx = " & detailidx
		sqlStr = sqlStr + " and itemno >= " & itemlackno
		sqlStr = sqlStr + " and 0 < " & itemlackno
		dbget.Execute sqlStr
	end if


	dim tmp_sendsmsmsg, tmp_sendmailmsg

	if ((MisendReason <> "05") and (MisendReason <> "66") and (MisendReason <> "06")) then
		tmp_sendsmsmsg = GetMichulgoSMSString(MisendReason)
		tmp_sendmailmsg = GetMichulgoMailString(MisendReason)

		tmp_sendmailmsg = Replace(tmp_sendmailmsg, "\n", "<br>")

		tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[출고예정일]", ipgodate)
	end if

    ''SMS 발송 + [CS메모에 저장 -> 같이 되어있음.]
    if (ckSendSMS="Y") then
        'if (application("Svr_Info")<>"Dev") then
            ' 출고지연. 카카오톡 알림톡 발송.   ' 2021.09.17 한용민 생성
            if MisendReason = "03" then
                set oneMisend = new COldMiSend
                    oneMisend.FRectDetailIDx = detailidx
                    oneMisend.getOneOldMisendItem

                buyhp = oneMisend.FOneItem.FBuyHP
                Itemname = replace(oneMisend.FOneItem.FItemname,vbcrlf,"")

                if buyhp<>"" and not(isnull(buyhp)) then
                    chulgodeleygubun=""
                    sqlStr = "select"
                    sqlStr = sqlStr & " l.idx"
                    sqlStr = sqlStr & " , (case"
                    sqlStr = sqlStr & "     when isnull(l.prevcode,'00')='05' and '"& prevcode &"'<>'03' and convert(varchar(10),ipgodate,121)<>'"& previpgodate &"' then '05_03'"   ' 품절출고불가 상품 출고지연전환 케이스. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
                    sqlStr = sqlStr & "     when isnull(l.prevcode,'00')<>'03' and '"& MisendReason &"'='03' and convert(varchar(10),ipgodate,121)<>'"& previpgodate &"' then '03'"    ' 출고지연 알림톡. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
                    sqlStr = sqlStr & "     when '"& prevcode &"'<>'03' and '"& MisendReason &"'='03' then '03'"   ' 다른사유를 출고지연전환 케이스. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
                    sqlStr = sqlStr & "     when '"& prevcode &"'='"& MisendReason &"' and '"& preisSendSMS &"'='N' and '"& preisSendEmail &"'='N' then '03'"	' 사유는 같으나 알림 발송을 위해 재저장 버튼을 누른 케이스
                    sqlStr = sqlStr & "     else '' end) as chulgodeleygubun"
                    sqlStr = sqlStr & " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
                    sqlStr = sqlStr & " where l.code = '03'"	' 출고지연
                    sqlStr = sqlStr & " and l.ipgodate is not null"
                    sqlStr = sqlStr & " and l.detailidx="& detailidx &""

                    'response.write sqlStr & "<br>"
                    rsget.CursorLocation = adUseClient
                    rsget.Open sqlStr, dbget, adOpenForwardOnly
                    if Not rsget.Eof then
                        chulgodeleygubun = rsget("chulgodeleygubun")
                    end if
                    rsget.close

                    ' 품절출고불가 상품 출고지연전환 케이스
                    if chulgodeleygubun="05_03" then
                        failtitle = "[텐바이텐]상품출고 안내"
                        fullText = "[10x10] 상품출고 안내" & vbCrLf & vbCrLf
                        fullText = fullText & "품절취소 안내드렸던 상품의 재고가 확보되어 발송 예정으로, 아래의 예정일까지 출발할 수 있도록 최선의 노력을 다하겠습니다." & vbCrLf & vbCrLf & vbCrLf
                        fullText = fullText & "■ 주문번호 : "& oneMisend.FOneItem.Forderserial &"" & vbCrLf
                        fullText = fullText & "■ 상품명 : "& Itemname &"" & vbCrLf
                        fullText = fullText & "■ 출발예정일 : "& ipgodate &"" & vbCrLf & vbCrLf
                        fullText = fullText & "감사합니다."
                        failText = fullText
                        btnJson = "{""button"":[{""name"":""주문내역 바로가기"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/L1izHiDBdjb""}]}"
                        call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0024",fullText,"LMS",failtitle,failText,btnJson,"",oneMisend.FOneItem.Fuserid)

                        sqlStr = "update db_temp.dbo.tbl_mibeasong_list set finishstr=N'품절상품 출고지연전환 알림톡 발송완료' where detailidx="& detailidx &"" & vbcrlf
		                dbget.Execute sqlStr

                        Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[품절상품 출고지연전환 알림톡 발송완료 "+ buyhp +"]" + html2db(fullText))

                    ' 출고지연 알림톡
                    elseif chulgodeleygubun="03" then
                        failtitle = "[텐바이텐]출고 지연 안내"
                        fullText = "[10x10] 출고 지연 안내" & vbCrLf & vbCrLf
                        fullText = fullText & "출고지연으로 양해의 말씀 드립니다." & vbCrLf
                        fullText = fullText & "주문하신 소중한 상품의 배송지연이 예상되오며," & vbCrLf
                        fullText = fullText & "아래의 예정일까지 출발할 수 있도록 최선의 노력을 다하겠습니다." & vbCrLf & vbCrLf
                        fullText = fullText & "■ 주문번호 : "& oneMisend.FOneItem.Forderserial &"" & vbCrLf
                        fullText = fullText & "■ 상품명 : "& Itemname &"" & vbCrLf
                        fullText = fullText & "■ 출발예정일 : "& ipgodate &"" & vbCrLf & vbCrLf
                        fullText = fullText & "감사합니다."
                        failText = fullText
                        btnJson = ""
                        call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0009",fullText,"LMS",failtitle,failText,btnJson,oneMisend.FOneItem.Forderserial,oneMisend.FOneItem.Fuserid)

                        sqlStr = "update db_temp.dbo.tbl_mibeasong_list set finishstr=N'출고지연 알림톡 발송완료' where detailidx="& detailidx &"" & vbcrlf
		                dbget.Execute sqlStr

                        Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송완료 "+ buyhp +"]" + html2db(fullText))
                    else
                        ' 오류시노티
                        Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송실패.미출고사유코드:"& chulgodeleygubun &".이전사유코드:"& prevcode &".이전입고예정일:"&previpgodate&"]")
                    end if
                else
                    ' 오류시노티
                    Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송실패.주문자휴대폰번호:"& buyhp &"]")
                end if
                set oneMisend = nothing

            ' 문자 발송
            else
			    Call SendMiChulgoSMSWithMessage(detailidx, tmp_sendsmsmsg)
            end if
        'end if
	end if

    ''EMail발송
    if (ckSendEmail="Y") then
        if (application("Svr_Info")<>"Dev") then
            ''call SendMiChulgoMail(detailidx)
			Call SendMiChulgoMailWithMessage(detailidx, tmp_sendmailmsg)
        end if
    end if

	if (MisendReason="05") then
        '// 품절출고불가 담당자 지정
		sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " & detailidx & " "
		dbget.Execute sqlStr
    end if

    if (ckSendSMS="Y") and (ckSendEmail="Y") then
        response.write "<script type='text/javascript'>alert('SMS와 메일이 발송 되었습니다.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script type='text/javascript'>alert('SMS가 발송 되었습니다.');</script>"
    elseif (ckSendEmail="Y") then
        response.write "<script type='text/javascript'>alert('메일이 발송 되었습니다.');</script>"
    else
        response.write "<script type='text/javascript'>alert('처리 되었습니다.');</script>"
    end if
    response.write "<script type='text/javascript'>opener.location.reload();</script>"
    response.write "<script type='text/javascript'>location.replace('" + CStr(referer) + "')</script>"
    dbget.close()	:	response.End
end if

Function getDefaultSongJangDiv(iMakerid)
    dim sqlStr, ret
    ret = 0
    sqlstr = " select top 1 IsNULL(defaultsongjangdiv,0) as defaultsongjangdiv from db_partner.dbo.tbl_partner where id='"&iMakerid&"'"

    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF Not (rsget.EOF OR rsget.BOF) THEN
    	ret = rsget("defaultsongjangdiv")
    END IF
    getDefaultSongJangDiv = ret
    rsget.Close
end function

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
