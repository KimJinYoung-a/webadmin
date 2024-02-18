<%@ language=vbscript %>
<%
option explicit
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim songjangnoArr , MisendReason, ipgodate, detailidx ,songjangdivArr ,detailidxArr
dim Overlap ,sqlStr,i ,ckSendSMS, ckSendEmail, ckSendCall, sendState
dim Sitemid, Sitemoption, itemSoldOut, optSoldOut , mode ,Makerid ,iMailmasteridxArr
Dim RectdetailidxArr, RectordernoArr, RectSongjangnoArr, RectSongjangdivArr, OrderCount
dim TotAssignedRow, AssignedRow, FailRow ,ordernoArr
	ordernoArr = request.Form("ordernoArr")
	songjangnoArr  = request.Form("songjangnoArr")
	songjangdivArr = request.Form("songjangdivArr")
	detailidxArr   = request.Form("detailidxArr")
	mode           = request.Form("mode")
	MisendReason   = request.Form("MisendReason")
	ipgodate       = request.Form("ipgodate")
	detailidx      = request.Form("detailidx")
	Makerid = session("ssBctID")
	iMailmasteridxArr=""

if (mode="SongjangInputCSV") then
    ''CSV 입력은 끝에 , 가 하나 없음. 콤마 사이에 공백 있음
    ordernoArr = Replace(ordernoArr," ","") & ","
    songjangnoArr  = Replace(songjangnoArr," ","") & ","
    songjangdivArr = Replace(songjangdivArr," ","") & ","
    detailidxArr   = Replace(detailidxArr," ","") & ","
end if

TotAssignedRow = 0
AssignedRow    = 0
FailRow        = 0

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

'' SongjangInputCSV CSV로 입력 추가
dim mibeasongSoldOutExists

if (mode="SongjangInput") or (mode="SongjangInputCSV") then
    RectdetailidxArr   = split(detailidxArr,",")
    RectordernoArr = split(ordernoArr,",")
    RectSongjangnoArr  = split(songjangnoArr,",")
    RectSongjangdivArr = split(songjangdivArr,",")

    if IsArray(RectdetailidxArr) then
        OrderCount = Ubound(RectdetailidxArr)

        if (OrderCount<>Ubound(RectordernoArr)) or (OrderCount<>Ubound(RectSongjangnoArr)) or (OrderCount<>Ubound(RectSongjangdivArr)) then
            response.write "<script>alert('전송된 데이터가 일치하지 않습니다.');</script>"
            dbget.close()	:	response.end
        end if
    end if

    if Right(detailidxArr,1)="," then detailidxArr = Left(detailidxArr,Len(detailidxArr)-1)
    if (Right(ordernoArr,1)=",") then ordernoArr=Left(ordernoArr,Len(ordernoArr)-1)
    ordernoArr = replace(ordernoArr,",","','")

    dim tmp
    dbget.beginTrans

    ''송장번호입력 루프
    for i=0 to OrderCount - 1
        if (Trim(RectdetailidxArr(i))<>"") then

            ''품절출고 불가 등록된경우 SKIP
            mibeasongSoldOutExists = false

            'sqlStr = "select count(*) as CNT" & VbCRLF
            'sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_mibeasong_list" & VbCRLF
            'sqlStr = sqlStr + " where detailidx=" & Trim(RectdetailidxArr(i))  & VbCRLF
            'sqlStr = sqlStr + " and orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            'sqlStr = sqlStr + " and code='05'" & VbCRLF

            'response.write sqlStr &"<br>"
            'rsget.CursorLocation = adUseClient
            'rsget.Open sqlStr, dbget, adOpenForwardOnly

        	'if Not rsget.Eof then
            '    mibeasongSoldOutExists = rsget("CNT")>0
            'end if

        	'rsget.close

        	if (mibeasongSoldOutExists) then
        	    FailRow = FailRow + 1
        	ELSE

                ''중복메일 방지용.
                sqlStr = ""
                sqlStr = "select d.masteridx"
                sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d"
                sqlStr = sqlStr + " Join db_shop.dbo.tbl_shopbeasong_order_master m"
                sqlStr = sqlStr + " on d.masteridx=m.masteridx"
                sqlStr = sqlStr + " where d.orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.

            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"
                end if

            	'response.write sqlStr &"<br>"
            	rsget.CursorLocation = adUseClient
                rsget.Open sqlStr, dbget, adOpenForwardOnly

            	if Not rsget.Eof then
            		tmp = ""
            		tmp = rsget("masteridx")&","

            	    if Not (InStr(iMailmasteridxArr,tmp)>0) then
            	        iMailmasteridxArr = iMailmasteridxArr + tmp
            	    end if
            	    tmp = ""
            	end if

            	rsget.close

                sqlStr = ""
            	sqlStr = "update D" & VbCRLF
            	sqlStr = sqlStr + " set currstate='7'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangno='" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangdiv='" & Trim(RectSongjangdivArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCRLF
            	sqlStr = sqlStr + " ,passday=IsNULL(db_sitemaster.dbo.fn_Ten_NetWorkDays(("
            	sqlStr = sqlStr + " 	select convert(varchar(10),baljudate,21)"
				sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master mm"
            	sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail dd"
            	sqlStr = sqlStr + " 	on mm.masteridx = dd.masteridx"
            	sqlStr = sqlStr + "		where dd.detailidx=" & Trim(RectdetailidxArr(i)) & ""
            	sqlStr = sqlStr + " 	),IsNULL(convert(varchar(10),d.beasongdate,21),convert(varchar(10),getdate(),21))),0)"& VbCRLF
                sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d"& VbCRLF
            	sqlStr = sqlStr + " Join db_shop.dbo.tbl_shopbeasong_order_master m"
                sqlStr = sqlStr + " on m.masteridx=d.masteridx"
            	sqlStr = sqlStr + " where d.orderno='" & Trim(RectordernoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.

            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"   ''완료후 송장번호 변경 할 수 있음.. :: 개별입력만 가능하도록.
                end if

				'response.write sqlStr &"<br>"
                dbget.Execute sqlStr, AssignedRow

                TotAssignedRow = TotAssignedRow + AssignedRow

                if (AssignedRow=0) then FailRow = FailRow + 1
            END IF
        end if

    next

	'마스터 일부출고 설정
    sqlStr = " update 																					" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master 									" & VbCRLF
    sqlStr = sqlStr + " set 																			" & VbCRLF
    sqlStr = sqlStr + " 	ipkumdiv='7' 																" & VbCRLF
    sqlStr = sqlStr + " 	, beadaldate=getdate() 														" & VbCRLF
    sqlStr = sqlStr + " where 																			" & VbCRLF
    sqlStr = sqlStr + " 	masteridx in ( 																" & VbCRLF
    sqlStr = sqlStr + " 		select 																	" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		from 																	" & VbCRLF
    sqlStr = sqlStr + " 			db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
    sqlStr = sqlStr + " 			join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
    sqlStr = sqlStr + " 			on 																	" & VbCRLF
    sqlStr = sqlStr + " 				m.masteridx=d.masteridx 										" & VbCRLF
    sqlStr = sqlStr + " 		where 																	" & VbCRLF
    sqlStr = sqlStr + " 			1 = 1 																" & VbCRLF
    sqlStr = sqlStr + " 			and d.itemid<>0 													" & VbCRLF
    sqlStr = sqlStr + " 			and m.masteridx in ( 												" & VbCRLF
    sqlStr = sqlStr + " 				select distinct 												" & VbCRLF
    sqlStr = sqlStr + " 					m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 				from 															" & VbCRLF
    sqlStr = sqlStr + " 					db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 					join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 					on 															" & VbCRLF
    sqlStr = sqlStr + " 						m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " 				where 															" & VbCRLF
    sqlStr = sqlStr + " 					1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 					and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 					and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 					and d.itemid<>0 											" & VbCRLF
    sqlStr = sqlStr + " 			) 																	" & VbCRLF
    sqlStr = sqlStr + " 		group by 																" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )>0 " & VbCRLF
    sqlStr = sqlStr + " 	) 																			" & VbCRLF

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

	'전부출고
    sqlStr = " update 																					" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master 									" & VbCRLF
    sqlStr = sqlStr + " set 																			" & VbCRLF
    sqlStr = sqlStr + " 	ipkumdiv='8' 																" & VbCRLF
    sqlStr = sqlStr + " 	, beadaldate=getdate() 														" & VbCRLF
	sqlStr = sqlStr + " where 																			" & VbCRLF
    sqlStr = sqlStr + " 	masteridx in ( 																" & VbCRLF
    sqlStr = sqlStr + " 		select 																	" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		from 																	" & VbCRLF
    sqlStr = sqlStr + " 			db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
    sqlStr = sqlStr + " 			join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
    sqlStr = sqlStr + " 			on 																	" & VbCRLF
    sqlStr = sqlStr + " 				m.masteridx=d.masteridx 										" & VbCRLF
    sqlStr = sqlStr + " 		where 																	" & VbCRLF
    sqlStr = sqlStr + " 			1 = 1 																" & VbCRLF
    sqlStr = sqlStr + " 			and d.itemid<>0 													" & VbCRLF
    sqlStr = sqlStr + " 			and m.masteridx in ( 												" & VbCRLF
    sqlStr = sqlStr + " 				select distinct 												" & VbCRLF
    sqlStr = sqlStr + " 					m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 				from 															" & VbCRLF
    sqlStr = sqlStr + " 					db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 					join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 					on 															" & VbCRLF
    sqlStr = sqlStr + " 						m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " 				where 															" & VbCRLF
    sqlStr = sqlStr + " 					1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 					and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 					and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 					and d.itemid<>0 											" & VbCRLF
    sqlStr = sqlStr + " 			) 																	" & VbCRLF
    sqlStr = sqlStr + " 		group by 																" & VbCRLF
    sqlStr = sqlStr + " 			m.masteridx 														" & VbCRLF
    sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
    sqlStr = sqlStr + " 	) 																			" & VbCRLF

    'response.write sqlStr &"<br>"
    dbget.Execute sqlStr

    ''메일보내기 루프
    iMailmasteridxArr = split(iMailmasteridxArr,",")

    if IsArray(iMailmasteridxArr) then
        for i=LBound(iMailmasteridxArr) to UBound(iMailmasteridxArr)

            if Trim(iMailmasteridxArr(i))<>"" then
                if (application("Svr_Info")<>"Dev") then
                    'call fcSendMailFinish_Dlv_Designer_off(iMailmasteridxArr(i),MakerID)
                end if
            end if
        next
    end if



	'문자발송
	dim reqhparr
	songjangdivarr = ""
	songjangnoarr = ""

    sqlStr = " select distinct 															" & VbCRLF
    sqlStr = sqlStr + " 	m.masteridx 												" & VbCRLF
    sqlStr = sqlStr + " 	, m.reqhp 													" & VbCRLF
    sqlStr = sqlStr + " 	, d.songjangdiv 											" & VbCRLF
    sqlStr = sqlStr + " 	, d.songjangno 												" & VbCRLF
    sqlStr = sqlStr + " from 															" & VbCRLF
    sqlStr = sqlStr + " 	db_shop.dbo.tbl_shopbeasong_order_master m 					" & VbCRLF
    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 			" & VbCRLF
    sqlStr = sqlStr + " 	on 															" & VbCRLF
    sqlStr = sqlStr + " 		m.masteridx=d.masteridx 								" & VbCRLF
    sqlStr = sqlStr + " where 															" & VbCRLF
    sqlStr = sqlStr + " 	1 = 1 														" & VbCRLF
    sqlStr = sqlStr + " 	and d.detailidx in (" & detailidxArr & ") 					" & VbCRLF
    sqlStr = sqlStr + " 	and m.cancelyn='N' 											" & VbCRLF
    sqlStr = sqlStr + " 	and d.itemid<>0 											" & VbCRLF

	rsget.open sqlStr ,dbget ,1

	if not(rsget.eof) then
		do until rsget.Eof
			reqhparr 		= reqhparr + "," + rsget("reqhp")
			songjangdivarr 	= songjangdivarr + "," + CStr(rsget("songjangdiv"))
			songjangnoarr	= songjangnoarr + "," + CStr(rsget("songjangno"))
			rsget.MoveNext
		loop
	end if
	rsget.close()

	if replace(reqhparr,"-","")<>"" then
	    reqhparr = split(reqhparr,",")
	    songjangdivarr = split(songjangdivarr,",")
	    songjangnoarr = split(songjangnoarr,",")

	    if IsArray(reqhparr) then
	        for i=LBound(reqhparr) to UBound(reqhparr)
	            if Trim(reqhparr(i))<>"" then
	                if (application("Svr_Info")<>"Dev") then
	                    'call SendNormalSMS(Trim(reqhparr(i)), "", "[텐바이텐샵] 상품이 출고되었습니다. [" & DeliverDivCd2Nm(Trim(songjangdivarr(i))) & "]" & Trim(songjangnoarr(i)) & "")
	                end if
	            end if
	        next
	    end if
	end if

	If Err.Number = 0 Then
	    dbget.CommitTrans
	Else
	    dbget.RollBackTrans
	End If

    dim AlertMsg
    AlertMsg = TotAssignedRow & "건 처리 되었습니다."
    if (FailRow>0) then
        AlertMsg = AlertMsg & "\n\n(" & FailRow & "건 입력 실패)"
    end if

    response.write "<script language='javascript'>alert('" & AlertMsg & "')</script>"

    if (mode="SongjangInputCSV") then
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>window.close();</script>"
    else
        response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
    end if
    dbget.close()	:	response.End

elseif mode="misendInputOne" then

    ''출고 지연 아니면 ipgodate 널
    sendState = "2"

    ''관리자경우
    if (C_ADMIN_USER) then
        ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
        ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")
        ckSendCall  = CHKIIF(request("ckSendCall")="on","Y","N")

        if (ckSendCall="Y") then sendState = "4"

        '//품절출고 불가
        if (MisendReason="05") then
            ipgodate    = ""
            ckSendSMS   = "N"
            ckSendEmail = "N"
            ckSendCall  = "N"
        else
            sendState = "4"
        end if
    else
        ''품절출고 불가 업체인경우
        if (MisendReason="05") then
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

    '/품절출고 불가
    if (MisendReason="05") then
        Sitemid     = RequestCheckVar(request("Sitemid"),10)
        Sitemoption = RequestCheckVar(request("Sitemoption"),4)
        itemSoldOut = RequestCheckVar(request("itemSoldOut"),4)

        if (Sitemid<>"") and (Sitemoption<>"") then
            if (Sitemoption="0000") then
                sqlStr = " update db_item.dbo.tbl_item" & VbCrlf
                sqlStr = sqlStr & " set sellyn='" & itemSoldOut & "'" & VbCrlf
                sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
                sqlStr = sqlStr & " where itemid=" & Sitemid

                'dbget.Execute sqlStr
            else
                optSoldOut = "N"

                sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(Sitemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(Sitemoption) + "'"

				'dbget.Execute sqlStr

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

               ' dbget.Execute sqlStr

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

            	'dbget.Execute sqlStr

            	'' 한정 판매 0 이면 일시 품절 처리
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='S'"
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and sellyn='Y'"
            	sqlStr = sqlStr + " and limityn='Y'"
            	sqlStr = sqlStr + " and limitno-limitSold<1"

                'dbget.Execute sqlStr

            	'' 판매중인 옵션이 없으면 품절처리
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='N'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and optioncnt=0"

                'dbget.Execute sqlStr

            end if
        end if
    end if

    sqlStr = " IF Exists(select mibeaidx from [db_shop].dbo.tbl_shopbeasong_mibeasong_list where detailidx=" & detailidx & ")"
    sqlStr = sqlStr + " BEGIN "
    sqlStr = sqlStr + "	    update [db_shop].dbo.tbl_shopbeasong_mibeasong_list"
    sqlStr = sqlStr + "	    set code='" & MisendReason & "'"

    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	    ,state='"&sendState&"'"                                         ''상태 변경 (기존 안내메일완료)
        sqlStr = sqlStr + "	    ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS발송완료
        sqlStr = sqlStr + "	    ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email발송완료
    end if
    if (ipgodate<>"") then
        sqlStr = sqlStr + "	,ipgodate='" & ipgodate & "'"
    else
        sqlStr = sqlStr + "	,ipgodate=NULL"
    end if

    sqlStr = sqlStr + "	    where detailidx=" & detailidx
    sqlStr = sqlStr + " END "
    sqlStr = sqlStr + " ELSE "
    sqlStr = sqlStr + " BEGIN "
    sqlStr = sqlStr + "	    insert into [db_shop].dbo.tbl_shopbeasong_mibeasong_list"
    sqlStr = sqlStr + "	    (detailidx, orderno, itemid, itemoption,"
    sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "

    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"             ''상태 변경 (기존 안내메일완료)
    end if

    sqlStr = sqlStr + "	    itemname, itemoptionname)"
    sqlStr = sqlStr + "	    select d.detailidx, d.orderno, d.itemid,d.itemoption,"
    sqlStr = sqlStr + "	    d.itemno, d.itemno, '" & MisendReason & "',"

    if (ipgodate<>"") then
        sqlStr = sqlStr + "	'" & ipgodate & "','',"
    else
        sqlStr = sqlStr + "	NULL,'',"
    end if
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
    end if

    sqlStr = sqlStr + "	od.itemname, od.itemoptionname"
	sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_detail d" + vbcrlf
	sqlStr = sqlStr + "	left join [db_shop].[dbo].tbl_shopjumun_detail od" +vbcrlf
	sqlStr = sqlStr + "	on d.orgdetailidx = od.idx" +vbcrlf
    sqlStr = sqlStr + "	where detailidx=" & detailidx
    sqlStr = sqlStr + " END "

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	    ''SMS 발송 + [CS메모에 저장 -> 같이 되어있음.]
	    if (ckSendSMS="Y") then
	    	response.write "이메일발송Y<br>"
	        if (application("Svr_Info")<>"Dev") then
	            '//당분간 발송안함
	            call SendMiChulgoSMS_off(detailidx)
	        end if
	    end if
	    ''EMail발송
	    if (ckSendEmail="Y") then
	    	response.write "문자발송Y<br>"
	        if (application("Svr_Info")<>"Dev") then
	            '//당분간 발송안함
	            call SendMiChulgoMail_off(detailidx)
	        end if
	    end if

    if (ckSendSMS="Y") and (ckSendEmail="Y") then
        response.write "<script language='javascript'>alert('SMS와 메일이 발송 되었습니다.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script language='javascript'>alert('SMS가 발송 되었습니다.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script language='javascript'>alert('메일이 발송 되었습니다.');</script>"
    else
        response.write "<script language='javascript'>alert('처리 되었습니다.');</script>"
    end if
    response.write "<script language='javascript'>opener.location.reload();</script>"
    response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->