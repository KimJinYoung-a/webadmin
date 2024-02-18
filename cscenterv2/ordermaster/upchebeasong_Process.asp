<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/smsLib.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/maillib.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lectureadmin/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/misendcls.asp"-->
<%

'개발서버인 경우 메일/SMS 발송이 안되도록 되어 있다.
dim SENDMAIL_YN
SENDMAIL_YN = "N"		'Y 인 경우 개발서버에서도 이메일을 발송하게 한다.

dim Makerid
Makerid = session("ssBctID")

dim mode
dim orderserialArr
dim songjangnoArr
dim songjangdivArr
dim detailidxArr
dim MisendReason, ipgodate, detailidx

orderserialArr = request.Form("orderserialArr")
songjangnoArr  = request.Form("songjangnoArr")
songjangdivArr = request.Form("songjangdivArr")
detailidxArr   = request.Form("detailidxArr")
mode           = RequestCheckvar(request.Form("mode"),16)
MisendReason   = RequestCheckvar(request.Form("MisendReason"),2)
ipgodate       = RequestCheckvar(request.Form("ipgodate"),10)
detailidx      = RequestCheckvar(request.Form("detailidx"),10)

if orderserialArr <> "" then
	if checkNotValidHTML(orderserialArr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if songjangnoArr <> "" then
	if checkNotValidHTML(songjangnoArr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if songjangdivArr <> "" then
	if checkNotValidHTML(songjangdivArr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if detailidxArr <> "" then
	if checkNotValidHTML(detailidxArr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if (mode <> "misendInputOne") then
	'// 다른 부분 사용안함
	response.end
end if



if (mode="SongjangInputCSV") then
    ''CSV 입력은 끝에 , 가 하나 없음. 콤마 사이에 공백 있음
    orderserialArr = Replace(orderserialArr," ","") & ","
    songjangnoArr  = Replace(songjangnoArr," ","") & ","
    songjangdivArr = Replace(songjangdivArr," ","") & ","
    detailidxArr   = Replace(detailidxArr," ","") & ","

end if

dim sqlStr,i
dim Overlap
Dim RectdetailidxArr, RectOrderSerialArr, RectSongjangnoArr, RectSongjangdivArr, OrderCount

dim TotAssignedRow, AssignedRow, FailRow
TotAssignedRow = 0
AssignedRow    = 0
FailRow        = 0
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

if (mode="SongjangInput") or (mode="SongjangInputCSV") then
    RectdetailidxArr   = split(detailidxArr,",")
    RectOrderSerialArr = split(orderserialArr,",")
    RectSongjangnoArr  = split(songjangnoArr,",")
    RectSongjangdivArr = split(songjangdivArr,",")

    if IsArray(RectdetailidxArr) then
        OrderCount = Ubound(RectdetailidxArr)

        ''2010-05-26 추가
        if (OrderCount<>Ubound(RectOrderSerialArr)) or (OrderCount<>Ubound(RectSongjangnoArr)) or (OrderCount<>Ubound(RectSongjangdivArr)) then
            response.write "<script>alert('전송된 데이터가 일치하지 않습니다.');</script>"
            dbACADEMYget.close()	:	response.end
        end if
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
            sqlStr = "select count(*) as CNT from db_academy.dbo.tbl_academy_mibeasong_list" & VbCRLF
            sqlStr = sqlStr + " where detailidx=" & Trim(RectdetailidxArr(i))  & VbCRLF
            sqlStr = sqlStr + " and orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
            sqlStr = sqlStr + " and code='05'" & VbCRLF
            rsACADEMYget.CursorLocation = adUseClient
            rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly
        	if Not rsACADEMYget.Eof then
                mibeasongSoldOutExists = rsACADEMYget("CNT")>0
            end if
        	rsACADEMYget.close

        	if (mibeasongSoldOutExists) then
        	    FailRow = FailRow + 1
        	ELSE

''response.write i&"="&Trim(RectOrderSerialArr(i))&"<br>"
                ''중복메일 방지용.
                sqlStr = "select d.orderserial from [db_academy].[dbo].tbl_academy_order_detail d"
                sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_master m"
                sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
                sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.

            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"
                end if

            	rsACADEMYget.CursorLocation = adUseClient
                rsACADEMYget.Open sqlStr, dbACADEMYget, adOpenForwardOnly
            	if Not rsACADEMYget.Eof then
            	    if Not (InStr(iMailOrderserialArr,rsACADEMYget("orderserial") + ",")>0) then
            	        iMailOrderserialArr = iMailOrderserialArr + rsACADEMYget("orderserial") + ","
            	    end if
            	end if
            	rsACADEMYget.close

            	sqlStr = "update D" & VbCRLF
            	sqlStr = sqlStr + " set currstate='7'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangno='" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,songjangdiv='" & Trim(RectSongjangdivArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCRLF
            	sqlStr = sqlStr + " ,passday=IsNULL(db_academy.dbo.fn_Ten_NetWorkDays((select convert(varchar(10),baljudate,21) from db_academy.dbo.tbl_academy_order_master where orderserial='" & Trim(RectOrderSerialArr(i)) & "'),IsNULL(convert(varchar(10),beasongdate,21),convert(varchar(10),getdate(),21))),0)"& VbCRLF
            	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_order_detail D"
            	sqlStr = sqlStr + "     Join [db_academy].[dbo].tbl_academy_order_master m"
                sqlStr = sqlStr + "     on m.orderserial=d.orderserial"
            	sqlStr = sqlStr + " where d.orderserial='" & Trim(RectOrderSerialArr(i)) & "'" & VbCRLF
            	sqlStr = sqlStr + " and d.detailidx =" & Trim(RectdetailidxArr(i))  & VbCRLF
            	sqlStr = sqlStr + " and d.makerid='" & Makerid & "'"
            	sqlStr = sqlStr + " and d.cancelyn<>'Y'"
            	sqlStr = sqlStr + " and m.cancelyn='N'"      '''취소 이전내역만.
            	if (mode="SongjangInputCSV") then
            	    sqlStr = sqlStr + " and IsNULL(d.currstate,0)<7"   ''완료후 송장번호 변경 할 수 있음.. :: 개별입력만 가능하도록.
                end if
    ''rw sqlStr
                dbACADEMYget.Execute sqlStr, AssignedRow

                TotAssignedRow = TotAssignedRow + AssignedRow

                if (AssignedRow=0) then FailRow = FailRow + 1
            END IF
        end if

    next


    '' ipkumdiv 8 추가

    sqlStr = "update [db_academy].[dbo].tbl_academy_order_master" & VbCRLF
    sqlStr = sqlStr + " set  ipkumdiv='7'" & VbCRLF
    sqlStr = sqlStr + " , beadaldate=getdate()" & VbCRLF
    sqlStr = sqlStr + " where orderserial in (" & VbCRLF
    sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_order_master m" & VbCRLF
    sqlStr = sqlStr + "         left join [db_academy].[dbo].tbl_academy_order_detail d" & VbCRLF
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
    sqlStr = sqlStr + "     where m.orderserial in ('" & orderserialArr & "')" & VbCRLF
    sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
    sqlStr = sqlStr + "     and jumundiv<>9" & VbCRLF
    sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
    sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )>0" & VbCRLF
    sqlStr = sqlStr + " ) "

    dbACADEMYget.Execute sqlStr


    sqlStr = "update [db_academy].[dbo].tbl_academy_order_master" & VbCRLF
    sqlStr = sqlStr + " set  ipkumdiv='8'" & VbCRLF
    sqlStr = sqlStr + " , beadaldate=getdate()" & VbCRLF
    sqlStr = sqlStr + " where orderserial in (" & VbCRLF
    sqlStr = sqlStr + "     select m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_order_master m" & VbCRLF
    sqlStr = sqlStr + "         left join [db_academy].[dbo].tbl_academy_order_detail d" & VbCRLF
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" & VbCRLF
    sqlStr = sqlStr + "     where m.orderserial in ('" & orderserialArr & "')" & VbCRLF
    sqlStr = sqlStr + "     and m.cancelyn='N'" & VbCRLF
    sqlStr = sqlStr + "     and m.jumundiv<>9" & VbCRLF
    sqlStr = sqlStr + "     and d.itemid<>0" & VbCRLF
    sqlStr = sqlStr + "     and d.cancelyn<>'Y'" & VbCRLF
    sqlStr = sqlStr + "     group by m.orderserial" & VbCRLF
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0"
    sqlStr = sqlStr + " ) "

    dbACADEMYget.Execute sqlStr



    ''#################################################
    ''메일보내기 루프
    ''#################################################
    iMailOrderserialArr = split(iMailOrderserialArr,",")

    if IsArray(iMailOrderserialArr) then
        for i=LBound(iMailOrderserialArr) to UBound(iMailOrderserialArr)
            if Trim(iMailOrderserialArr(i))<>"" then
                On Error resume Next
                if (application("Svr_Info")<>"Dev") or (SENDMAIL_YN = "Y") then
                    call fcSendMail_UpcheSendItem(iMailOrderserialArr(i), MakerID)
                end if
                on Error Goto 0
            end if
        next
    end if



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
    dbACADEMYget.close()	:	response.End

elseif (mode="misendInputOne") then
    ''출고 지연 아니면 ipgodate 널
    dim ckSendSMS, ckSendEmail, ckSendCall, sendState
    dim Sitemid, Sitemoption, itemSoldOut, optSoldOut

    sendState = "2"

    ''업체인경우
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

    if (MisendReason="05") then
        Sitemid     = RequestCheckVar(request("Sitemid"),10)
        Sitemoption = RequestCheckVar(request("Sitemoption"),4)
        itemSoldOut = RequestCheckVar(request("itemSoldOut"),4)

        if (Sitemid<>"") and (Sitemoption<>"") then
            if (Sitemoption="0000") then
                sqlStr = " update db_academy.dbo.tbl_diy_item" & VbCrlf
                sqlStr = sqlStr & " set sellyn='" & itemSoldOut & "'" & VbCrlf
                sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
                sqlStr = sqlStr & " where itemid=" & Sitemid

                dbget.Execute sqlStr
            else
                optSoldOut = "N"

                sqlStr = "update [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(Sitemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(Sitemoption) + "'"

				dbget.Execute sqlStr

				''옵션갯수
                sqlStr = "update [db_academy].[dbo].tbl_diy_item" + VBCrlf
                sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
                sqlStr = sqlStr + " from (" + VBCrlf
                sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
                sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
                sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
                sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
                sqlStr = sqlStr + " ) T" + VBCrlf
                sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item.itemid=" + CStr(Sitemid) + VBCrlf

                dbget.Execute sqlStr

                ''상품한정수량
            	sqlStr = "update [db_academy].[dbo].tbl_diy_item" + VBCrlf
            	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
            	sqlStr = sqlStr + " from (" + VBCrlf
            	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
            	sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
            	sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
            	sqlStr = sqlStr + " ) T" + VBCrlf
            	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item.itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " and [db_academy].[dbo].tbl_diy_item.optioncnt>0"

            	dbget.Execute sqlStr

            	'' 한정 판매 0 이면 일시 품절 처리
                sqlStr = " update [db_academy].[dbo].tbl_diy_item "
            	sqlStr = sqlStr + " set sellyn='S'"
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and sellyn='Y'"
            	sqlStr = sqlStr + " and limityn='Y'"
            	sqlStr = sqlStr + " and limitno-limitSold<1"

                dbget.Execute sqlStr

            	'' 판매중인 옵션이 없으면 품절처리
                sqlStr = " update [db_academy].[dbo].tbl_diy_item "
            	sqlStr = sqlStr + " set sellyn='N'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and optioncnt=0"

                dbget.Execute sqlStr

            end if
        end if
    end if

    sqlStr = " IF Exists(select idx from [db_academy].dbo.tbl_academy_mibeasong_list where detailidx=" & detailidx & ")"
    sqlStr = sqlStr + " BEGIN "
    sqlStr = sqlStr + "	    update [db_academy].dbo.tbl_academy_mibeasong_list"
    sqlStr = sqlStr + "	    set code='" & MisendReason & "'"
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	    ,state='"&sendState&"'"                                         ''상태 변경 (기존 안내메일완료)
        sqlStr = sqlStr + "	    ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS발송완료
        sqlStr = sqlStr + "	    ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email발송완료
        '''sqlStr = sqlStr + "	    ,isSendCall=(CASE WHEN isSendCall='Y' then 'Y' ELSE '"&ckSendCall&"' END)"  '' CALL완료 : 따로 처리
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
    sqlStr = sqlStr + "	    insert into [db_academy].dbo.tbl_academy_mibeasong_list"
    sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
    sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"             ''상태 변경 (기존 안내메일완료)
        ''sqlStr = sqlStr + "	isSendCall,"
    end if
    sqlStr = sqlStr + "	    itemname, itemoptionname)"
    sqlStr = sqlStr + "	    select detailidx, orderserial, itemid,itemoption,"
    sqlStr = sqlStr + "	    itemno, itemno, '" & MisendReason & "',"

    if (ipgodate<>"") then
        sqlStr = sqlStr + "	'" & ipgodate & "','',"
    else
        sqlStr = sqlStr + "	NULL,'',"
    end if
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
        ''sqlStr = sqlStr + "	 '"&ckSendCall&"',"
    end if
    sqlStr = sqlStr + "	    itemname, itemoptionname"
    sqlStr = sqlStr + "	    from [db_academy].[dbo].tbl_academy_order_detail"
    sqlStr = sqlStr + "	    where detailidx=" & detailidx
    sqlStr = sqlStr + " END "
''rw   sqlStr
    dbget.Execute sqlStr


    ''SMS 발송 + [CS메모에 저장 -> 같이 되어있음.]
    if (ckSendSMS="Y") then
        if (application("Svr_Info")<>"Dev") or (SENDMAIL_YN = "Y") then
            call SendMiChulgoSMS(detailidx)
        end if
   end if
    ''EMail발송
    if (ckSendEmail="Y") then
        if (application("Svr_Info")<>"Dev") or (SENDMAIL_YN = "Y") then
            call fcSendMail_SendMiChulgoMail(detailidx)
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

''dim chkCount, chkIdx
''dim ArrChkVal
''
''chkCount = request("chkIdx").count
''
'''rw "chkCount=" & chkCount
'''rw "ArrChkVal=" & request("ArrChkVal")
'''rw "chkidx=" & request("chkidx")
'''rw "detailidx=" & request("detailidx")
'''rw "MisendReason=" & request("MisendReason")
''
''ArrChkVal = split(request("ArrChkVal"),",")
''
''if (mode="misendInput") then
''    for i=1 to chkCount
''        chkIdx      = ArrChkVal(i-1)
''''rw "chkIdx="&chkIdx
''        detailidx   = request("detailidx")(chkIdx)
''        MisendReason= request("MisendReason")(chkIdx)
''        ipgodate    = request("ipgodate" + CStr(chkIdx-1))
''
''
''        ''출고 지연 아니면 ipgodate 널
''        if (MisendReason="05") then
''            ipgodate =""
''        end if
''
''        sqlStr = " IF Exists(select idx from [db_academy].dbo.tbl_academy_mibeasong_list where detailidx=" & detailidx & ")"
''        sqlStr = sqlStr + " BEGIN "
''        sqlStr = sqlStr + "	    update [db_academy].dbo.tbl_academy_mibeasong_list"
''        sqlStr = sqlStr + "	    set code='" & MisendReason & "'"
''        if (ipgodate<>"") then
''            sqlStr = sqlStr + "	,ipgodate='" & ipgodate & "'"
''        else
''            sqlStr = sqlStr + "	,ipgodate=NULL"
''        end if
''        sqlStr = sqlStr + "	    where detailidx=" & detailidx
''        sqlStr = sqlStr + " END "
''        sqlStr = sqlStr + " ELSE "
''        sqlStr = sqlStr + " BEGIN "
''        sqlStr = sqlStr + "	    insert into [db_academy].dbo.tbl_academy_mibeasong_list"
''        sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
''        sqlStr = sqlStr + "	    itemno, itemlackno, code, state, ipgodate, reqstr, "
''        sqlStr = sqlStr + "	    itemname, itemoptionname)"
''        sqlStr = sqlStr + "	    select idx, orderserial, itemid,itemoption,"
''        sqlStr = sqlStr + "	    itemno, itemno, '" & MisendReason & "', 0,"
''        if (ipgodate<>"") then
''            sqlStr = sqlStr + "	'" & ipgodate & "',"
''        else
''            sqlStr = sqlStr + "	NULL,"
''        end if
''        sqlStr = sqlStr + "	    '', "
''        sqlStr = sqlStr + "	    itemname, itemoptionname"
''        sqlStr = sqlStr + "	    from [db_academy].[dbo].tbl_academy_order_detail"
''        sqlStr = sqlStr + "	    where idx=" & detailidx
''        sqlStr = sqlStr + " END "
''''rw   sqlStr
''        dbget.Execute sqlStr
''    next
''
''    response.write "<script language='javascript'>alert('처리 되었습니다.')</script>"
''    response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
''    dbget.close()	:	response.End
''end if
%>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
