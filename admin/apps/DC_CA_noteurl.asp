<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<%

    dim mailContents
    dim sqlStr, paramInfo, retParamInfo
    dim TmpOrderserial, TmpUserID, Tmpbuyhp

    if (request("LGD_CASFLAG")="") then response.end

    ''call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 incomming", request("LGD_OID") & "|" & request("LGD_CASFLAG"))



if   (request("LGD_OID")="18082778608") or  (request("LGD_OID")="18051675934") or  (request("LGD_OID")="18062970563") or   (request("LGD_OID")="18062865953") or  (request("LGD_OID")="18062969848") or (request("LGD_OID")="18062970460") or (request("LGD_OID")="16061720390") or (request("LGD_OID")="16061723044") or (request("LGD_OID")="16061723082") or  (request("LGD_OID")="16050805362") or   (request("LGD_OID")="16050908108") or  (request("LGD_OID")="16050694309") or  (request("LGD_OID")="16061723007") or (request("LGD_OID")="16061723026") or (request("LGD_OID")="16050908698") or  (request("LGD_OID")="16050908685") or  (request("LGD_OID")="16061723009") or (request("LGD_OID")="16061619185") or (request("LGD_OID")="16061618882") or  (request("LGD_OID")="16061720610") or (request("LGD_OID")="16061722706") or (request("LGD_OID")="16061721883") or (request("LGD_OID")="14112909580") or (request("LGD_OID")="14112909888") or (request("LGD_OID")="16061722753") or (request("LGD_OID")="16061723090") or (request("LGD_OID")="13030216705") or (request("LGD_OID")="13031156078") or (request("LGD_OID")="13030952062")  or (request("LGD_OID")="16061723051") or (request("LGD_OID")="13032626552") then  ''10041273170 ''10051389839 ''iisue00
    response.write "OK"
    response.end
end if

if (request("LGD_OID")="18091193461") or  (request("LGD_OID")="18062970487") or  (request("LGD_OID")="18031198038") or  (request("LGD_OID")="18031198265") or  (request("LGD_OID")="18031198214") then
    response.write "OK"
    response.end
end if

if (request("LGD_OID")="18062751384") or  (request("LGD_OID")="18062865607") or  (request("LGD_OID")="18031196413") or  (request("LGD_OID")="18031198177") or  (request("LGD_OID")="18031198164") then
    response.write "OK"
    response.end
end if

if (request("LGD_OID")="18062970124") or  (request("LGD_OID")="18031198160") or  (request("LGD_OID")="18031197532") or  (request("LGD_OID")="18031198150") or  (request("LGD_OID")="18062970577") then
    response.write "OK"
    response.end
end if    

if (request("LGD_OID")="18062648172") or  (request("LGD_OID")="18062865746") or  (request("LGD_OID")="18062970535") or  (request("LGD_OID")="18031198203") or  (request("LGD_OID")="18030667639") then
    response.write "OK"
    response.end
end if 

if (request("LGD_OID")="18031198294") or  (request("LGD_OID")="18031198147") or  (request("LGD_OID")="18031198105") or  (request("LGD_OID")="18031197165") or  (request("LGD_OID")="18062968861") then
    response.write "OK"
    response.end
end if 

if (request("LGD_OID")="18062970283") or  (request("LGD_OID")="18062969935") or  (request("LGD_OID")="18062970564") or  (request("LGD_OID")="18062970558") or  (request("LGD_OID")="18062970539") then
    response.write "OK"
    response.end
end if 

if (request("LGD_OID")="18031198175") or  (request("LGD_OID")="18031198170") or  (request("LGD_OID")="18031198046") or  (request("LGD_OID")="18031198045") or  (request("LGD_OID")="18031198056") then
    response.write "OK"
    response.end
end if 


    
    '/*
    ' * [상점 결제결과처리(DB) 페이지]
    ' *
    ' * 1) 위변조 방지를 위한 hashdata값 검증은 반드시 적용하셔야 합니다.
    ' *
    ' */
    dim LGD_RESPCODE    : LGD_RESPCODE            = trim(request("LGD_RESPCODE"))             '// 응답코드: 0000(성공) 그외 실패
    dim LGD_RESPMSG     : LGD_RESPMSG             = trim(request("LGD_RESPMSG"))              '// 응답메세지
    dim LGD_MID         : LGD_MID                 = trim(request("LGD_MID"))                  '// 상점아이디
    dim LGD_OID         : LGD_OID                 = trim(request("LGD_OID"))                  '// 주문번호
    dim LGD_AMOUNT      : LGD_AMOUNT              = trim(request("LGD_AMOUNT"))               '// 거래금액
    dim LGD_TID         : LGD_TID                 = trim(request("LGD_TID"))                  '// 데이콤이 부여한 거래번호
    dim LGD_PAYTYPE     : LGD_PAYTYPE             = trim(request("LGD_PAYTYPE"))              '// 결제수단코드
    dim LGD_PAYDATE     : LGD_PAYDATE             = trim(request("LGD_PAYDATE"))              '// 거래일시(승인일시/이체일시)
    dim LGD_HASHDATA    : LGD_HASHDATA            = trim(request("LGD_HASHDATA"))             '// 해쉬값
    dim LGD_FINANCECODE : LGD_FINANCECODE         = trim(request("LGD_FINANCECODE"))          '// 결제기관코드(은행코드)
    dim LGD_FINANCENAME : LGD_FINANCENAME         = trim(request("LGD_FINANCENAME"))          '// 결제기관이름(은행이름)
    dim LGD_ESCROWYN    : LGD_ESCROWYN            = trim(request("LGD_ESCROWYN"))             '// 에스크로 적용여부
    dim LGD_TIMESTAMP   : LGD_TIMESTAMP           = trim(request("LGD_TIMESTAMP"))            '// 타임스탬프
    dim LGD_ACCOUNTNUM  : LGD_ACCOUNTNUM          = trim(request("LGD_ACCOUNTNUM"))           '// 계좌번호(무통장입금)
    dim LGD_CASTAMOUNT  : LGD_CASTAMOUNT          = trim(request("LGD_CASTAMOUNT"))           '// 입금총액(무통장입금)
    dim LGD_CASCAMOUNT  : LGD_CASCAMOUNT          = trim(request("LGD_CASCAMOUNT"))           '// 현입금액(무통장입금)
    dim LGD_CASFLAG     : LGD_CASFLAG             = trim(request("LGD_CASFLAG"))              '// 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
    dim LGD_CASSEQNO    : LGD_CASSEQNO            = trim(request("LGD_CASSEQNO"))             '// 입금순서(무통장입금)
    dim LGD_CASHRECEIPTNUM      : LGD_CASHRECEIPTNUM      = trim(request("LGD_CASHRECEIPTNUM"))       '// 현금영수증 승인번호
    dim LGD_CASHRECEIPTSELFYN   : LGD_CASHRECEIPTSELFYN   = trim(request("LGD_CASHRECEIPTSELFYN"))    '// 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
    dim LGD_CASHRECEIPTKIND     : LGD_CASHRECEIPTKIND     = trim(request("LGD_CASHRECEIPTKIND"))      '// 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
	dim LGD_PAYER               : LGD_PAYER            	  = trim(request("LGD_PAYER"))             	'// 입금자명


	''tbl_cyberAcctNoti_Log
	paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	        ,Array("@LGD_RESPCODE"  , adVarchar	, adParamInput, 4 , LGD_RESPCODE)	_
	        ,Array("@LGD_RESPMSG"  , adVarchar	, adParamInput, 160 , LGD_RESPMSG)	_
	        ,Array("@LGD_MID"  , adVarchar	, adParamInput, 15 , LGD_MID)	_
	        ,Array("@LGD_OID"  , adVarchar	, adParamInput, 64 , LGD_OID)	_
	        ,Array("@LGD_AMOUNT"  , adCurrency	, adParamInput,  , LGD_AMOUNT)	_
	        ,Array("@LGD_TID"  , adVarchar	, adParamInput, 24 , LGD_TID)	_
	        ,Array("@LGD_PAYTYPE"  , adVarchar	, adParamInput, 6 , LGD_PAYTYPE)	_
	        ,Array("@LGD_PAYDATE"  , adVarchar	, adParamInput, 14 , LGD_PAYDATE)	_
	        ,Array("@LGD_FINANCECODE"  , adVarchar	, adParamInput, 10 , LGD_FINANCECODE)	_
	        ,Array("@LGD_FINANCENAME"  , adVarchar	, adParamInput, 20 , LGD_FINANCENAME)	_
	        ,Array("@LGD_ESCROWYN"  , adVarchar	, adParamInput, 1 , LGD_ESCROWYN)	_
	        ,Array("@LGD_TIMESTAMP"  , adVarchar	, adParamInput, 14 , LGD_TIMESTAMP)	_
	        ,Array("@LGD_ACCOUNTNUM"  , adVarchar	, adParamInput, 15 , LGD_ACCOUNTNUM)	_
	        ,Array("@LGD_CASTAMOUNT"  , adCurrency	, adParamInput,  , LGD_CASTAMOUNT)	_
	        ,Array("@LGD_CASCAMOUNT"  , adCurrency	, adParamInput,  , LGD_CASCAMOUNT)	_
	        ,Array("@LGD_CASFLAG"  , adVarchar	, adParamInput, 10 , LGD_CASFLAG)	_
	        ,Array("@LGD_CASSEQNO"  , adVarchar	, adParamInput, 3 , LGD_CASSEQNO)	_
	        ,Array("@LGD_CASHRECEIPTNUM"  , adVarchar	, adParamInput, 9 , LGD_CASHRECEIPTNUM)	_
	        ,Array("@LGD_CASHRECEIPTSELFYN"  , adVarchar	, adParamInput, 1 , LGD_CASHRECEIPTSELFYN)	_
	        ,Array("@LGD_CASHRECEIPTKIND"  , adVarchar	, adParamInput, 4 , LGD_CASHRECEIPTKIND)	_
	        ,Array("@LGD_PAYER"  , adVarchar	, adParamInput, 16 , LGD_PAYER)	_
	)

	sqlStr = "db_order.dbo.sp_Ten_CyberAcct_LogSave"
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

    dim LogIdx : LogIdx      = GetValue(retParamInfo, "@RETURN_VALUE")   ' 에러내용

    if (LogIdx<0) then
        'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]로그 저장중 오류", request("LGD_OID") & "|" & request("LGD_CASFLAG"))

        response.write "ERR:1"
        response.end
    end if

    '/*
    ' * 구매정보
    ' */
''    dim LGD_BUYER       : LGD_BUYER               = trim(request("LGD_BUYER"))                '// 구매자
''    dim LGD_PRODUCTINFO : LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))          '// 상품명
''    dim LGD_BUYERID     : LGD_BUYERID             = trim(request("LGD_BUYERID"))              '// 구매자 ID
''    dim LGD_BUYERADDRESS: LGD_BUYERADDRESS        = trim(request("LGD_BUYERADDRESS"))         '// 구매자 주소
''    dim LGD_BUYERPHONE  : LGD_BUYERPHONE          = trim(request("LGD_BUYERPHONE"))           '// 구매자 전화번호
''    dim LGD_BUYEREMAIL  : LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))           '// 구매자 이메일
''    dim LGD_BUYERSSN    : LGD_BUYERSSN            = trim(request("LGD_BUYERSSN"))             '// 구매자 주민번호
''    dim LGD_PRODUCTCODE : LGD_PRODUCTCODE         = trim(request("LGD_PRODUCTCODE"))          '// 상품코드
''    dim LGD_RECEIVER    : LGD_RECEIVER            = trim(request("LGD_RECEIVER"))             '// 수취인
''    dim LGD_RECEIVERPHONE   : LGD_RECEIVERPHONE   = trim(request("LGD_RECEIVERPHONE"))        '// 수취인 전화번호
''    dim LGD_DELIVERYINFO    : LGD_DELIVERYINFO    = trim(request("LGD_DELIVERYINFO"))         '// 배송지


    '/*
    ' * hashdata 검증을 위한 mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다.
    ' * LG데이콤에서 발급한 상점키로 반드시변경해 주시기 바랍니다.
    ' */
    dim LGD_MERTKEY : LGD_MERTKEY = "1af44018218ae6e8f6e14b3797b3f094"  '//mertkey
    dim LGD_HASHDATA2 : LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

    '/*
    ' * 상점 처리결과 리턴메세지
    ' *
    ' * OK  : 상점 처리결과 성공
    ' * 그외 : 상점 처리결과 실패
    ' *
    ' * ※ 주의사항 : 성공시 'OK' 문자이외의 다른문자열이 포함되면 실패처리 되오니 주의하시기 바랍니다.
    ' */
    dim resultMSG : resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 결과값을 입력해 주시기 바랍니다."

    dim orderserial
    dim RetErr, RetMsg, retval
    dim userid, buyname, buyhp, buyEmail, jumundiv
    dim osms
    if UCASE(LGD_HASHDATA2) = UCASE(LGD_HASHDATA) then
        '//해쉬값 검증이 성공이면
        if LGD_RESPCODE = "0000" then
            '//결제가 성공이면
            if LGD_CASFLAG = "R" then
                '/*
                ' * 무통장 할당 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */
                resultMSG = "OK"
            elseif LGD_CASFLAG = "I" then
                '/*
                ' * 무통장 입금 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */

                ''db_order.dbo.tbl_order_CyberAccountLog 확인
                paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                    ,Array("@Orderserial"  , adVarchar	, adParamInput, 11 , LGD_OID)	_
                    ,Array("@BackUserID", adVarchar	, adParamInput, 32, "system")	_
        			,Array("@LGD_TID", adVarchar	, adParamInput, 24, LGD_TID)	_
        			,Array("@LGD_AMOUNT", adCurrency	, adParamInput,, LGD_AMOUNT)	_
        			,Array("@LGD_CASTAMOUNT", adCurrency	, adParamInput,, LGD_CASTAMOUNT)	_
        			,Array("@LGD_FINANCECODE", adVarchar	, adParamInput,8, LGD_FINANCECODE)	_
        			,Array("@LGD_ACCOUNTNUM", adVarchar	, adParamInput,16, LGD_ACCOUNTNUM)	_
        			,Array("@LGD_CASHRECEIPTNUM", adVarchar	, adParamInput,16, LGD_CASHRECEIPTNUM)	_
        			,Array("@LGD_CASHRECEIPTSELFYN", adVarchar	, adParamInput,16, LGD_CASHRECEIPTSELFYN)	_
        			,Array("@LGD_CASHRECEIPTKIND", adchar	, adParamInput,1, LGD_CASHRECEIPTKIND)	_
        			,Array("@LGD_PAYER", adVarchar	, adParamInput,32, LGD_PAYER)	_
        			,Array("@LogIdx", adInteger	, adParamInput,, LogIdx)	_
        			,Array("@RetVal"	, adInteger  , adParamOutput,, 0) _
        			,Array("@RetMsg", adVarchar	, adParamOutput,100,"") _
        			,Array("@MatchOrderSerial", adVarchar	, adParamOutput,11,"") _
        		)

                sqlStr = "db_order.dbo.sp_Ten_IpkumConfirm_Cyber_Proc"
                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

                RetErr      = GetValue(retParamInfo, "@RETURN_VALUE")   ' 에러내용
                retval      = GetValue(retParamInfo, "@RetVal")         '
                RetMsg      = GetValue(retParamInfo, "@RetMsg")
                orderserial  = GetValue(retParamInfo, "@MatchOrderSerial") ' 매칭된 주문번호

                if (RetErr=0) then
                    if (retval=1)  then
                        if(orderserial<>"") then
                            sqlStr = "select top 1 userid, buyname, buyhp, buyemail, jumundiv from [db_order].[dbo].tbl_order_master"
                            sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                            sqlStr = sqlStr + " and cancelyn='N'"

                            rsget.Open sqlStr,dbget,1
                            if Not rsget.Eof then
                                userid  = rsget("userid")
                            	buyname = db2html(rsget("buyname"))
                            	buyhp = db2html(rsget("buyhp"))
                            	buyemail = db2html(rsget("buyemail"))
                            	jumundiv = db2html(rsget("jumundiv"))
                            end if
                            rsget.close

                            ''SMS 발송 : 데이콤에서 발송하므로 발송안함.

                            'set osms = new CSMSClass
                            'osms.SendAcctIpkumOkMsg buyhp,orderserial
                            'set osms = Nothing

                            ''Email발송
                            IF (jumundiv="7") or (jumundiv="4") then
                                call sendmailbankokNoDLV(buyemail,buyname,orderserial)
                            ELSE
                                call sendmailbankok(buyemail,buyname,orderserial)
                            END IF

                            ''네이트온 메세지.
''                            dim oXML
''                            If (userid<>"") then
''                                On Error resume Next
''                            		'// POST로 전송
''                            		'실서버측 알림전송 처리 페이지로 정보 전달
''                            		set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
''                                    if (application("Svr_Info")<>"Dev") then
''                            			oXML.open "POST", "http://www1.10x10.co.kr/apps/nateon/interface/check_alarmSend.asp", false
''                            		else
''                            			oXML.open "POST", "http://2009www.10x10.co.kr/apps/nateon/interface/check_alarmSend.asp", false
''                            		end if
''                            		oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                            		oXML.send "arid=166&ordsn=" & orderserial	'파라메터 전송
''                            		Set oXML = Nothing	'컨퍼넌트 해제
''                                on Error Goto 0
''                            End If
                        end if

                        resultMSG = "OK"
                        if (LGD_CASSEQNO<>"001") then
                            'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 완료" & orderserial, resultMSG)
                        end if
                    else

                        resultMSG = "ERR : orderserial=" & LGD_OID & " : retval=" & retval & " : RetMsg=" & RetMsg
                        'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 오류 발생" & orderserial, resultMSG)

                        if (retval=-1) then
                            ''기 취소된 무통장건에 대해 입금처리 요청이 들어올경우.

                        end if

                        ''[803][Already Match in Log]:  2019/01/15 부터 자주 생김.
                        if (retval=-8) then
                            response.write "OK"
                            dbget.close()	:	response.End
                        end if
                    end if


                ELSE
                    resultMSG = "ERR"
                    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 ERR", trim(request("LGD_OID")) & "|"&trim(request("LGD_ACCOUNTNUM"))&"|"&resultMSG&"|"&RetErr)

                END IF


            elseif LGD_CASFLAG = "C" then
                '/*
                ' * 무통장 입금취소 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */
                '' 결제후 입금완료상태 입금취소인경우 => 입금이전으로
                ''
                '' 결제전 입금취소인경우
                sqlStr = " select orderserial, userid, buyhp from db_order.dbo.tbl_order_master"
                sqlStr = sqlStr & " where orderserial='" + LGD_OID + "'" & VbCRLF
                sqlStr = sqlStr & " and ipkumdiv=4"
                sqlStr = sqlStr & " and cancelyn='N'"
                rsget.Open sqlStr,dbget,1
                if Not rsget.Eof then
                    TmpOrderserial = rsget("orderserial")
                    TmpUserID      = rsget("userid")
                    Tmpbuyhp      = rsget("buyhp")
                end if
                rsget.Close

                if (TmpOrderserial<>"") then
                    sqlStr = " update db_order.dbo.tbl_order_master"
                    sqlStr = sqlStr & " set ipkumdiv=2"
                    sqlStr = sqlStr & " , ipkumdate=NULL"
                    sqlStr = sqlStr & " where orderserial='" + LGD_OID + "'" & VbCRLF
                    sqlStr = sqlStr + " and ipkumdiv=4"
                    sqlStr = sqlStr + " and cancelyn='N'"

                    dbget.Execute sqlStr

                    ''로그 다시 미 매칭으로 변경..
                    sqlStr = " update  db_order.dbo.tbl_order_CyberAccountLog"
                    sqlStr = sqlStr + " set isMatched='N'"
                    sqlStr = sqlStr + " where orderserial='" + LGD_OID + "'"
                    dbget.Execute sqlStr


                    ''메모남김.
                    sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo"
                    sqlStr = sqlStr + " (orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
                    sqlStr = sqlStr + " values('" + CStr(TmpOrderserial) + "','1','" + CStr(TmpUserID) + "','0','99','','system','system','입금취소 - 가상계좌 " & trim(request("LGD_ACCOUNTNUM")) & "," & trim(request("LGD_AMOUNT")) & " ','Y',getdate(),getdate()) "
                    dbget.Execute sqlStr

                    resultMSG = "OK"
                    On Error Resume Next
                    set osms = new CSMSClass
                    osms.SendAcctIpkumCancelMsg Tmpbuyhp,LGD_OID
                    set osms = Nothing
                    On Error Goto 0

                    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금취소 완료", trim(request("LGD_OID")) & "|"&trim(request("LGD_ACCOUNTNUM"))&"|"&resultMSG&"|"&RetErr)

                else
                    resultMSG = "ERR"

                    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 오류 - 취소분", trim(request("LGD_OID")) & "|"&trim(request("LGD_ACCOUNTNUM"))&"|"&resultMSG&"|"&RetErr)

                end if

            end if
        else
            '//결제가 실패이면
            '/*
            ' * 거래실패 결과 상점 처리(DB) 부분
            ' * 상점결과 처리가 정상이면 "OK"
            ' */
            resultMSG = "OK"
        end if
    else
        '//해쉬값이 검증이 실패이면
        '/*
        ' * hashdata검증 실패 로그를 처리하시기 바랍니다.
        ' */
        resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 해쉬값 검증이 실패하였습니다."

        mailContents = resultMSG & "<br>"
        mailContents = mailContents & "LGD_MID=" & LGD_MID & "<br>"
        mailContents = mailContents & "LGD_OID=" & LGD_OID & "<br>"
        mailContents = mailContents & "LGD_AMOUNT=" & LGD_AMOUNT & "<br>"
        mailContents = mailContents & "LGD_RESPCODE=" & LGD_RESPCODE & "<br>"
        mailContents = mailContents & "LGD_TIMESTAMP=" & LGD_TIMESTAMP & "<br>"
        mailContents = mailContents & "LGD_MERTKEY=" & LGD_MERTKEY & "<br>"

        mailContents = mailContents & "LGD_HASHDATA2=" & LGD_HASHDATA2 & "<br>"
        mailContents = mailContents & "LGD_HASHDATA=" & LGD_HASHDATA & "<br>"

        'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[가상계좌]["&Left(now(),10)&"]입금확인 incomming", mailContents)

    end if

    Response.Write(resultMSG)
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->