<%@ language=vbscript %>
<% option explicit %>
<%
Response.CharSet = "euc-kr"
%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbAcademyHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<%

    Dim ref_ip :ref_ip   = request.ServerVariables( "REMOTE_ADDR" )

    IF (LEFT(ref_ip,10)<>"61.252.133") or (LEFT(ref_ip,10)<>"203.238.36") then
        response.write ""
    END IF

    dim mailContents
    dim sqlStr, paramInfo, retParamInfo
    dim TmpOrderserial, TmpUserID, Tmpbuyhp


    Dim site_cd     : site_cd      = request( "site_cd"  )                   ' ����Ʈ �ڵ�
    Dim tno         : tno          = request( "tno"      )                   ' KCP �ŷ���ȣ
    Dim order_no    : order_no     = request( "order_no" )                   ' �ֹ���ȣ
    Dim tx_cd       : tx_cd        = request( "tx_cd"    )                   ' ����ó�� ���� �ڵ�
    Dim tx_tm       : tx_tm        = request( "tx_tm"    )                   ' ����ó�� �Ϸ� �ð�
    '/* = -------------------------------------------------------------------------- = */
    Dim ipgm_name   : ipgm_name    = ""                                      ' �ֹ��ڸ�
    Dim remitter    : remitter     = ""                                      ' �Ա��ڸ�
    Dim ipgm_mnyx   : ipgm_mnyx    = ""                                      ' �Ա� �ݾ�
    Dim bank_code   : bank_code    = ""                                      ' �����ڵ�
    Dim account     : account      = ""                                      ' ������� �Աݰ��¹�ȣ
    Dim op_cd       : op_cd        = ""                                      ' ó������ �ڵ�
    Dim noti_id     : noti_id      = ""                                      ' �뺸 ���̵�

    if tx_cd = "TX00" then

        ipgm_name = request( "ipgm_name"  )            ' �ֹ��ڸ�
        remitter  = request( "remitter"   )            ' �Ա��ڸ�
        ipgm_mnyx = request( "ipgm_mnyx"  )            ' �Ա� �ݾ�
        bank_code = request( "bank_code"  )            ' �����ڵ�
        account   = request( "account"    )            ' ������� �Աݰ��¹�ȣ
        op_cd     = request( "op_cd"      )            ' ó������ �ڵ�
        noti_id   = request( "noti_id"    )            ' �뺸 ���̵�

    end if

    dim buf
    buf = "site_cd="&site_cd&"<br>"
    buf = buf &"tno="&tno&"<br>"
    buf = buf &"order_no="&order_no&"<br>"
    buf = buf &"tx_cd="&tx_cd&"<br>"
    buf = buf &"tx_tm="&tx_tm&"<br>"
    buf = buf &"ipgm_name="&ipgm_name&"<br>"
    buf = buf &"remitter="&remitter&"<br>"
    buf = buf &"ipgm_mnyx="&ipgm_mnyx&"<br>"
    buf = buf &"bank_code="&bank_code&"<br>"
    buf = buf &"account="&account&"<br>"
    buf = buf &"op_cd="&op_cd&"<br>"
    buf = buf &"noti_id="&noti_id&"<br>"

    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������KCP]["&Left(now(),10)&"]", buf)


    if (request("order_no")="10082588076") then  ''10041273170 ''10051389839
        response.write "<html><body><form><input type=""hidden"" name=""result"" value=""0000""></form></body></html>"
        response.end
    end if

    ''UTF-8 �ε�.
    IF (application("Svr_Info")	= "Dev") then
        ipgm_name="ipgm_name"
        remitter="remitter"
        if (ipgm_mnyx="") then ipgm_mnyx="0"
    END IF

''    site_cd="T0000"
''    tno="20101110910731"
''    order_no="Y0111077872"
''    tx_cd="TX00"
''    tx_tm="20101110183911"
''    ipgm_name="KKKK"
''    remitter="HHH"
''    ipgm_mnyx="11500"
''    bank_code="03"
''    account="T0300000030044"
''    op_cd="18"
''    noti_id="10111003381951040102"


    ''tbl_cyberAcctNoti_Log

	paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
	        ,Array("@LGD_RESPCODE"  , adVarchar	, adParamInput, 4 , tx_cd)	_
	        ,Array("@LGD_RESPMSG"  , adVarchar	, adParamInput, 160 , tx_tm)	_
	        ,Array("@LGD_MID"  , adVarchar	, adParamInput, 15 , site_cd)	_
	        ,Array("@LGD_OID"  , adVarchar	, adParamInput, 64 , order_no)	_
	        ,Array("@LGD_AMOUNT"  , adCurrency	, adParamInput,  , ipgm_mnyx)	_
	        ,Array("@LGD_TID"  , adVarchar	, adParamInput, 24 , tno)	_
	        ,Array("@LGD_PAYTYPE"  , adVarchar	, adParamInput, 6 , op_cd)	_
	        ,Array("@LGD_PAYDATE"  , adVarchar	, adParamInput, 14 , tx_tm)	_
	        ,Array("@LGD_FINANCECODE"  , adVarchar	, adParamInput, 10 , bank_code)	_
	        ,Array("@LGD_FINANCENAME"  , adVarchar	, adParamInput, 20 , bank_code)	_
	        ,Array("@LGD_ESCROWYN"  , adVarchar	, adParamInput, 1 , "N")	_
	        ,Array("@LGD_TIMESTAMP"  , adVarchar	, adParamInput, 14 , tx_tm)	_
	        ,Array("@LGD_ACCOUNTNUM"  , adVarchar	, adParamInput, 15 , account)	_
	        ,Array("@LGD_CASTAMOUNT"  , adCurrency	, adParamInput,  , ipgm_mnyx)	_
	        ,Array("@LGD_CASCAMOUNT"  , adCurrency	, adParamInput,  , ipgm_mnyx)	_
	        ,Array("@LGD_CASFLAG"  , adVarchar	, adParamInput, 10 , tx_cd)	_
	        ,Array("@LGD_CASSEQNO"  , adVarchar	, adParamInput, 3 , "0")	_
	        ,Array("@LGD_CASHRECEIPTNUM"  , adVarchar	, adParamInput, 9 , "")	_
	        ,Array("@LGD_CASHRECEIPTSELFYN"  , adVarchar	, adParamInput, 1 , "")	_
	        ,Array("@LGD_CASHRECEIPTKIND"  , adVarchar	, adParamInput, 4 , "")	_
	        ,Array("@LGD_PAYER"  , adVarchar	, adParamInput, 16 , remitter)	_
	)

	sqlStr = "db_academy.dbo.sp_ACA_CyberAcct_LogSave"
    retParamInfo = dbacademy_fnExecSPOutput(sqlStr,paramInfo)

    dim LogIdx : LogIdx      = GetValue(retParamInfo, "@RETURN_VALUE")   ' ��������

    if (LogIdx<0) then
        'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������KCP]["&Left(now(),10)&"]�α� ������ ����", request("order_no") & "|" & request("tx_cd") & "|" & request("op_cd"))

        response.write "ERR:1"
        response.end
    end if

''Response.Write("<html><body><form><input type=""hidden"" name=""result"" value=""0000""></form></body></html>")
''Response.end

    dim resultMSG

    dim orderserial
    dim RetErr, RetMsg, retval
    dim userid, buyname, buyhp, buyEmail
    dim osms

    if (tx_cd = "TX00") then
        '//������ �����̸�
        if (tx_cd = "TX00") and (op_cd <>"13") then
            '/*
            ' * ������ �Ա� ���� ��� ���� ó��(DB) �κ�
            ' * ���� ��� ó���� �����̸� "OK"
            ' */

            ''db_order.dbo.tbl_order_CyberAccountLog Ȯ��
            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@Orderserial"  , adVarchar	, adParamInput, 11 , order_no)	_
                ,Array("@BackUserID", adVarchar	, adParamInput, 32, "system")	_
    			,Array("@LGD_TID", adVarchar	, adParamInput, 24, tno)	_
    			,Array("@LGD_AMOUNT", adCurrency	, adParamInput,, ipgm_mnyx)	_
    			,Array("@LGD_CASTAMOUNT", adCurrency	, adParamInput,, ipgm_mnyx)	_
    			,Array("@LGD_FINANCECODE", adVarchar	, adParamInput,8, bank_code)	_
    			,Array("@LGD_ACCOUNTNUM", adVarchar	, adParamInput,16, account)	_
    			,Array("@LGD_CASHRECEIPTNUM", adVarchar	, adParamInput,16, "")	_
    			,Array("@LGD_CASHRECEIPTSELFYN", adVarchar	, adParamInput,16, "")	_
    			,Array("@LGD_CASHRECEIPTKIND", adchar	, adParamInput,1, "")	_
    			,Array("@LGD_PAYER", adVarchar	, adParamInput,32, remitter)	_
    			,Array("@LogIdx", adInteger	, adParamInput,, LogIdx)	_
    			,Array("@RetVal"	, adInteger  , adParamOutput,, 0) _
    			,Array("@RetMsg", adVarchar	, adParamOutput,100,"") _
    			,Array("@MatchOrderSerial", adVarchar	, adParamOutput,11,"") _
    		)

            sqlStr = "db_academy.dbo.sp_ACA_IpkumConfirm_Cyber_Proc"
            retParamInfo = dbacademy_fnExecSPOutput(sqlStr,paramInfo)

            RetErr      = GetValue(retParamInfo, "@RETURN_VALUE")   ' ��������
            retval      = GetValue(retParamInfo, "@RetVal")         '
            RetMsg      = GetValue(retParamInfo, "@RetMsg")
            orderserial  = GetValue(retParamInfo, "@MatchOrderSerial") ' ��Ī�� �ֹ���ȣ

            if (RetErr=0) then
                if (retval=1)  then
                    if(orderserial<>"") then
                        '''�ֹ� ���ϸ��� ������Ʈ---------------------------------------------------
                    	Dim totmile : totmile = 0
						Dim michulgoMile : michulgoMile = 0

                    	sqlStr = "select userid from [db_academy].[dbo].tbl_academy_order_master where orderserial='"&orderserial&"'"
                    	rsAcademyget.Open sqlStr,dbAcademyget,1
                        if Not rsAcademyget.Eof then
                            userid    = rsAcademyget("userid")
                        end if
                        rsAcademyget.Close


                    	IF (userid<>"") THEN
                        	sqlStr = "select sum(totalmileage) as totmile, IsNull(sum(case when sitename = 'academy' and ipkumdiv < 7 then totalmileage when sitename = 'diyitem' and ipkumdiv < 8 then totalmileage else 0 end),0) as michulgoMile" + VbCrlf
                            sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_order_master" + VbCrlf
                            sqlStr = sqlStr + "     where userid='"&userid&"' " + VbCrlf
                            sqlStr = sqlStr + "     and sitename in ('academy','diyitem')" + VbCrlf
                            sqlStr = sqlStr + "     and cancelyn='N'" + VbCrlf
                            sqlStr = sqlStr + "     and ipkumdiv>3" + VbCrlf
                            rsAcademyget.Open sqlStr,dbAcademyget,1
                            if Not rsAcademyget.Eof then
                                totmile    		= rsAcademyget("totmile")
								michulgoMile    = rsAcademyget("michulgoMile")
                            end if
                            rsAcademyget.Close



                        	'�ֹ����ϸ��� ��� ����
                            sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + VbCrlf
                            sqlStr = sqlStr + " set academymileage=" & totmile & ",michulmileACA=" & michulgoMile & VbCrlf
                            sqlStr = sqlStr + " where userid='" + CStr(userid) + "' " + VbCrlf

                            dbget.Execute sqlStr
                        END IF
                        ''''------------------------------------------------------------------------

                        sqlStr = "select top 1 userid, buyname, buyhp, buyemail from [db_academy].[dbo].tbl_academy_order_master"
                        sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
                        sqlStr = sqlStr + " and cancelyn='N'"

                        rsAcademyget.Open sqlStr,dbAcademyget,1
                        if Not rsAcademyget.Eof then
                            userid  = rsAcademyget("userid")
                        	buyname = db2html(rsAcademyget("buyname"))
                        	buyhp = db2html(rsAcademyget("buyhp"))
                        	buyemail = db2html(rsAcademyget("buyemail"))
                        end if
                        rsAcademyget.close

                        ''SMS �߼� : �����޿��� �߼��ϹǷ� �߼۾���.

                        set osms = new CSMSClass
                        osms.SendAcctIpkumOkMsgACADEMY buyhp,orderserial
                        set osms = Nothing

                        ''Email�߼�
                        ''call sendmailbankok(buyemail,buyname,orderserial)


                    end if

                    resultMSG = "OK"
                    ''if (LGD_CASSEQNO<>"001") then
                        'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������KCP]["&Left(now(),10)&"]�Ա�Ȯ�� �Ϸ�" & orderserial, resultMSG)
                    ''end if
                else

                    resultMSG = "ERR : orderserial=" & order_no & " : retval=" & retval & " : RetMsg=" & RetMsg
                    'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������KCP]["&Left(now(),10)&"]�Ա�Ȯ�� ���� �߻�" & orderserial, resultMSG)

                    if (retval=-1) then
                        ''�� ��ҵ� ������ǿ� ���� �Ա�ó�� ��û�� ���ð��.

                    end if
                end if


            ELSE
                resultMSG = "ERR"
                'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������KCP]["&Left(now(),10)&"]�Ա�Ȯ�� ERR", request("order_no") & "|" & request("tx_cd") & "|" & request("op_cd")&"|"&resultMSG&"|"&RetErr)

            END IF


        elseif (FALSE) and op_cd = "13" then
            '/*
            ' * ������ �Ա���� ���� ��� ���� ó��(DB) �κ�
            ' * ���� ��� ó���� �����̸� "OK"
            ' */
            '' ������ �ԱݿϷ���� �Ա�����ΰ�� => �Ա���������
            ''
            '' ������ �Ա�����ΰ��
            sqlStr = " select orderserial, userid, buyhp from db_academy.dbo.tbl_academy_order_master"
            sqlStr = sqlStr & " where orderserial='" + LGD_OID + "'"
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
                sqlStr = " update db_academy.dbo.tbl_academy_order_master"
                sqlStr = sqlStr & " set ipkumdiv=2"
                sqlStr = sqlStr & " , ipkumdate=NULL"
                sqlStr = sqlStr & " where orderserial='" + LGD_OID + "'"
                sqlStr = sqlStr + " and ipkumdiv=4"
                sqlStr = sqlStr + " and cancelyn='N'"

                dbget.Execute sqlStr

                ''�α� �ٽ� �� ��Ī���� ����..
                sqlStr = " update  db_academy.dbo.tbl_academy_order_CyberAccountLog"
                sqlStr = sqlStr + " set isMatched='N'"
                sqlStr = sqlStr + " where orderserial='" + LGD_OID + "'"
                dbget.Execute sqlStr


                ''�޸𳲱�.
                sqlStr = " insert into [db_academy].[dbo].tbl_academy_cs_memo"
                sqlStr = sqlStr + " (orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
                sqlStr = sqlStr + " values('" + CStr(TmpOrderserial) + "','1','" + CStr(TmpUserID) + "','0','99','','system','system','�Ա���� - ������� " & trim(request("LGD_ACCOUNTNUM")) & "," & trim(request("LGD_AMOUNT")) & " ','Y',getdate(),getdate()) "
                dbget.Execute sqlStr

                resultMSG = "OK"
                On Error Resume Next
                set osms = new CSMSClass
                osms.SendAcctIpkumCancelMsg Tmpbuyhp,LGD_OID
                set osms = Nothing
                On Error Goto 0

                'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������]["&Left(now(),10)&"]�Ա���� �Ϸ�", request("order_no") & "|" & request("tx_cd") & "|" & request("op_cd")&"|"&resultMSG&"|"&RetErr)

            else
                resultMSG = "ERR"

                'call SendMail("webserver@10x10.co.kr", "kobula@10x10.co.kr", "[�������]["&Left(now(),10)&"]�Ա�Ȯ�� ���� - ��Һ�",request("order_no") & "|" & request("tx_cd") & "|" & request("op_cd")&"|"&resultMSG&"|"&RetErr)

            end if

        end if
    else
        '//������ �����̸�
        '/*
        ' * �ŷ����� ��� ���� ó��(DB) �κ�
        ' * ������� ó���� �����̸� "OK"
        ' */
        resultMSG = "OK"
    end if

    IF resultMSG="OK" then
        Response.Write("<html><body><form><input type=""hidden"" name=""result"" value=""0000""></form></body></html>")
    else
        Response.Write("<html><body><form><input type=""hidden"" name=""result"" value=""" & resultMSG & """></form></body></html>")
    end if


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
