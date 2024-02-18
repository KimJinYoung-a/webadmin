<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs����
' History : �̻� ����
'			2018.04.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->

<!-- #include virtual="/lib/email/smslib.asp"-->
<%
dim orderserial, mode, customNumber, acctdiv, paygatetid, authcode, IsOldOrder
dim buyname, buyphone, buyhp, buyemail, accountname, checkappexists, existsorderserial
dim reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress, comment
dim cardribbon, message, fromname, yyyy, mm, dd, tt,  reqdate, reqtime, successpayetcresult
dim sqlStr
dim iAsID, divcd, reguserid, title, gubun01, gubun02, contents_jupsu, finishuser, contents_finish
dim ipkumdiv, userid, cancelyn, emailok, smsok
dim requiredetail, detailidx, jumundiv, totPaymentEtc
dim ojumun, IsAppExists, IsTempOrderExists, ErrStr, isfailorder, affectedRows


''' html2db : �Է½� ���. : 2���� Case RegCSMaster������ html2db ������� ����.
	customNumber 	= requestCheckVar(request("customNumber"),13)
    orderserial = request("orderserial")
	mode        	= requestCheckVar(request("mode"),32)
    buyname     = request("buyname")
    buyphone    = request("buyphone")
    buyhp       = request("buyhp")
    buyemail    = request("buyemail")
    accountname = request("accountname")
    reguserid   = session("ssbctid")
    reqname     = request("reqname")
    reqphone    = request("reqphone")
    reqhp       = request("reqhp")
    reqzipcode  = request("reqzipcode")
    reqzipaddr  = request("reqzipaddr")
    reqaddress  = request("reqaddress")
    comment     = request("comment")
    cardribbon  = request("cardribbon")
    message     = request("message")
    fromname    = request("fromname")
    yyyy        = request("yyyy")
    mm          = request("mm")
    dd          = request("dd")
    tt          = request("tt")

    reqdate     = yyyy + "-" + dd + "-" + dd
    reqtime     = tt

    ipkumdiv    = request("ipkumdiv")
    userid      = request("userid")
    emailok     = request("emailok")
    smsok       = request("smsok")
    requiredetail = request("requiredetail")
    detailidx     = request("detailidx")
    acctdiv 	= requestCheckVar(request("acctdiv"),3)
    paygatetid 	= requestCheckVar(request("paygatetid"),50)
    authcode 	= requestCheckVar(request("authcode"),32)
    checkappexists 	= requestCheckVar(request("checkappexists"),2)
IsOldOrder = false

dim errcode
dim osms
const CNEXT = " => "

dim myorderdetail,i

isfailorder = false

function getDateFormatedWithDash(DateVal)
	dim rtnDateStr
	dim m, d, h, Min, sec

    rtnDateStr = year(DateVal)
    m = month(DateVal)
    d = day(DateVal)
    h = Hour(DateVal)
    Min = Minute(DateVal)
    sec = second(DateVal)

    if month(DateVal)<10 then
        m = "0"&month(DateVal)
    end if

    if day(DateVal)<10 then
        d = "0"&day(DateVal)
    end if

    if Hour(DateVal)<10 then
        h = "0"&Hour(DateVal)
    end if

    if Minute(DateVal)<10 then
        Min = "0"&Minute(DateVal)
    end if

    if second(DateVal)<10 then
        sec = "0"&second(DateVal)
    end if


    rtnDateStr = rtnDateStr&"-"&m&"-"&d&" "&h&":"&Min&":"&sec
    getDateFormatedWithDash = rtnDateStr
end function

if (mode = "modifybuyerinfo") then
    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "�ֹ��� ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""


        sqlStr = " select top 1 IsNULL(buyname,'') as buyname"
        sqlStr = sqlStr + " ,IsNULL(buyphone,'') as buyphone"
        sqlStr = sqlStr + " ,IsNULL(buyhp,'') as buyhp"
        sqlStr = sqlStr + " ,IsNULL(buyemail,'') as buyemail"
        sqlStr = sqlStr + " ,IsNULL(accountname,'') as accountname"
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master "
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            contents_jupsu = contents_jupsu & "���� ����" & VbCrlf

            if (db2html(rsget("buyname"))<>buyname) then
                contents_jupsu = contents_jupsu & "�ֹ��ڸ�: " & rsget("buyname") & CNEXT & buyname & VbCrlf
            end if

            if (rsget("buyphone")<>buyphone) then
                contents_jupsu = contents_jupsu & "�ֹ�����ȭ: " & rsget("buyphone") & CNEXT & buyphone & VbCrlf
            end if

            if (rsget("buyhp")<>buyhp) then
                contents_jupsu = contents_jupsu & "�ֹ����ڵ���: " & rsget("buyhp") & CNEXT & buyhp & VbCrlf
            end if

            if (db2html(rsget("buyemail"))<>buyemail) then
                contents_jupsu = contents_jupsu & "�ֹ����̸���: " & rsget("buyemail") & CNEXT & buyemail & VbCrlf
            end if

            if (db2html(rsget("accountname"))<>accountname) then
                contents_jupsu = contents_jupsu & "�Ա��ڸ�: " & rsget("accountname") & CNEXT & accountname & VbCrlf
            end if
        end if
        rsget.Close

        contents_finish = contents_jupsu

    end if

    If Err.Number = 0 Then
        errcode = "002"


        sqlStr = " update [db_order].[dbo].tbl_order_master "     + VbCrlf
        sqlStr = sqlStr + " set buyname='" + html2db(buyname) + "' "   + VbCrlf
        sqlStr = sqlStr + " ,buyphone = '" + CStr(buyphone) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,buyhp = '" + CStr(buyhp) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,buyemail = '" + html2db(buyemail) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,accountname = '" + html2db(accountname) + "' "    + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf

        dbget.Execute sqlStr

    end if


    If Err.Number = 0 Then
        errcode = "003"
        ''html2db ������� ����.
        iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.top.listFrame.location.reload();</script>"
    response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="modifyreceiverinfo") then
    if orderserial="" or isnull(orderserial) then
        response.write "<script>alert('�ֹ���ȣ�� �����ϴ�.')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    set ojumun = new COrderMaster
        ojumun.FRectOrderSerial = orderserial
        ojumun.QuickSearchOrderMaster

        if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
            ojumun.FRectOldOrder = "on"
            ojumun.QuickSearchOrderMaster

            if (ojumun.FResultCount>0) then
                IsOldOrder = true
            end if
        end if
    set ojumun = nothing

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "������ ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""

		reqzipaddr = Replace(reqzipaddr, "��", "/")
		reqaddress = Replace(reqaddress, "��", "/")

        sqlStr = " select top 1 IsNULL(reqname,'') as reqname"
        sqlStr = sqlStr + " , IsNULL(reqphone,'') as reqphone"
        sqlStr = sqlStr + " , IsNULL(reqhp,'') as reqhp"
        sqlStr = sqlStr + " , IsNULL(reqzipcode,'') as reqzipcode"
        sqlStr = sqlStr + " , Replace(IsNULL(reqzipaddr,''), char(63), '/') as reqzipaddr"
        sqlStr = sqlStr + " , Replace(IsNULL(reqaddress,''), char(63), '/') as reqaddress"
        sqlStr = sqlStr + " ,IsNULL(comment,'') as comment"

        if IsOldOrder then
            sqlStr = sqlStr & " from db_log.dbo.tbl_old_order_master_2003 with (nolock)"
        else
            sqlStr = sqlStr & " from db_order.dbo.tbl_order_master with (nolock)"
        end if

        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            contents_jupsu = contents_jupsu & "������ ����" & VbCrlf
            if (db2html(rsget("reqname"))<>reqname) then
                contents_jupsu = contents_jupsu & "�����θ�: " & rsget("reqname") & CNEXT & reqname & VbCrlf
            end if

            if (rsget("reqphone")<>reqphone) then
                contents_jupsu = contents_jupsu & "��������ȭ: " & rsget("reqphone") & CNEXT & reqphone & VbCrlf
            end if

            if (rsget("reqhp")<>reqhp) then
                contents_jupsu = contents_jupsu & "�������ڵ���: " & rsget("reqhp") & CNEXT & reqhp & VbCrlf
            end if

            if (rsget("reqzipcode")<>reqzipcode) or (rsget("reqzipaddr")<>reqzipaddr) or (db2html(rsget("reqaddress"))<>reqaddress)  then
                if (orderserial = "21040162513") then
                    contents_jupsu = contents_jupsu & "�������ּ�: " & CNEXT & "[" & reqzipcode & "] " & reqzipaddr & " " & reqaddress & VbCrlf
                else
                    contents_jupsu = contents_jupsu & "�������ּ�: [" & rsget("reqzipcode") & "] " & rsget("reqzipaddr") & " " & rsget("reqaddress") & CNEXT & "[" & reqzipcode & "] " & reqzipaddr & " " & reqaddress & VbCrlf
                end if
            end if

            if (db2html(rsget("comment"))<>comment) then
                contents_jupsu = contents_jupsu & "��Ÿ����: " & rsget("comment") & CNEXT & comment & VbCrlf
            end if
        end if
        rsget.Close

        contents_finish = contents_jupsu
    end if

     If Err.Number = 0 Then
        errcode = "002"

        if IsOldOrder then
            sqlStr = " update db_log.dbo.tbl_old_order_master_2003" & VbCrlf
        else
            sqlStr = " update db_order.dbo.tbl_order_master" & VbCrlf
        end if

        sqlStr = sqlStr + " set reqname='" + html2db(reqname) + "' "   + VbCrlf
        sqlStr = sqlStr + " ,reqphone = '" + CStr(reqphone) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,reqhp = '" + CStr(reqhp) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,reqzipcode = '" + CStr(reqzipcode) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,reqzipaddr = '" + CStr(reqzipaddr) + "' "    + VbCrlf
        sqlStr = sqlStr + " ,reqaddress = '" + html2db(reqaddress) + "' "    + VbCrlf
        sqlStr = sqlStr & " ,comment = '" + html2db(comment) + "' where" & VbCrlf
        sqlStr = sqlStr & " orderserial='" + CStr(orderserial) + "'" & VbCrlf

        dbget.Execute sqlStr

        '' �̴Ϸ�Ż �ֹ��� ��� �̴Ͻý��� ����� ���� �� ������ ��
        If acctdiv = "150" Then
            dim xmlHttp, postdata, strData, iniMid, inimodifyAuthUrl, oJSON, resultCode
            IF application("Svr_Info")="Dev" THEN
                iniMid = "teenxtest1"
                inimodifyAuthUrl = "https://inirt.inicis.com/apis/v1/rental/modify"
                Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
            Else
                iniMid = "teenxteenr"
                inimodifyAuthUrl = "https://inirt.inicis.com/apis/v1/rental/modify"
                Set xmlHttp = CreateObject("Msxml2.ServerXMLHTTP.3.0")
            End If

            postdata = "mid="&CStr(iniMid)
            postdata = postdata&"&type=Modify"
            postdata = postdata&"&clientIp="&CStr(request.ServerVariables("LOCAL_ADDR"))
            postdata = postdata&"&"&CStr("timestamp")&"="&DateDiff("s", "1970-01-01 09:00:00", now)*1000+clng(timer)
            postdata = postdata&"&tid="&Cstr(paygatetid)
            postdata = postdata&"&recvName="&Server.URLEncode(Trim(html2db(reqname)))
            postdata = postdata&"&recvPost="&CStr(Trim(reqzipcode))
            postdata = postdata&"&recvAddr1="&Server.URLEncode(Trim(html2db(reqzipaddr)))
            postdata = postdata&"&recvAddr2="&Server.URLEncode(Trim(html2db(reqaddress)))
            postdata = postdata&"&recvTel="&CStr(replace(reqhp,"-",""))

            xmlHttp.open "POST",inimodifyAuthUrl, False
            xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset �ʿ�.
            xmlHttp.setTimeouts 90000,90000,90000,90000 ''2013/03/14 �߰�
            xmlHttp.Send postdata	'post data send
            strData = BinaryToText(xmlHttp.responseBody, "UTF-8")

            Set xmlHttp = nothing

            Set oJSON = New aspJSON
            oJSON.loadJSON(strData)
            resultCode = oJSON.data("resultCode")
            Set oJSON = Nothing
        End If
    end if

    If Err.Number = 0 Then
        errcode = "003"
        iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        ''contents_finish = Replace(contents_finish, CHAR(39), " ")
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    '// �̴Ͻý����� ���������� ����� ������ ���� �ʾ�����
    If acctdiv = "150" Then
        If resultCode<>"00" Then
            dbget.RollBackTrans
            response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + resultCode + ")')</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If
    End If

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write contents_finish & "<br />"
        ''response.write Err.line & "<br />"
        response.write Err.description & "<br />"
        ''response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    'if session("ssBctId") <> "tozzinet" and session("ssBctId") <> "icommang" then
    Call SendCsActionMail(iAsID)
	'end if

    On Error Goto 0

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.top.listFrame.location.reload();</script>"
    response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="modifyflowerinfo") then
    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "�ö�� ��� ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""


        sqlStr = " select top 1 IsNULL(cardribbon,'') as cardribbon"
        sqlStr = sqlStr + " ,IsNULL(reqdate,'') as reqdate"
        sqlStr = sqlStr + " ,IsNULL(reqtime,'') as reqtime"
        sqlStr = sqlStr + " ,IsNULL(fromname,'') as fromname"
        sqlStr = sqlStr + " ,IsNULL(message,'') as message"
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master "
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then

            contents_jupsu = contents_jupsu & "������ ����" & VbCrlf

            if (rsget("cardribbon")<>cardribbon) then
                contents_jupsu = contents_jupsu & "ī�帮��: " & getCardRibonName(rsget("cardribbon")) & CNEXT & getCardRibonName(cardribbon) & VbCrlf
            end if

            if (rsget("reqdate")<>reqdate) or (rsget("reqtime")<>reqtime) then
                contents_jupsu = contents_jupsu & "��ۿ�û��: " & rsget("reqdate") & " " & rsget("reqtime") & "~" & (rsget("reqtime")+2) & CNEXT & reqdate & " " & reqtime & "~" & (reqtime+2) & VbCrlf
            end if

            if (db2html(rsget("fromname"))<>fromname) then
                contents_jupsu = contents_jupsu & "From: " & rsget("fromname") & CNEXT & fromname & VbCrlf
            end if

            if (db2html(rsget("message"))<>message) then
                contents_jupsu = contents_jupsu & "�޼���: " & rsget("message") & CNEXT & message & VbCrlf
            end if
        end if
        rsget.Close

    end if

     If Err.Number = 0 Then
        errcode = "002"


        sqlStr = " update [db_order].[dbo].tbl_order_master "     + VbCrlf
        sqlStr = sqlStr + " set cardribbon='" + CStr(cardribbon) + "' "   + VbCrlf
        sqlStr = sqlStr + " ,reqdate = '" + yyyy + "-" + mm + "-" + dd + "' "  + VbCrlf
        sqlStr = sqlStr + " ,reqtime = '" + CStr(tt) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,fromname = '" + html2db(fromname) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,message = '" + html2db(message) + "' "    + VbCrlf
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf

        dbget.Execute sqlStr

    end if


    If Err.Number = 0 Then
        errcode = "003"
        iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    On Error Goto 0

    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script>opener.top.listFrame.location.reload();</script>"
    response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="ipkumdivnextstep") then
    if (ipkumdiv="2") then
        divcd   = "A900"
        title   = "�����Ϸ� ��������"
        gubun01 = "C004"
        gubun02 = "CD99"

        ''�޸�� �Է��ϰ� ����

        sqlStr = "select top 1 userid, buyname, buyhp, buyemail, cancelyn, jumundiv "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master"
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            userid      = rsget("userid")
        	buyname     = db2html(rsget("buyname"))
        	buyhp       = db2html(rsget("buyhp"))
        	buyemail    = db2html(rsget("buyemail"))
        	cancelyn    = rsget("cancelyn")
        	jumundiv    = db2html(rsget("jumundiv"))
        end if
        rsget.close

        if (cancelyn="") or (cancelyn="Y") or (cancelyn="D") then
            response.write "<script>alert('���� �ֹ����� �ƴϰų� ��ҵ� �����Դϴ�.')</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

        On Error Resume Next
        dbget.beginTrans

        If Err.Number = 0 Then
            errcode = "001"

            sqlStr =	"update [db_order].[dbo].tbl_order_master " & vbCrlf
    		sqlStr = sqlStr & " set ipkumdiv='4'" & vbCrlf
    		sqlStr = sqlStr & " ,ipkumdate=getdate()" & vbCrlf
    		sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    		dbget.Execute sqlStr

    		''��� ������Ʈ
    		sqlStr = " exec db_summary.dbo.sp_ten_RealtimeStock_regIpkum '" & orderserial & "'"
    		dbget.Execute sqlStr
	    end IF

	    If Err.Number = 0 Then
            errcode = "002"
		    '' �ֹ� ���ϸ��� ������Ʈ
		    CALL updateUserMileage(userid)
		end IF

		If Err.Number = 0 Then
            errcode = "005"
            call AddCsMemo(orderserial,"1",userid,reguserid,"�����Ϸ� ��������")
        end if


		If Err.Number = 0 Then
            dbget.CommitTrans

            ''2015/08/17 �߰�
            if (smsok<>"") Then
                'set osms = new CSMSClass
                '    osms.SendAcctIpkumOkMsg buyhp,orderserial
                'set osms = Nothing

                Call SendNormalSMS_LINK(buyhp,"1644-6030","[�ٹ�����]�Ա�Ȯ�� �Ǿ����ϴ�. �ֹ���ȣ�� " + orderserial + "�Դϴ�.�����մϴ�.")
            end if

            if (emailok<>"") Then
                IF (jumundiv="7") or (jumundiv="4") then
		            Call sendmailbankokNoDLV(buyemail,buyname,orderserial)
		        ELSE
		            Call SendMailBankOk(buyemail,buyname,orderserial)
		        END IF
            end if
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If


        response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
        response.write "<script>opener.top.listFrame.location.reload();</script>"
        response.write "<script>opener.top.detailFrame.location.reload();</script>"
        response.write "<script>opener.focus(); window.close();</script>"
        dbget.close()	:	response.End

    ' ���� �ֹ� ���� �ֹ����� ����
	elseif (ipkumdiv="1") or (ipkumdiv="0") then
        ' ������ ��Ʈ�� �̰ų�, ������ ����
		if Not(C_CSPowerUser or C_ADMIN_AUTH) then
			response.write "<script>alert('������ �����ϴ�.');</script>"
			response.write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if
        orderserial=trim(orderserial)
		if orderserial="" or isnull(orderserial) then
			response.write "<script>alert('�ֹ���ȣ�� �����ϴ�.');</script>"
			response.write "<script>history.back()</script>"
			dbget.close()	:	response.End
		end if

		set ojumun = new COrderMaster
		ojumun.FRectOrderSerial = orderserial
		ojumun.QuickSearchOrderMaster

        IF ojumun.FTotalCount < 1 THEN
            response.write "<script>alert('�������� �ֹ����� �ƴմϴ�.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        IF ojumun.FOneItem.Fcancelyn<>"N" THEN
            response.write "<script>alert('��ҵ� ���� �Դϴ�.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

		IsAppExists = False
		IsTempOrderExists = False
        '// ���γ��� �ִ���
        IsAppExists = ojumun.getAppLogExists()
        IsTempOrderExists = ojumun.getTempOrderExists()

		if (ojumun.FOneItem.Fipkumdiv = "1") or (ojumun.FOneItem.Fipkumdiv = "0") then
            isfailorder=true
		end if
        IF not(isfailorder) THEN
            response.write "<script>alert('���е� �ֹ��� �ƴմϴ�.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        if lcase(checkappexists)="on" then
            if Not(IsAppExists) then
                if datediff("d",left(ojumun.FOneItem.Fregdate,10),date())>2 then
                    response.write "<script>alert('�������� 2���� �ʰ� �Ǿ�����, PG�翡�� �Ѿ�� ���γ����� ���� �ֹ� �Դϴ�.\n�������� ���� �Ͻðų�, PG����γ���üũ�� ������ �ּ���.');</script>"
                    response.write "<script>history.back()</script>"
                    dbget.close()	:	response.End
                end if
            end if
        end if
        acctdiv=trim(getNumeric(acctdiv))
        IF acctdiv="" or isnull(acctdiv) THEN
            response.write "<script>alert('��������� ������ �ּ���.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        IF not(acctdiv="100" or acctdiv="20" or acctdiv="400" or acctdiv="7") THEN
            response.write "<script>alert('��������� �ſ�ī��,�ǽð���ü,�ڵ�������,������ �� ���ð��� �մϴ�.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        paygatetid=trim(paygatetid)
        IF paygatetid="" or isnull(paygatetid) THEN
            response.write "<script>alert('PG�� TID�� �Է��� �ּ���.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        IF len(paygatetid)<10 THEN
            response.write "<script>alert('�������� PG�� TID�� �ƴմϴ�.');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if
        authcode=trim(getNumeric(authcode))
        if authcode<>"" then
            IF len(authcode)>10 THEN
                response.write "<script>alert('�������� ���ι�ȣ�� �ƴմϴ�.');</script>"
                response.write "<script>history.back()</script>"
                dbget.close()	:	response.End
            end if
        end if

        existsorderserial=""
		sqlStr = " select top 1 m.orderserial" & vbCrLf
		sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (readuncommitted)" & vbCrLf
		sqlStr = sqlStr & " where m.orderserial<>'" & orderserial & "'" & vbCrLf
        sqlStr = sqlStr & " and m.paygatetid='" & paygatetid & "'" & vbCrLf

        'response.write sqlStr & "<br>"
    	rsget.Open sqlStr,dbget,1
    	if not rsget.Eof Then
    		existsorderserial = rsget("orderserial")
    	end if
    	rsget.close
        IF existsorderserial<>"" THEN
            response.write "<script>alert('�ٸ��ֹ�("& existsorderserial &")���� �̹� ������ TID �Դϴ�[0].');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

        existsorderserial=""
		sqlStr = " select top 1 t.orderserial" & vbCrLf
		sqlStr = sqlStr & " from db_order.dbo.tbl_order_temp t with (readuncommitted)" & vbCrLf
		sqlStr = sqlStr & " where t.orderserial<>'" & orderserial & "'" & vbCrLf
        sqlStr = sqlStr & " and t.p_tid='" & paygatetid & "'" & vbCrLf

        'response.write sqlStr & "<br>"
    	rsget.Open sqlStr,dbget,1
    	if not rsget.Eof Then
    		existsorderserial = rsget("orderserial")
    	end if
    	rsget.close
        IF existsorderserial<>"" THEN
            response.write "<script>alert('�ٸ��ֹ�("& existsorderserial &")���� �̹� ������ TID �Դϴ�[1].');</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        end if

        'Ʈ�����ǽ���
		dbget.BeginTrans

		sqlStr = " update m " & vbCrLf
		sqlStr = sqlStr & " set m.ipkumdate=l.appdate" & vbCrLf
		sqlStr = sqlStr & " from db_order.[dbo].[tbl_order_master] m with (readuncommitted)" & vbCrLf
		sqlStr = sqlStr & " join db_order.dbo.tbl_onlineApp_log l with (readuncommitted)" & vbCrLf
        sqlStr = sqlStr & "     on m.paygateTid=l.PGkey " & vbCrLf
		sqlStr = sqlStr & " where m.orderserial = '" & CStr(orderserial) & "'" & vbCrLf
		sqlStr = sqlStr & " and l.orderserial is NULL " & vbCrLf
		sqlStr = sqlStr & " and l.appDivCode = 'A' " & vbCrLf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

        IF (Err) then
		    ErrStr = "[Err-ORD-000]" & Err.Description
		    dbget.RollBackTrans
			response.write ErrStr
		    dbget.close()	:	response.End
		end if

        sqlStr = "update m set" & vbCrLf
        sqlStr = sqlStr & " m.ipkumdiv=4, m.ipkumdate=(case when m.ipkumdate is null then m.regdate else m.ipkumdate end)" & vbCrLf
        sqlStr = sqlStr & " , m.paygatetid='"& paygatetid &"' , m.authcode='"& authcode &"'" & vbCrLf
        sqlStr = sqlStr & " , m.totalvat = t.itemvat" & vbCrLf
        sqlStr = sqlStr & " , m.totalmileage = t.mileage" & vbCrLf
        sqlStr = sqlStr & " , m.totalsum = t.totalsum" & vbCrLf
        sqlStr = sqlStr & " , m.subtotalpricecouponnotapplied = t.totalsum" & vbCrLf
        sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (readuncommitted)" & vbCrLf
        sqlStr = sqlStr & " join (" & vbCrLf
        sqlStr = sqlStr & " 	select orderserial" & vbCrLf
        sqlStr = sqlStr & " 	, sum(itemvat) as itemvat, sum(mileage) as mileage, sum(itemcost*itemno) as totalsum, sum(reducedprice*itemno) as reducedprice" & vbCrLf
        sqlStr = sqlStr & " 	from db_order.dbo.tbl_order_detail with (readuncommitted)" & vbCrLf
        sqlStr = sqlStr & " 	where orderserial='"& orderserial &"'" & vbCrLf
        sqlStr = sqlStr & " 	and cancelyn<>'Y'" & vbCrLf
        sqlStr = sqlStr & " 	group by orderserial" & vbCrLf
        sqlStr = sqlStr & " ) as t" & vbCrLf
        sqlStr = sqlStr & " 	on m.orderserial = t.orderserial" & vbCrLf
        sqlStr = sqlStr & " where m.orderserial = '"& orderserial &"'" & vbCrLf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

        IF (Err) then
		    ErrStr = "[Err-ORD-001]" & Err.Description
		    dbget.RollBackTrans
			response.write ErrStr
		    dbget.close()	:	response.End
		end if

		sqlStr = " update m" + vbCrLf
		sqlStr = sqlStr & " set m.accountdiv='" & acctdiv & "'" & vbCrLf
		sqlStr = sqlStr & " from db_order.[dbo].[tbl_order_master] m " & vbCrLf
		sqlStr = sqlStr & " where m.orderserial = '" + CStr(orderserial) + "' and m.accountdiv<>'" & acctdiv & "'" & vbCrLf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

        IF (Err) then
		    ErrStr = "[Err-ORD-002]" & Err.Description
		    dbget.RollBackTrans
			response.write ErrStr
		    dbget.close()	:	response.End
		end if

		'// �ֹ����� �ٽ� �б�(�Ա��� ��������)
		ojumun.QuickSearchOrderMaster

        if ojumun.FOneItem.FPgGubun<>"" then
            ' pg�纰 �ֱ� ���� �ڵ带 �޾ƿ´�.
            sqlStr = " select top 1 e.payetcresult" & vbCrLf
            sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (readuncommitted)" & vbCrLf
            sqlStr = sqlStr & " join db_order.dbo.tbl_order_PaymentEtc e with (readuncommitted)" & vbCrLf
            sqlStr = sqlStr & "     on m.orderserial=e.orderserial and m.accountdiv=e.acctdiv" & vbCrLf
            sqlStr = sqlStr & " where m.accountdiv = '"& ojumun.FOneItem.Faccountdiv &"'" & vbCrLf
            sqlStr = sqlStr & " and m.pggubun = '"& ojumun.FOneItem.FPgGubun &"'" & vbCrLf
            sqlStr = sqlStr & " and m.ipkumdiv>3 and m.cancelyn='N'" & vbCrLf
            sqlStr = sqlStr & " order by m.orderserial desc" & vbCrLf

            'response.write sqlStr & "<br>"
            rsget.Open sqlStr,dbget,1
            if not rsget.Eof Then
                successpayetcresult = rsget("payetcresult")
            end if
            rsget.close

            IF (Err) then
                ErrStr = "[Err-ORD-003]" & Err.Description
                dbget.RollBackTrans
                response.write ErrStr
                dbget.close()	:	response.End
            end if
        end if

		''########## ��븶�ϸ��� �α� ##########
		if ojumun.FOneItem.Fmiletotalprice > 0 and ojumun.FOneItem.Fuserid <> "" then
			sqlStr = "insert into [db_user].[dbo].tbl_mileagelog(userid,mileage,jukyocd,jukyo,orderserial, regdate)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(ojumun.FOneItem.Fuserid) + "'," + CStr(-1*CLng(ojumun.FOneItem.Fmiletotalprice)) + ",'02','��ǰ����','" + orderserial + "', '" & getDateFormatedWithDash(ojumun.FOneItem.Fipkumdate) & "')"
			dbget.Execute(sqlStr)

			sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + vbCrlf
			sqlStr = sqlStr + " set spendmileage=spendmileage + " + CStr(ojumun.FOneItem.Fmiletotalprice) + vbCrlf
			sqlStr = sqlStr + " where userid='" + CStr(ojumun.FOneItem.Fuserid) + "'"

			dbget.Execute(sqlStr)

			IF (Err) then
    		    ErrStr = "[Err-ORD-004]" & Err.Description
    		    dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
    		end if
		end if

		sqlStr = " if not exists(select top 1 orderserial from [db_order].[dbo].[tbl_order_temp] with (readuncommitted) where orderserial = '" + CStr(orderserial) + "') " + vbCrLf
		sqlStr = sqlStr + " begin " + vbCrLf
		sqlStr = sqlStr + " 	update t " + vbCrLf
		sqlStr = sqlStr + " 	set t.orderserial = m.orderserial " + vbCrLf
		sqlStr = sqlStr + " 	from " + vbCrLf
		sqlStr = sqlStr + " 		[db_order].[dbo].[tbl_order_master] m " + vbCrLf
		sqlStr = sqlStr + " 		join [db_order].[dbo].[tbl_order_temp] t on m.paygateTid = t.P_TID " + vbCrLf
		sqlStr = sqlStr + " 	where " + vbCrLf
		sqlStr = sqlStr + " 		m.orderserial = '" + CStr(orderserial) + "' and t.orderserial = '' " + vbCrLf
		sqlStr = sqlStr + " end " + vbCrLf
		dbget.Execute(sqlStr)

        if ojumun.FOneItem.FPgGubun="NP" then
		    sqlStr = " update T " + vbCrLf
		    sqlStr = sqlStr & " set T.Tn_paymethod='" & acctdiv & "'" & vbCrLf
		    sqlStr = sqlStr & " from [db_order].[dbo].[tbl_order_temp] T " & vbCrLf
		    sqlStr = sqlStr & " where T.orderserial = '" + CStr(orderserial) + "' and T.Tn_paymethod = '900' " & vbCrLf
		    'response.write sqlStr & "<br>"
		    dbget.Execute sqlStr
        end if

		sqlStr = " select top 1 spendtencash, spendgiftmoney, IsNull(pDiscount,0) as pDiscount, IsNull(pDiscount2, 0) as pDiscount2, tn_paymethod " + vbCrLf
		sqlStr = sqlStr + " from " + vbCrLf
		sqlStr = sqlStr + " [db_order].[dbo].[tbl_order_temp] with (readuncommitted)" + vbCrLf
		sqlStr = sqlStr + " where orderserial = '" + orderserial + "' " + vbCrLf
    	rsget.Open sqlStr,dbget,1
    	if not rsget.Eof Then
    		ojumun.FOneItem.FspendTenCash = rsget("spendTenCash")
    		ojumun.FOneItem.Fspendgiftmoney = rsget("spendgiftmoney")
			ojumun.FOneItem.FpDiscount = rsget("pDiscount")
			ojumun.FOneItem.FpDiscount2 = rsget("pDiscount2")
			ojumun.FOneItem.Faccountdiv = rsget("tn_paymethod")
		else
			ojumun.FOneItem.FspendTenCash = 0
    		ojumun.FOneItem.Fspendgiftmoney = 0
			ojumun.FOneItem.FpDiscount = 0
			ojumun.FOneItem.FpDiscount2 = 0
    	end if
    	rsget.close

		''########## ��뿹ġ�� �α� ##########
		if ojumun.FOneItem.Fspendtencash > 0 and ojumun.FOneItem.Fuserid <> "" then
			sqlStr = "insert into [db_user].[dbo].tbl_depositlog(userid,deposit,jukyocd,jukyo,orderserial, regdate)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(ojumun.FOneItem.Fuserid) + "'," + CStr(-1*CLng(ojumun.FOneItem.Fspendtencash)) + ",100,'��ǰ����','" + orderserial + "', '" & getDateFormatedWithDash(ojumun.FOneItem.Fipkumdate) & "')"
			dbget.Execute(sqlStr)

			sqlStr = "update [db_user].[dbo].tbl_user_current_deposit" + vbCrlf
			sqlStr = sqlStr + " set spenddeposit=spenddeposit + " + CStr(ojumun.FOneItem.Fspendtencash) + vbCrlf
			sqlStr = sqlStr + " ,currentdeposit=currentdeposit - " + CStr(ojumun.FOneItem.Fspendtencash) + vbCrlf   '''+-Ȯ��.
			sqlStr = sqlStr + " where userid='" + CStr(ojumun.FOneItem.Fuserid) + "'"

			dbget.Execute(sqlStr)

			sqlStr = " insert into db_order.dbo.tbl_order_PaymentEtc"
			sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate)"
			sqlStr = sqlStr + " values('"&orderserial&"'"
			sqlStr = sqlStr + " ,'200'"
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.Fspendtencash&""
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.Fspendtencash&""
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " )"

			dbget.Execute sqlStr

			IF (Err) then
    		    ErrStr = "[Err-ORD-005]" & Err.Description
    		    dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
    		end if
		end if

		''########## ���Giftī�� �α� ##########
		if ojumun.FOneItem.Fspendgiftmoney > 0 and ojumun.FOneItem.Fuserid <> "" then
			sqlStr = "insert into [db_user].[dbo].tbl_depositlog(userid,deposit,jukyocd,jukyo,orderserial, regdate)" + vbCrlf
			sqlStr = sqlStr + " values('" + CStr(ojumun.FOneItem.Fuserid) + "'," + CStr(-1*CLng(ojumun.FOneItem.Fspendtencash)) + ",100,'��ǰ����','" + orderserial + "', '" & getDateFormatedWithDash(ojumun.FOneItem.Fipkumdate) & "')"
			dbget.Execute(sqlStr)

			sqlStr = "update [db_user].[dbo].tbl_user_current_deposit" + vbCrlf
			sqlStr = sqlStr + " set spenddeposit=spenddeposit + " + CStr(ojumun.FOneItem.Fspendtencash) + vbCrlf
			sqlStr = sqlStr + " ,currentdeposit=currentdeposit - " + CStr(ojumun.FOneItem.Fspendtencash) + vbCrlf   '''+-Ȯ��.
			sqlStr = sqlStr + " where userid='" + CStr(ojumun.FOneItem.Fuserid) + "'"

			dbget.Execute(sqlStr)

			sqlStr = " insert into db_order.dbo.tbl_order_PaymentEtc"
			sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate)"
			sqlStr = sqlStr + " values('"&orderserial&"'"
			sqlStr = sqlStr + " ,'900'"
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.Fspendgiftmoney&""
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.Fspendgiftmoney&""
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " )"

			dbget.Execute sqlStr

			IF (Err) then
    		    ErrStr = "[Err-ORD-006]" & Err.Description
    		    dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
    		end if
		end if

		''########## �ֹ����ϸ��� ���� ##########
		if (ojumun.FOneItem.Fuserid<>"") and (ojumun.FOneItem.Fsitename="10x10") then
			''## �ֹ� ���ϸ��� ������Ʈ ##''
			sqlStr = "update [db_user].[dbo].tbl_user_current_mileage" + VbCrlf
			sqlStr = sqlStr + " set jumunmileage=jumunmileage+" + CStr(ojumun.FOneItem.Ftotalmileage) + VbCrlf
			sqlStr = sqlStr + " ,michulmile=michulmile+" + CStr(ojumun.FOneItem.Ftotalmileage) + VbCrlf  ''2015/03/06 �߰�
			sqlStr = sqlStr + " where userid='" + CStr(ojumun.FOneItem.Fuserid) + "'"

			dbget.Execute(sqlStr)

			IF (Err) then
    		    ErrStr = "[Err-ORD-007]" & Err.Description
    		    dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
    		end if
		end if

		sqlStr = " select IsNull(sum(acctamount), 0) as totPaymentEtc from [db_order].[dbo].[tbl_order_PaymentEtc] with (readuncommitted) where orderserial = '"&orderserial&"' "
    	rsget.Open sqlStr,dbget,1
    	if not rsget.Eof Then
    		totPaymentEtc = rsget("totPaymentEtc")
		else
			totPaymentEtc = 0
    	end if
    	rsget.close

		'// /www/lib/classes/ordercls/shoppingbagDBcls.asp
		'// SaveOrderResultDB
		if ojumun.FOneItem.FPgGubun="NP" then
			'�ְ��� ���� ����
			sqlStr = " insert into db_order.dbo.tbl_order_PaymentEtc"
			sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate,PayEtcResult,pDiscount)"
			sqlStr = sqlStr + " values('"&orderserial&"'"
			sqlStr = sqlStr + " ,'"& ojumun.FOneItem.Faccountdiv &"'"
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FpDiscount-totPaymentEtc&""
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FpDiscount-totPaymentEtc&""
			sqlStr = sqlStr + " ,convert(varchar(32),'" & authcode & "')"
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " ,'"& successpayetcresult &"',0"
			sqlStr = sqlStr + " );" & vbCrLf

			'���̹�����Ʈ ���� ���� (���̹�����Ʈ: 120)
			if ojumun.FOneItem.FpDiscount>0 then
				sqlStr = sqlStr + " insert into db_order.dbo.tbl_order_PaymentEtc"
				sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate,PayEtcResult,pDiscount)"
				sqlStr = sqlStr + " values('"&orderserial&"'"
				sqlStr = sqlStr + " ,'120'"
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FpDiscount&""
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FpDiscount&""
				sqlStr = sqlStr + " ,convert(varchar(32),'" + ojumun.FOneItem.Fauthcode + "')"
				sqlStr = sqlStr + " ,'','',0"
				sqlStr = sqlStr + " )"
			end If

		elseif ojumun.FOneItem.FPgGubun="PY" then
			sqlStr = " insert into db_order.dbo.tbl_order_PaymentEtc"
			sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate,PayEtcResult,pDiscount, pAddParam)"
			sqlStr = sqlStr + " values('"&orderserial&"'"
			sqlStr = sqlStr + " ,'"& ojumun.FOneItem.Faccountdiv &"'"
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FpDiscount2-totPaymentEtc&""
			sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FpDiscount2-totPaymentEtc&""
			sqlStr = sqlStr + " ,convert(varchar(32),'" & authcode & "')"
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " ,'"& successpayetcresult &"'"
			sqlStr = sqlStr + " ,'"&ojumun.FOneItem.FpDiscount&"'"
			sqlStr = sqlStr + " ,''"
			sqlStr = sqlStr + " );" & vbCrLf

			'����������Ʈ ���� ���� (����������Ʈ: 120)
			if ojumun.FOneItem.FpDiscount2>0 then
				sqlStr = sqlStr + " insert into db_order.dbo.tbl_order_PaymentEtc"
				sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate,PayEtcResult,pDiscount, pAddParam)"
				sqlStr = sqlStr + " values('"&orderserial&"'"
				sqlStr = sqlStr + " ,'120'"
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FpDiscount2&""
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FpDiscount2&""
				sqlStr = sqlStr + " ,convert(varchar(32),'" + ojumun.FOneItem.Fauthcode + "')"
				sqlStr = sqlStr + " ,'','',0"
				sqlStr = sqlStr + " ,''"
				sqlStr = sqlStr + " )"
			end If
		else
			'// �Ϲ� ������ ó��
			sqlStr = " insert into db_order.dbo.tbl_order_PaymentEtc"
			sqlStr = sqlStr + " (orderserial,acctdiv,acctamount,realPayedSum,acctAuthCode,acctAuthDate,pDiscount)"
			sqlStr = sqlStr + " values('"&orderserial&"'"
			IF (ojumun.FOneItem.Faccountdiv="110") THEN  ''�ſ�+OK ����
    			ErrStr = "[Err-ORD-008] �ſ�+OK ���� ó���ȵ�"
    			dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
				sqlStr = sqlStr + " ,'100'"
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FOKCashbagSpend&""
			ELSE
				sqlStr = sqlStr + " ,'"&ojumun.FOneItem.Faccountdiv&"'"
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice&""
			ENd IF

			IF (ojumun.FOneItem.Faccountdiv="110") THEN
    			ErrStr = "[Err-ORD-009] �ſ�+OK ���� ó���ȵ�"
    			dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice-ojumun.FOneItem.FOKCashbagSpend&""
			ELSE
				sqlStr = sqlStr + " ,"&ojumun.FOneItem.FSubtotalPrice&""
			ENd IF

			sqlStr = sqlStr + " ,convert(varchar(32),'" & authcode & "')"
			sqlStr = sqlStr + " ,''"
			''sqlStr = sqlStr + " ,'"&ojumun.FOneItem.FPayEtcResult&"'"
			sqlStr = sqlStr + " ,'"&ojumun.FOneItem.FpDiscount&"'"
			sqlStr = sqlStr + " )"
		end if

		''response.write sqlStr
        dbget.Execute sqlStr

        IF (Err) then
    		ErrStr = "[Err-ORD-010]" & Err.Description
    		dbget.RollBackTrans
			response.write ErrStr
			dbget.close()	:	response.End
    	end if

		if ((CLng(ojumun.FOneItem.Fspendtencash)>0) or (CLng(ojumun.FOneItem.Fspendgiftmoney)>0)) then    ''��Ÿ������ �հ�.
		    sqlStr = " update M "
            sqlStr = sqlStr + " set M.sumPaymentEtc=IsNULL("
            sqlStr = sqlStr + " 		(select sum(acctamount) as totamount "
            sqlStr = sqlStr + " 		from db_order.dbo.tbl_order_PaymentEtc "
            sqlStr = sqlStr + " 		where orderserial='"&orderserial&"' and acctdiv in ('200','900')),0)"
            sqlStr = sqlStr + " from db_order.dbo.tbl_order_master M with (readuncommitted)"
            sqlStr = sqlStr + " where M.orderserial='"&orderserial&"'"

            dbget.Execute sqlStr

		    IF (Err) then
    		    ErrStr = "[Err-ORD-011]" & Err.Description
    		    dbget.RollBackTrans
				response.write ErrStr
				dbget.close()	:	response.End
    		end if
	    end if

		sqlStr = " update l " & vbCrLf
		sqlStr = sqlStr & " set l.orderserial = m.orderserial" & vbCrLf
		sqlStr = sqlStr & " from db_order.[dbo].[tbl_order_master] m" & vbCrLf
		sqlStr = sqlStr & " join db_order.dbo.tbl_onlineApp_log l" & vbCrLf
        sqlStr = sqlStr & "     on m.paygateTid=l.PGkey " & vbCrLf
		sqlStr = sqlStr & " where m.orderserial = '" & CStr(orderserial) & "'" & vbCrLf
		sqlStr = sqlStr & " and l.orderserial is NULL " & vbCrLf
		sqlStr = sqlStr & " and l.appDivCode = 'A' " & vbCrLf

		'response.write sqlStr & "<br>"
		dbget.Execute sqlStr

		If Err.Number = 0 Then
            call AddCsMemo(orderserial,"1",ojumun.FOneItem.Fuserid,reguserid,"��� ������ ���� �������а� -> �����Ϸ� ó��")
        end if

        IF (Err) then
		    ErrStr = "[Err-ORD-012]" &Err.Description
		    dbget.RollBackTrans
			response.write ErrStr
		    dbget.close()	:	response.End
		ELSE
		    dbget.CommitTrans
		end if
    else
        response.write "<script>alert('�Ա� ��� ���¿����� ���� ���·� ���� �����մϴ�.');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

	response.write "OK"

elseif (mode="edithandmadereq") then


    set myorderdetail = new COrderMaster
    myorderdetail.FRectOrderserial = orderserial
    myorderdetail.FRectDetailIdx = detailidx
    myorderdetail.GetOneOrderDetail


    ''������ ������ �϶�
    if (myorderdetail.FOneItem.FItemNo>1) then
        requiredetail = ""
        for i=0 to myorderdetail.FOneItem.FItemNo-1
            if (request.form("requiredetail"&i)<>"") then
                requiredetail = requiredetail & request.form("requiredetail"&i) & CAddDetailSpliter
            end if
        next

        if Right(requiredetail,2)=CAddDetailSpliter then
            requiredetail = Left(requiredetail,Len(requiredetail)-2)
        end if
    end if
    set myorderdetail = Nothing

    On Error Resume Next
    dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd       = "A900"
        title       = "�ֹ����� ��ǰ ���� ����"
        gubun01     = "C004"
        gubun02     = "CD99"

        contents_jupsu  = ""
        contents_finish = ""
        finishuser      = CFINISH_SYSTEM

        sqlStr = " select top 1 IsNULL(requiredetail,'') as requiredetail"
        sqlStr = sqlStr + " ,IsNULL(itemname,'') as itemname"
        sqlStr = sqlStr + " ,IsNULL(itemoptionname,'') as itemoptionname"
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail "
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf
        sqlStr = sqlStr + " and idx=" + CStr(detailidx)


        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then

            contents_jupsu = contents_jupsu & "���� ����" & VbCrlf

            if (db2html(rsget("requiredetail"))<>requiredetail) then
                contents_jupsu = contents_jupsu & "��ǰ��(�ɼ�): " & db2html(rsget("itemname"))
                if (rsget("itemoptionname")<>"") then
                    contents_jupsu = contents_jupsu & "(" & db2html(rsget("itemoptionname")) & ")" & VbCrlf
                end if
                contents_jupsu = contents_jupsu & "����: " & rsget("requiredetail") & VbCrlf & CNEXT & VbCrlf & requiredetail & VbCrlf
            end if

        end if
        rsget.Close

    end if

    if (contents_jupsu="") then
        response.write "<script language='javascript'>alert('�����Ͻ� ������ ���� ������ ��ġ�մϴ�. �������� �ʾҽ��ϴ�.');</script>"
        response.write "<script language='javascript'>history.back();</script>"
        dbget.close()	:	response.End
    else
        contents_jupsu = "���� ����" & VbCrlf & contents_jupsu
        contents_finish = contents_jupsu
    end if

    If Err.Number = 0 Then
        errcode = "002"

        sqlStr = "update [db_order].[dbo].tbl_order_detail" + VbCrlf
        sqlStr = sqlStr + " set requiredetail='" + html2db(requiredetail) + "'" + VbCrlf
        sqlStr = sqlStr + " where idx=" + CStr(detailidx)

        dbget.Execute sqlStr
    end if


    If Err.Number = 0 Then
        errcode = "003"
        '' html2db ������� ����.
        iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"
        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

    end if

    If Err.Number = 0 Then
        dbget.CommitTrans

        response.write "<script language='javascript'>alert('���� �Ǿ����ϴ�.');</script>"
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>window.close();</script>"
        dbget.close()	:	response.End
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If
    On Error Goto 0

elseif (mode="editforeigndirectpurchase") then
	If customNumber = "" Or orderserial = "" Then
		Response.Write "<script type='text/javascript'>alert('�߸��� �����Դϴ�.'); history.back();</script>"
		dbget.close(): Response.End
	End IF

	dbget.beginTrans

    If Err.Number = 0 Then
        errcode = "001"

        divcd   = "A900"
        title   = "�ؿ����� ���� ����"
        gubun01 = "C004"
        gubun02 = "CD99"

        contents_jupsu = ""
        finishuser      = reguserid
        contents_finish = ""

        sqlStr = " select top 1 IsNULL(customnumber,'') as customnumber"
        sqlStr = sqlStr + " from db_order.dbo.tbl_order_custom_number"
        sqlStr = sqlStr + " where orderserial='" + CStr(orderserial) + "' " + VbCrlf
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            contents_jupsu = contents_jupsu & "������ ����" & VbCrlf
            if (rsget("customnumber")<>customnumber) then
                contents_jupsu = contents_jupsu & "������� ������ȣ: " & rsget("customnumber") & CNEXT & customnumber & VbCrlf
            end if
        end if
        rsget.Close

        contents_finish = contents_jupsu
    end if

	If Err.Number = 0 Then
		errcode = "002"

		'// ������� ������ȣ ����
		sqlStr = "UPDATE db_order.dbo.tbl_order_custom_number SET customnumber = '"& customNumber &"', lastupdate = getdate()  WHERE orderserial = '" & orderserial & "' "
		dbget.Execute sqlStr
	end if

    If Err.Number = 0 Then
        errcode = "003"

        iAsID = RegCSMaster(divcd , orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
    end if

    If Err.Number = 0 Then
        errcode = "004"

        Call FinishCSMaster(iAsid, finishuser, html2db(contents_finish))

        Call AddCustomerOpenContents(iAsid, html2db(contents_finish))
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script>alert('����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + ")')</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail(iAsID)

    response.write "<script type='text/javascript'>"
    response.write "	alert('���� �Ǿ����ϴ�.');"
    response.write "	opener.top.listFrame.location.reload();"
    response.write "	opener.top.detailFrame.location.reload();"
    response.write "	opener.focus();"
    response.write "	window.close();"
    response.write "</script>"
    dbget.close()	:	response.End

elseif (mode="chgtoextordr") then

    sqlStr = " update m "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	m.jumundiv = o.jumundiv, m.accountdiv = o.accountdiv, m.accountno = o.accountno "
    sqlStr = sqlStr + " 	, m.beadaldiv = o.beadaldiv, m.sitename = o.sitename, m.paygatetid = o.paygatetid "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] o "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		m.linkorderserial = o.orderserial "
    sqlStr = sqlStr + " where m.orderserial = '" & orderserial & "' and m.sitename = '10x10' and o.jumundiv = 5 "
    dbget.Execute sqlStr, affectedRows

    if (affectedRows > 0) then
        call AddCsMemo(orderserial,"1","",reguserid,"���޸� �ֹ���ȯ")

        response.write "<script type='text/javascript'>"
        response.write "	alert('���� �Ǿ����ϴ�.');"
        response.write "	opener.focus();"
        response.write "	window.close();"
        response.write "</script>"
    else
        response.write "���� : ��ȯ���� �ʾҽ��ϴ�<br />(���ֹ��� �����ֹ��� �ƴմϴ�.)"
    end if
    dbget.close()	:	response.End
elseif (mode="chgtotenordr") then

    sqlStr = " update m "
    sqlStr = sqlStr + " set "
    sqlStr = sqlStr + " 	m.jumundiv = '1', m.accountdiv = '7' "
    sqlStr = sqlStr + " 	, m.beadaldiv = '1', m.sitename = '10x10' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_master] m "
    sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] o "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		m.linkorderserial = o.orderserial "
    sqlStr = sqlStr + " where m.orderserial = '" & orderserial & "' and m.sitename <> '10x10' and o.jumundiv = 5 "
    dbget.Execute sqlStr, affectedRows

    if (affectedRows > 0) then
        call AddCsMemo(orderserial,"1","",reguserid,"���� �ֹ���ȯ")

        response.write "<script type='text/javascript'>"
        response.write "	alert('���� �Ǿ����ϴ�.');"
        response.write "	opener.focus();"
        response.write "	window.close();"
        response.write "</script>"
    else
        response.write "���� : ��ȯ���� �ʾҽ��ϴ�<br />(���ֹ��� �����ֹ��� �ƴմϴ�.)"
    end if
    dbget.close()	:	response.End
Else
    response.write "<script type='text/javascript'>"
    response.write "	alert('�����ڰ� �����ϴ�.')"
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if



%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
