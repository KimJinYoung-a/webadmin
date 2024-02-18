<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/smscls.asp" -->
<%


dim orderserial, mode
dim buyname, buyphone, buyhp, buyemail, accountname
dim reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress, comment
dim cardribbon, message, fromname, yyyy, mm, dd, tt,  reqdate, reqtime
dim sqlStr
dim iAsID, divcd, reguserid, title, gubun01, gubun02, contents_jupsu, finishuser, contents_finish
dim ipkumdiv, userid, cancelyn, emailok, smsok
dim requiredetail, detailidx

''' html2db : �Է½� ���. : 2���� Case RegCSMaster������ html2db ������� ����.


orderserial = request("orderserial")
mode        = request("mode")

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

dim errcode
dim osms
const CNEXT = " => "

dim myorderdetail,i

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
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " "
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


        sqlStr = " update " & TABLE_ORDERMASTER & " "     + VbCrlf
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
    'response.write "<script>opener.top.listFrame.location.reload();</script>"
    'response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="modifyreceiverinfo") then
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


        sqlStr = " select top 1 IsNULL(reqname,'') as reqname"
        sqlStr = sqlStr + " ,IsNULL(reqphone,'') as reqphone"
        sqlStr = sqlStr + " ,IsNULL(reqhp,'') as reqhp"
        sqlStr = sqlStr + " ,IsNULL(reqzipcode,'') as reqzipcode"
        sqlStr = sqlStr + " ,IsNULL(reqzipaddr,'') as reqzipaddr"
        sqlStr = sqlStr + " ,IsNULL(reqaddress,'') as reqaddress"
        sqlStr = sqlStr + " ,IsNULL(comment,'') as comment"
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " "
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
                contents_jupsu = contents_jupsu & "�������ּ�: [" & rsget("reqzipcode") & "] " & rsget("reqzipaddr") & " " & rsget("reqaddress") & CNEXT & "[" & reqzipcode & "] " & reqzipaddr & " " & reqaddress & VbCrlf
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


        sqlStr = " update " & TABLE_ORDERMASTER & " "     + VbCrlf
        sqlStr = sqlStr + " set reqname='" + html2db(reqname) + "' "   + VbCrlf
        sqlStr = sqlStr + " ,reqphone = '" + CStr(reqphone) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,reqhp = '" + CStr(reqhp) + "' "        + VbCrlf
        sqlStr = sqlStr + " ,reqzipcode = '" + CStr(reqzipcode) + "' "  + VbCrlf
        sqlStr = sqlStr + " ,reqzipaddr = '" + CStr(reqzipaddr) + "' "    + VbCrlf
        sqlStr = sqlStr + " ,reqaddress = '" + html2db(reqaddress) + "' "    + VbCrlf
        sqlStr = sqlStr + " ,comment = '" + html2db(comment) + "' "    + VbCrlf
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
    'response.write "<script>opener.top.listFrame.location.reload();</script>"
    'response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.location.reload();</script>"
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
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & " "
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


        sqlStr = " update " & TABLE_ORDERMASTER & " "     + VbCrlf
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
    'response.write "<script>opener.top.listFrame.location.reload();</script>"
    'response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>opener.focus(); window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="ipkumdivnextstep") then
    if (ipkumdiv="2") then
        divcd   = "A900"
        title   = "�����Ϸ� ��������"
        gubun01 = "C004"
        gubun02 = "CD99"

        ''�޸�� �Է��ϰ� ����

        sqlStr = "select top 1 userid, buyname, buyhp, buyemail, cancelyn "
        sqlStr = sqlStr + " from " & TABLE_ORDERMASTER & ""
        sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            userid      = rsget("userid")
        	buyname     = db2html(rsget("buyname"))
        	buyhp       = db2html(rsget("buyhp"))
        	buyemail    = db2html(rsget("buyemail"))
        	cancelyn    = rsget("cancelyn")
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

            sqlStr =	"update " & TABLE_ORDERMASTER & " " & vbCrlf
    		sqlStr = sqlStr & " set ipkumdiv='4'" & vbCrlf
    		sqlStr = sqlStr & " ,ipkumdate=getdate()" & vbCrlf
    		sqlStr = sqlStr & " where orderserial='" & orderserial & "'"

    		dbget.Execute sqlStr

    		''��� ������Ʈ
    		'sqlStr = " exec db_summary.dbo.sp_ten_RealtimeStock_regIpkum '" & orderserial & "'"
    		'dbget.Execute sqlStr
	    end IF

	    If Err.Number = 0 Then
            errcode = "002"
		    '' �ֹ� ���ϸ��� ������Ʈ
		    CALL updateUserMileage(userid)
		end IF

		If (Err.Number = 0) and (smsok<>"") Then
            errcode = "003"

		    '' SMS �߼�
            set osms = new CSMSClass
                osms.SendAcctIpkumOkMsg buyhp,orderserial
            set osms = Nothing

	    end IF

		If (Err.Number = 0) and (emailok<>"") Then
            errcode = "004"

		    '' Email �߼�
		        Call SendMailBankOk(buyemail,buyname,orderserial)

		end IF


		If Err.Number = 0 Then
            errcode = "005"
            call AddCsMemo(orderserial,"1",userid,reguserid,"�����Ϸ� ��������")
        end if


		If Err.Number = 0 Then
            dbget.CommitTrans
			errcode = "006"
			''2017/01/17 �߰� corpse2 ��ü��� �ֹ� ������ ���� ��� push �߼�
			sqlStr = "exec db_academy.[dbo].[sp_ACA_sendPushMsgOrderSuccess_Artist] '" & Cstr(orderserial) & "'"
			dbget.execute(sqlStr)
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.\r\n(�����ڵ� : " + CStr(errcode) + "|" + Replace(Err.description,Vbcrlf," ") + ")" + Chr(34) + ")</script>"
            response.write "<script>history.back()</script>"
            dbget.close()	:	response.End
        End If


        response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
    'response.write "<script>opener.top.listFrame.location.reload();</script>"
    'response.write "<script>opener.top.detailFrame.location.reload();</script>"
    response.write "<script>opener.location.reload();</script>"
        response.write "<script>opener.focus(); window.close();</script>"
        dbget.close()	:	response.End

    else
        response.write "<script>alert('�Ա� ��� ���¿����� ���� ���·� ���� �����մϴ�.');</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if
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
        sqlStr = sqlStr + " from " & TABLE_ORDERDETAIL & " "
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

        sqlStr = "update " & TABLE_ORDERDETAIL & "" + VbCrlf
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
end if



%>
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->