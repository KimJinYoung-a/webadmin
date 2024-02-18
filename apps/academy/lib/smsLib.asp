<%
CONST CNORMALCALLBAKC = "1644-1557"

function SendNormalSMS(reqhp,callback,smstext)
    dim sqlStr, RetRows
    if callback="" then callback=CNORMALCALLBAKC

    sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
	sqlStr = sqlStr + " values('" + reqhp + "',"
	sqlStr = sqlStr + " '" + callback + "',"
	sqlStr = sqlStr + " '1',"
	sqlStr = sqlStr + " getdate(),"
	sqlStr = sqlStr + " '" + html2db(smstext) + "')"

	dbget.Execute sqlStr, RetRows

	SendNormalSMS = (RetRows=1)
end function


function SendOverLengthSMS(reqhp,callback,smstext)
    dim smstext1, smstext2, smstext3
    dim retVal : retVal=false
    if callback="" then callback=CNORMALCALLBAKC

    if LenB(smstext)>160 then
        smstext1 = LeftB(smstext,80)
        smstext2 = MidB(smstext,81,80)
        smstext3 = MidB(smstext,161,80)
    elseif LenB(smstext)>80 then
        smstext1 = LeftB(smstext,80)
        smstext2 = MidB(smstext,81,80)
    else
        smstext1 = smstext
    end if

    if (Trim(smstext1)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext1)
    end if

    if (retVal) and (Trim(smstext2)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext2)
    end if

    if (retVal) and (Trim(smstext3)<>"") then
        retVal = SendNormalSMS(reqhp,callback,smstext3)
    end if

    SendOverLengthSMS = retVal
end function

function SendMultiRowsSMS(reqhp,callback,smstext,spliter)
    dim MaxRows : MaxRows=10
    dim smstextArr, i : i=0
    dim retVal : retVal=false
    if (callback="") then callback=CNORMALCALLBAKC
    if (spliter="") then spliter=VbCrlf

    smstextArr = split(smstext,spliter)

    if IsArray(smstextArr) then
        for i=LBound(smstextArr) to UBound(smstextArr)
            if (i>MaxRows) then Exit for
            if (Trim(smstextArr(i))<>"") then
                retVal = SendNormalSMS(reqhp,callback,smstextArr(i))
            end if
        next
    else
        retVal =SendNormalSMS(reqhp,callback,smstext)
    end if
    SendMultiRowsSMS = retVal
end function

function SendMiChulgoSMS(detailidx)
    dim oneMisend, smstext, buyhp
    set oneMisend = new COldMiSend
        oneMisend.FRectDetailIDx = detailidx
        oneMisend.getOneOldMisendItem

        smstext = oneMisend.FOneItem.getSMSText
        buyhp = oneMisend.FOneItem.FBuyHP


    if (smstext<>"") and (buyhp<>"") then
        SendMiChulgoSMS = SendMultiRowsSMS(buyhp,"",smstext,vbCrlf)

        'call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[SMS "+ buyhp +"]" + html2db(smstext))
    end if
    set oneMisend = Nothing
end function

%>