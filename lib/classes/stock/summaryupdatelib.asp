<%
'입고 상품 업데이트 - 입고일, 상품구분, 상품코드, 상품옵션, 갯수, 반품구분
sub QuickUpdateItemIpgoSummary(byval yyyymmdd, byval itemgubun, byval itemid, byval itemoption, byval itemno, byval isreturn)
    dim found, sqlStr

	if IsNULL(yyyymmdd) then exit sub
	if (yyyymmdd="") then exit sub

    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_logisstock_summary "
    sqlStr = sqlStr + " where itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close


    if (found) then
        if (isreturn) then
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set reipgono = reipgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totipgono = totipgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno)
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)

        else
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set ipgono = ipgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totipgono = totipgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + " "
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)
        end if

        'response.write sqlStr
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,reipgono,totipgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,ipgono,totipgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if

        'response.write sqlStr
    end if


    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_daily_logisstock_summary "
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "' and itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close

    if (found = true) then
        if (isreturn) then
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
            sqlStr = sqlStr + " set reipgono = reipgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totipgono = totipgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        else
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
            sqlStr = sqlStr + " set ipgono = ipgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totipgono = totipgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        end if
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,reipgono,totipgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,ipgono,totipgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if
    end if

end sub



'샾 출고 상품 업데이트 - 입고일, 상품구분, 상품코드, 상품옵션, 갯수, 반품구분
sub QuickUpdateItemChulgoSummary(byval yyyymmdd, byval itemgubun, byval itemid, byval itemoption, byval itemno, byval isreturn)
    dim found, sqlStr

	if IsNULL(yyyymmdd) then exit sub
	if (yyyymmdd="") then exit sub

    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_logisstock_summary "
    sqlStr = sqlStr + " where itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close

    if (found = true) then
        if (isreturn) then
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set offrechulgono = offrechulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno)
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)

        else
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set offchulgono = offchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + " "
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)
        end if
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,offrechulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,offchulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if
    end if


    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_daily_logisstock_summary "
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "' and itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close

    if (found = true) then
        if (isreturn) then
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
            sqlStr = sqlStr + " set offrechulgono = offrechulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        else
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
           sqlStr = sqlStr + " set offchulgono = offchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        end if
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,offrechulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,offchulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if
    end if

end sub

'기타 출고 상품 업데이트 - 입고일, 상품구분, 상품코드, 상품옵션, 갯수, 반품구분
sub QuickUpdateItemEtcChulgoSummary(byval yyyymmdd, byval itemgubun, byval itemid, byval itemoption, byval itemno, byval isreturn)
    dim found, sqlStr

	if IsNULL(yyyymmdd) then exit sub
	if (yyyymmdd="") then exit sub

    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_current_logisstock_summary "
    sqlStr = sqlStr + " where itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close

    if (found = true) then
        if (isreturn) then
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set etcrechulgono = etcrechulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno)
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)

        else
            sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary "
            sqlStr = sqlStr + " set etcchulgono = etcchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno)
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno)
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + " "
            sqlStr = sqlStr + " , shortageno = shortageno + " + CStr(itemno)
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where itemgubun = '" + CStr(itemgubun) + "'"
            sqlStr = sqlStr + " and itemid = " + CStr(itemid)
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' "

            dbget.Execute(sqlStr)
        end if
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,etcrechulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary "
        	sqlStr = sqlStr + " (itemgubun,itemid,itemoption,etcchulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if
    end if


    found = false
    sqlStr = " select top 1 itemid "
    sqlStr = sqlStr + " from [db_summary].[dbo].tbl_daily_logisstock_summary "
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "' and itemgubun='" + itemgubun + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' "
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        found = true
    end if
    rsget.close

    if (found = true) then
        if (isreturn) then
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
            sqlStr = sqlStr + " set etcrechulgono = etcrechulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        else
        	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary " + VbCrlf
           sqlStr = sqlStr + " set etcchulgono = etcchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totchulgono = totchulgono + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , totsysstock = totsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , availsysstock = availsysstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " , realstock = realstock + " + CStr(itemno) + VbCrlf
            sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
            sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
            sqlStr = sqlStr + " and itemgubun = '" + CStr(itemgubun) + "'" + VbCrlf
            sqlStr = sqlStr + " and itemid = " + CStr(itemid) + VbCrlf
            sqlStr = sqlStr + " and itemoption = '" + CStr(itemoption) + "' " + VbCrlf

            dbget.Execute(sqlStr)
        end if
    else
        if (isreturn) then
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,etcrechulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        else
        	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary "
        	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,etcchulgono,totchulgono,totsysstock,availsysstock,realstock)"
        	sqlStr = sqlStr + " values("
        	sqlStr = sqlStr + " '" + yyyymmdd + "'"
        	sqlStr = sqlStr + " ,'" + itemgubun + "'"
        	sqlStr = sqlStr + " ," + CStr(itemid)
        	sqlStr = sqlStr + " ,'" + itemoption + "'"
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " ," + CStr(itemno)
        	sqlStr = sqlStr + " )"

        	dbget.Execute(sqlStr)
        end if
    end if

end sub


'입고코드를 이용해 신규입고데이타 반영 - 입출코드, 삭제
sub QuickUpdateNewIpgoDetailSummary(byval stcode, byval isdelete)
    dim found, sqlStr
    dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr
    dim itemgubun, itemid, itemoption, itemno
	dim STOCKBASEDATE, yyyymmdd, ipchulflag
	dim squareFactor

	if (isdelete) then
		squareFactor = -1
	else
		squareFactor = 1
	end if

	''최근 2달 내역만 업데이트 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  convert(varchar(7),dateadd(m,-1,getdate()),21)+'-01' as STOCKBASEDATE, m.executedt, m.ipchulflag"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.code='" + stcode + "'"

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		STOCKBASEDATE = rsget("STOCKBASEDATE")
		yyyymmdd = rsget("executedt")
		ipchulflag = rsget("ipchulflag")
	end if
	rsget.close

	if IsNULL(yyyymmdd) or (yyyymmdd="") then Exit sub
	if (CDate(STOCKBASEDATE)>CDate(yyyymmdd)) then Exit sub

	yyyymmdd = Left(CStr(yyyymmdd),10)

    sqlStr = " select d.iitemgubun as itemgubun, d.itemid, d.itemoption, d.itemno "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d "
    sqlStr = sqlStr + " where d.mastercode= '" + CStr(stcode) + "' "
    sqlStr = sqlStr + " and d.deldt is null "
    rsget.Open sqlStr,dbget,1

    if Not rsget.Eof then
		do until rsget.eof
			itemgubunarr = itemgubunarr + CStr(rsget("itemgubun")) + "|"
			itemidarr = itemidarr + CStr(rsget("itemid")) + "|"
			itemoptionarr = itemoptionarr + CStr(rsget("itemoption")) + "|"
			itemnoarr = itemnoarr + CStr(rsget("itemno")) + "|"

			rsget.movenext
			i=i+1
		loop
    end if
    rsget.close

    itemgubunarr = split(itemgubunarr, "|")
    itemidarr = split(itemidarr, "|")
    itemoptionarr = split(itemoptionarr, "|")
    itemnoarr = split(itemnoarr, "|")

	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = trim(itemnoarr(i))

			if ipchulflag="I" then
				QuickUpdateItemIpgoSummary yyyymmdd, itemgubun, itemid, itemoption, itemno * squareFactor ,(itemno<0)
			elseif ipchulflag="S" then
				QuickUpdateItemChulgoSummary yyyymmdd, itemgubun, itemid, itemoption, itemno * squareFactor ,(itemno>0)
			elseif ipchulflag="E" then
				QuickUpdateItemEtcChulgoSummary yyyymmdd, itemgubun, itemid, itemoption, itemno * squareFactor ,(itemno>0)
			end if
		end if
	next
end sub







''기주문 수량 업데이트 :
function PreOrderUpdateByBrand(targetid)
	dim sqlStr

	if targetid="10x10" then exit function
    
    sqlStr = " exec db_summary.dbo.ten_preOrder_Update '" & targetid & "','',0,''"
    dbget.Execute sqlStr
    Exit function 
    
    
	''초기화.
	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=0"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"
	sqlStr = sqlStr + " where i.makerid='" + targetid + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun='10'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=i.itemid"

	rsget.Open sqlStr,dbget,1


	''상품이 존재 하지 않을경우 1..
	sqlStr = " insert into [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " (itemgubun, itemid, itemoption, preorderno, preordernofix)"
	sqlStr = sqlStr + " select T.itemgubun, T.itemid, T.itemoption, T.preorderno, T.preordernofix "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " ("
	sqlStr = sqlStr + " 	select d.itemgubun, d.itemid, d.itemoption, sum(baljuitemno) as preorderno, sum(realitemno) as preordernofix  "
	sqlStr = sqlStr + " 	from [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " 	where m.idx=d.masteridx"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and m.ipgodate is null"
	sqlStr = sqlStr + " 	and datediff(d,m.scheduledate,getdate())<10"
	sqlStr = sqlStr + " 	and m.baljuid='10x10'"
	sqlStr = sqlStr + " 	and m.targetid='" + targetid + "'"
	sqlStr = sqlStr + " 	and m.statecd<9"
	sqlStr = sqlStr + " 	and m.divcode in ('300','301','302')"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by d.itemgubun, d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_current_logisstock_summary s "
	sqlStr = sqlStr + " on T.itemgubun=s.itemgubun"
	sqlStr = sqlStr + " and T.itemid=s.itemid"
	sqlStr = sqlStr + " and T.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemgubun is null"

	rsget.Open sqlStr,dbget,1

	''상품이 존재 할 경우 2..
	sqlStr = " update [db_summary].[dbo].tbl_current_logisstock_summary"
	sqlStr = sqlStr + " set preorderno=IsNULL(T.preorderno,0)"
	sqlStr = sqlStr + " , preordernofix=IsNULL(T.preordernofix,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select d.itemgubun, d.itemid, d.itemoption, sum(baljuitemno) as preorderno, sum(realitemno) as preordernofix  "
	sqlStr = sqlStr + " 	from [db_storage].[dbo].tbl_ordersheet_master m,"
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_detail d"
	sqlStr = sqlStr + " 	where m.idx=d.masteridx"
	sqlStr = sqlStr + " 	and m.deldt is null"
	sqlStr = sqlStr + " 	and m.ipgodate is null"
	sqlStr = sqlStr + " 	and datediff(d,m.scheduledate,getdate())<10"
	sqlStr = sqlStr + " 	and m.baljuid='10x10'"
	sqlStr = sqlStr + " 	and m.targetid='" + targetid + "'"
	sqlStr = sqlStr + " 	and m.statecd<9"
	sqlStr = sqlStr + " 	and m.divcode in ('300','301','302')"
	sqlStr = sqlStr + " 	and d.deldt is null"
	sqlStr = sqlStr + " 	group by d.itemgubun, d.itemid, d.itemoption"
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"

	rsget.Open sqlStr,dbget,1
end function


sub OffStockUpdateUpcheIpgoByIdx(byval iidx)
	dim sqlStr
	dim makerid, shopid
	sqlStr = "select top 1 chargeid, shopid from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(iidx)

	rsget.Open sqlStr,dbget,1
	if Not rsget.Eof then
		makerid = rsget("chargeid")
		shopid = rsget("shopid")
	end if
	rsget.close

	if ((makerid="10x10") or (makerid="")) then
		exit sub
	end if

	OffStockUpdateUpcheIpgo makerid, shopid
end sub

sub OffStockUpdateUpcheIpgo(byval makerid,byval shopid )
	dim sqlStr
	sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlf
	sqlStr = sqlStr + " set upcheipno=T.upcheipno" + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " (" + VbCrlf
	sqlStr = sqlStr + " select m.shopid,d.itemgubun,d.shopitemid,d.itemoption,sum(d.itemno) as upcheipno" + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_master m," + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_day_stock s," + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail d" + VbCrlf
	sqlStr = sqlStr + " where m.chargeid<>'10x10'" + VbCrlf
	sqlStr = sqlStr + " and m.shopid=s.shopid" + VbCrlf
	sqlStr = sqlStr + " and m.execdt>=s.lastrealdate" + VbCrlf
	sqlStr = sqlStr + " and m.deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and m.idx=d.masteridx" + VbCrlf
	sqlStr = sqlStr + " and s.makerid='" + makerid + "'" + VbCrlf
	sqlStr = sqlStr + " and s.shopid='" + shopid + "'" + VbCrlf
	sqlStr = sqlStr + " and d.itemgubun=s.itemgubun" + VbCrlf
	sqlStr = sqlStr + " and d.shopitemid=s.itemid" + VbCrlf
	sqlStr = sqlStr + " and d.itemoption=s.itemoption" + VbCrlf
	sqlStr = sqlStr + " and d.deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and d.itemno>0" + VbCrlf
	sqlStr = sqlStr + " group  by m.shopid,d.itemgubun,d.shopitemid,d.itemoption" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_day_stock.shopid=T.shopid" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemgubun=T.itemgubun" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemid=T.shopitemid" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemoption=T.itemoption" + VbCrlf

	dbget.Execute(sqlStr)



	sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlf
	sqlStr = sqlStr + " set upchereno=T.upchereno" + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " (" + VbCrlf
	sqlStr = sqlStr + " select m.shopid,d.itemgubun,d.shopitemid,d.itemoption,sum(d.itemno*-1) as upchereno" + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_master m," + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_day_stock s," + VbCrlf
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail d" + VbCrlf
	sqlStr = sqlStr + " where m.chargeid<>'10x10'" + VbCrlf
	sqlStr = sqlStr + " and m.shopid=s.shopid" + VbCrlf
	sqlStr = sqlStr + " and m.execdt>=s.lastrealdate" + VbCrlf
	sqlStr = sqlStr + " and m.deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and m.idx=d.masteridx" + VbCrlf
	sqlStr = sqlStr + " and s.makerid='" + makerid + "'" + VbCrlf
	sqlStr = sqlStr + " and s.shopid='" + shopid + "'" + VbCrlf
	sqlStr = sqlStr + " and d.itemgubun=s.itemgubun" + VbCrlf
	sqlStr = sqlStr + " and d.shopitemid=s.itemid" + VbCrlf
	sqlStr = sqlStr + " and d.itemoption=s.itemoption" + VbCrlf
	sqlStr = sqlStr + " and d.deleteyn='N'" + VbCrlf
	sqlStr = sqlStr + " and d.itemno<0" + VbCrlf
	sqlStr = sqlStr + " group  by m.shopid,d.itemgubun,d.shopitemid,d.itemoption" + VbCrlf
	sqlStr = sqlStr + " ) T" + VbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_day_stock.shopid=T.shopid" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemgubun=T.itemgubun" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemid=T.shopitemid" + VbCrlf
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_day_stock.itemoption=T.itemoption" + VbCrlf

	dbget.Execute(sqlStr)



	sqlStr = " update [db_shop].[dbo].tbl_shop_day_stock" + VbCrlf
	sqlStr = sqlStr + " set currno=lastrealno+ipno-reno+upcheipno-upchereno-sellno" + VbCrlf
	sqlStr = sqlStr + " where makerid='" + makerid + "'" + VbCrlf
	sqlStr = sqlStr + " and shopid='" + shopid + "'" + VbCrlf

	dbget.Execute(sqlStr)
end sub

%>