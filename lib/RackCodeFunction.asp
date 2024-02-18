<%
'###########################################################
' Description : 랙코드 공통함수
' Hieditor : 이상구 생성
'			 2021.03.24 한용민 수정
'###########################################################

''/lib/RackCodeFunction.asp
'' SCM 과 로직스 모두 동일한 내용이어야 한다.

'// ============================================================================
'// 상품 랙코드 입력
'// function RF_SetItemRackCode(itemgubun, itemid, rackcode)

'// 상품 보조랙코드 입력
'// function RF_SetSubItemRackCode(itemgubun, itemid, rackcode)

'// 상품 옵션별 랙코드 입력
'// function RF_SetItemRackCodeByOption(itemgubun, itemid, itemoption, rackcode)

function RF_SetBrandRackCode(makerid, rackcode)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, prerackcode, prerackcode4, prerackcode8

    sqlStr = "select userid, IsNULL(rackcodeByBrand, IsNULL(prtidx,'9999')) as prtidx from [db_user].[dbo].tbl_user_c" + VbCrlf
	sqlStr = sqlStr + " where userid='" + makerid + "'" + VbCrlf
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		prerackcode = rsget("prtidx")
	end if
	rsget.Close

    Call RF_GetRackCodeBy4By8(rackcode, rackcode4, rackcode8)
    Call RF_GetRackCodeBy4By8(prerackcode, prerackcode4, prerackcode8)

    sqlStr = "update [db_user].[dbo].tbl_user_c" & VbCrlf
    sqlStr = sqlStr & " set prtidx='" & rackcode4 & "', rackcodeByBrand='" & rackcode8 & "' " & VbCrlf
    sqlStr = sqlStr & " where userid='" & makerid & "'"

    if (makerid<>"") and (rackcode4<>"") then
        dbget.Execute sqlStr
    end if

    Call RF_CopyItemRackCodeByBrand(makerid)

    if (rackcode8 <> prerackcode8) then
        sqlStr = " update s "
	    sqlStr = sqlStr + " set s.rackcodeByOption = '" & rackcode8 & "' "
	    sqlStr = sqlStr + " from "
	    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
	    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] s "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		1 = 1 "
	    sqlStr = sqlStr + " 		and i.makerid = '" & makerid & "' "
	    sqlStr = sqlStr + " 		and s.itemgubun = '10' "
	    sqlStr = sqlStr + " 		and i.itemid = s.itemid "
	    sqlStr = sqlStr + " 		and s.itemoption >= '0000' "
	    sqlStr = sqlStr + " 		and s.rackcodeByOption = '" & prerackcode8 & "' "
        dbget.Execute sqlStr
    end if

    if (rackcode4 <> prerackcode4) then
        sqlStr = " update s "
	    sqlStr = sqlStr + " set s.optrackcode = '" & rackcode4 & "' "
	    sqlStr = sqlStr + " from "
	    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
	    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option] s "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		1 = 1 "
	    sqlStr = sqlStr + " 		and i.makerid = '" & makerid & "' "
	    sqlStr = sqlStr + " 		and i.itemid = s.itemid "
	    sqlStr = sqlStr + " 		and s.optrackcode = '" & prerackcode4 & "' "
        dbget.Execute sqlStr

        sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
	    sqlStr = sqlStr + " set itemrackcode='" + CStr(rackcode4) + "'" + VbCrlf
	    sqlStr = sqlStr + " where makerid='" + makerid + "'" + VbCrlf
	    sqlStr = sqlStr + " and itemrackcode='" + CStr(prerackcode4) + "'"
	    dbget.Execute sqlStr
    end if
end function

function RF_SetItemRackCode(itemgubun, itemid, rackcode)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, prerackcode, prerackcode4, prerackcode8

    Call RF_GetRackCodeBy4By8(rackcode, rackcode4, rackcode8)

    Call RF_CopyItemRackCodeByItemID(itemgubun, itemid)

    sqlStr = "select IsNull(rackcodeByOption, '99990000') as rackcodeByOption from [db_item].[dbo].[tbl_item_option_stock]" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + trim(itemgubun) + "' and itemid='" + trim(itemid) + "' and itemoption='0000' " + VbCrlf

    'response.write sqlStr & "<BR>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		prerackcode = rsget("rackcodeByOption")
	end if
	rsget.Close

    Call RF_GetRackCodeBy4By8(prerackcode, prerackcode4, prerackcode8)

    if prerackcode<>rackcode then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '10' as itemgubun, i.itemid, isnull(s.itemoption,'0000') as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate, s.rackcodeByOption, s.subRackcodeByOption, '랙코드변경' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='10'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '0000'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
	    dbget.Execute sqlStr
    end if

    sqlStr = " update s "
	sqlStr = sqlStr + " set s.rackcodeByOption = '" & rackcode8 & "' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item_option_stock] s "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.itemgubun = '" + itemgubun + "' "
	sqlStr = sqlStr + " 		and s.itemid = '" + itemid + "' "
	sqlStr = sqlStr + " 		and s.itemoption >= '0000' "
	sqlStr = sqlStr + " 		and (s.rackcodeByOption = '" & prerackcode8 & "' or s.rackcodeByOption is NULL) "

    'response.write sqlStr & "<BR>"
    dbget.Execute sqlStr

    if (itemgubun = "10") then
	    sqlStr = "update [db_item].[dbo].[tbl_item_option]" + VBCrlf
	    sqlStr = sqlStr + " set optrackcode='" + rackcode4 + "'" + VBCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and (optrackcode = '" & prerackcode4 & "') "
	    dbget.Execute sqlStr

	    sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	    sqlStr = sqlStr + " set lastupdate = getdate()" + VBCrlf
	    sqlStr = sqlStr + " , itemrackcode='" + rackcode4 + "'" + VBCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	    dbget.Execute sqlStr
    else
		sqlStr = "update db_shop.dbo.tbl_shop_item" + VbCrlf
		sqlStr = sqlStr + " set offitemrackcode='" + rackcode4 + "'" + VbCrlf
		sqlStr = sqlStr + " where shopitemid=" + CStr(itemid)  + VbCrlf
		sqlStr = sqlStr + " and itemgubun='" + CStr(itemgubun)  + "'" + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + CStr(itemoption)  + "'" + VbCrlf
        dbget.Execute sqlStr
    end if
end function

function RF_SetSubItemRackCode(itemgubun, itemid, rackcode)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, prerackcode, prerackcode4, prerackcode8

    if (rackcode = "") then
        rackcode = "9999"
    end if

    Call RF_GetRackCodeBy4By8(rackcode, rackcode4, rackcode8)

    Call RF_CopyItemRackCodeByItemID(itemgubun, itemid)

    sqlStr = "select IsNull(subRackcodeByOption, '99990000') as subRackcodeByOption from [db_item].[dbo].[tbl_item_option_stock]" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "' and itemid='" + itemid + "' and itemoption='0000' " + VbCrlf
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		prerackcode = rsget("subRackcodeByOption")
	end if
	rsget.Close

    Call RF_GetRackCodeBy4By8(prerackcode, prerackcode4, prerackcode8)

    if prerackcode<>rackcode then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '10' as itemgubun, i.itemid, isnull(s.itemoption,'0000') as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate, s.rackcodeByOption, s.subRackcodeByOption, '보조랙코드변경' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='10'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '0000'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
	    dbget.Execute sqlStr
    end if

    sqlStr = " update s "
	sqlStr = sqlStr + " set s.subRackcodeByOption = '" & rackcode8 & "' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item_option_stock] s "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and s.itemgubun = '" + itemgubun + "' "
	sqlStr = sqlStr + " 		and s.itemid = '" + itemid + "' "
	sqlStr = sqlStr + " 		and s.itemoption >= '0000' "
	sqlStr = sqlStr + " 		and (s.subRackcodeByOption = '" & prerackcode8 & "' or s.subRackcodeByOption is NULL) "
    dbget.Execute sqlStr

    if (itemgubun = "10") then
	    sqlStr = "update [db_item].[dbo].[tbl_item_logics_addinfo]" + VBCrlf
	    sqlStr = sqlStr + " set subitemrackcode='" + rackcode4 + "'" + VBCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	    dbget.Execute sqlStr
    end if
end function

function RF_SetItemRackCodeByOption(itemgubun, itemid, itemoption, rackcode)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, prerackcode, preitemoptionrackcode, presubrackcode, presubitemoptionrackcode

    Call RF_GetRackCodeBy4By8(rackcode, rackcode4, rackcode8)

    if itemgubun = "10" then
	    sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
	    sqlStr = sqlStr + " set optrackcode='" + rackcode4 + "'" + VBCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
	    dbget.Execute sqlStr

        if (itemoption = "0000") then
	        sqlStr = "update [db_item].[dbo].[tbl_item]" + VBCrlf
	        sqlStr = sqlStr + " set itemrackcode='" + rackcode4 + "'" + VBCrlf
	        sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	        dbget.Execute sqlStr
        end if
    else
		sqlStr = "update db_shop.dbo.tbl_shop_item" + VBCrlf
		sqlStr = sqlStr + " set offitemrackcode='" + rackcode4 + "'" + VBCrlf
		sqlStr = sqlStr + " , updt=getdate()" + VBCrlf
		sqlStr = sqlStr + " where itemgubun='" & CStr(itemgubun) & "' and shopitemid=" & CStr(itemid) & " and itemoption='" & CStr(itemoption) & "' and IsNull(offitemrackcode, '') <> '" + rackcode4 + "' "
	    dbget.Execute sqlStr
    end if

    sqlStr = "select IsNull(rackcodeByOption, '99990000') as rackcodeByOption, isnull(subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr + " from [db_item].[dbo].[tbl_item_option_stock] with (nolock)" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + trim(itemgubun) + "' and itemid='" + trim(itemid) + "' and itemoption='0000' " + VbCrlf

    'response.write sqlStr & "<BR>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		prerackcode = rsget("rackcodeByOption")
		presubrackcode = rsget("subRackcodeByOption")
	end if
	rsget.Close

    sqlStr = "select isnull(s.rackcodeByOption,'') as rackcodeByOption, isnull(s.subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
    sqlStr = sqlStr & " where itemgubun='" & trim(itemgubun) & "' and itemid='" & trim(itemid) & "' and itemoption='" & trim(itemoption) & "' "

    'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        preitemoptionrackcode = rsget("rackcodeByOption")
        presubitemoptionrackcode = rsget("subRackcodeByOption")
    end if
    rsget.Close

    if rackcode8<>"" and preitemoptionrackcode<>rackcode8 then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '" & trim(itemgubun) & "' as itemgubun, i.itemid, '" & trim(itemoption) & "' as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate"

        ' 로그작성 이전에 랙코드값이 없을경우 기본값으로 옵션 랙코드를 넣는 로직이 있어서 저장시 무조건 로그가 쌓이는 문제가 있어서
        ' 옵션없음 랙코드 앞자리4와 0000를 합한값이 옵션랙코드와 같은 경우에는 이전 데이터 없음으로 넣는다.
        if (left(prerackcode,4)+"0000"=preitemoptionrackcode)then
            sqlStr = sqlStr & " , '' as rackcodeByOption"
        else
            sqlStr = sqlStr & " , s.rackcodeByOption"
        end if
        if (left(presubrackcode,4)+"0000"=presubitemoptionrackcode)then
            sqlStr = sqlStr & " , '' as subRackcodeByOption"
        else
            sqlStr = sqlStr & " , s.subRackcodeByOption"
        end if

        sqlStr = sqlStr & " , '랙코드변경' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='" & trim(itemgubun) & "'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '" & trim(itemoption) & "'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
        dbget.Execute sqlStr
    end if

	sqlStr = "update [db_item].[dbo].[tbl_item_option_stock]" + VBCrlf
	sqlStr = sqlStr + " set  rackcodeByOption='" + rackcode8 + "'" + VBCrlf
	sqlStr = sqlStr + " where itemgubun = '" & itemgubun & "' and itemid=" + CStr(itemid) + " and itemoption = '" & itemoption & "' "

    'response.write sqlStr & "<br>"
	dbget.Execute sqlStr, affectedRows

    if (affectedRows = 0) then
        sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption) "
        sqlStr = sqlStr + " select '" & itemgubun & "' as itemgubun, " & itemid & " as itemid, '" & itemoption & "' as itemoption, '" & rackcode8 & "' as rackcodeByOption "

        'response.write sqlStr & "<br>"
        dbget.Execute sqlStr
    end if
end function

function RF_DelItemRackCodeByOption(itemgubun, itemid, itemoption)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, rackcodecount, prerackcode, preitemoptionrackcode

    sqlStr = "select IsNull(rackcodeByOption, '99990000') as rackcodeByOption from [db_item].[dbo].[tbl_item_option_stock] with (nolock)" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + trim(itemgubun) + "' and itemid='" + trim(itemid) + "' and itemoption='0000' " + VbCrlf

    'response.write sqlStr & "<BR>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		prerackcode = rsget("rackcodeByOption")
	end if
	rsget.Close

    sqlStr = "select isnull(s.rackcodeByOption,'') as rackcodeByOption, isnull(s.subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
    sqlStr = sqlStr & " where itemgubun='" & trim(itemgubun) & "' and itemid='" & trim(itemid) & "' and itemoption='" & trim(itemoption) & "' "

    'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        preitemoptionrackcode = rsget("rackcodeByOption")
    end if
    rsget.Close

    ' 로그작성 이전에 랙코드값이 없을경우 기본값으로 옵션 랙코드를 넣는 로직이 있어서 저장시 무조건 로그가 쌓이는 문제가 있어서
    ' 이전옵션랙코드가 있고, 옵션없음 랙코드 앞자리4와 0000를 합한값이 옵션랙코드와 다른값일 경우에만 로그를 쌓는다.
    if preitemoptionrackcode<>"" and (left(prerackcode,4)+"0000"<>preitemoptionrackcode)then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '" & trim(itemgubun) & "' as itemgubun, i.itemid, '" & trim(itemoption) & "' as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate, s.rackcodeByOption, s.subRackcodeByOption, '랙코드삭제' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='" & trim(itemgubun) & "'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '" & trim(itemoption) & "'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
	    dbget.Execute sqlStr
    end if

    if itemgubun = "10" then
	    sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
	    sqlStr = sqlStr + " set optrackcode=NULL" + VBCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
        sqlStr = sqlStr + " and itemoption='" & itemoption & "'"
	    dbget.Execute sqlStr
    end if

	sqlStr = "update [db_item].[dbo].[tbl_item_option_stock]" + VBCrlf
	sqlStr = sqlStr + " set  rackcodeByOption=NULL" + VBCrlf
	sqlStr = sqlStr + " where itemgubun = '" & itemgubun & "' and itemid=" + CStr(itemid) + " and itemoption = '" & itemoption & "' "
	dbget.Execute sqlStr, affectedRows
end function

function RF_SetSubItemRackCodeByOption(itemgubun, itemid, itemoption, rackcode)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, presubrackcode, presubitemoptionrackcode, prerackcode, preitemoptionrackcode

    Call RF_GetRackCodeBy4By8(rackcode, rackcode4, rackcode8)

    sqlStr = "select IsNull(rackcodeByOption, '99990000') as rackcodeByOption, isnull(subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr + " from [db_item].[dbo].[tbl_item_option_stock] with (nolock)" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + trim(itemgubun) + "' and itemid='" + trim(itemid) + "' and itemoption='0000' " + VbCrlf

    'response.write sqlStr & "<BR>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		prerackcode = rsget("rackcodeByOption")
		presubrackcode = rsget("subRackcodeByOption")
	end if
	rsget.Close

    sqlStr = "select isnull(s.rackcodeByOption,'') as rackcodeByOption, isnull(s.subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
    sqlStr = sqlStr & " where itemgubun='" & trim(itemgubun) & "' and itemid='" & trim(itemid) & "' and itemoption='" & trim(itemoption) & "' "

    'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        preitemoptionrackcode = rsget("rackcodeByOption")
        presubitemoptionrackcode = rsget("subRackcodeByOption")
    end if
    rsget.Close

    if rackcode8<>"" and presubitemoptionrackcode<>rackcode8 then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '" & trim(itemgubun) & "' as itemgubun, i.itemid, '" & trim(itemoption) & "' as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate"

        ' 로그작성 이전에 랙코드값이 없을경우 기본값으로 옵션 랙코드를 넣는 로직이 있어서 저장시 무조건 로그가 쌓이는 문제가 있어서
        ' 옵션없음 랙코드 앞자리4와 0000를 합한값이 옵션랙코드와 같은 경우에는 이전 데이터 없음으로 넣는다.
        if (left(prerackcode,4)+"0000"=preitemoptionrackcode)then
            sqlStr = sqlStr & " , '' as rackcodeByOption"
        else
            sqlStr = sqlStr & " , s.rackcodeByOption"
        end if
        if (left(presubrackcode,4)+"0000"=presubitemoptionrackcode)then
            sqlStr = sqlStr & " , '' as subRackcodeByOption"
        else
            sqlStr = sqlStr & " , s.subRackcodeByOption"
        end if

        sqlStr = sqlStr & " , '보조랙코드변경' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='" & trim(itemgubun) & "'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '" & trim(itemoption) & "'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
        dbget.Execute sqlStr
    end if

	sqlStr = "update [db_item].[dbo].[tbl_item_option_stock]" + VBCrlf
	sqlStr = sqlStr + " set  subrackcodeByOption='" + rackcode8 + "'" + VBCrlf
	sqlStr = sqlStr + " where itemgubun = '" & itemgubun & "' and itemid=" + CStr(itemid) + " and itemoption = '" & itemoption & "' "
	dbget.Execute sqlStr, affectedRows

    if (affectedRows = 0) then
        sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, subrackcodeByOption) "
        sqlStr = sqlStr + " select '" & itemgubun & "' as itemgubun, " & itemid & " as itemid, '" & itemoption & "' as itemoption, '" & rackcode8 & "' as subrackcodeByOption "
        dbget.Execute sqlStr
    end if
end function

function RF_DelSubItemRackCodeByOption(itemgubun, itemid, itemoption)
    dim sqlStr, affectedRows
    dim rackcode4, rackcode8, presubrackcode, presubitemoptionrackcode

    sqlStr = "select isnull(subRackcodeByOption,'') as subRackcodeByOption from [db_item].[dbo].[tbl_item_option_stock] with (nolock)" + VbCrlf
	sqlStr = sqlStr + " where itemgubun='" + trim(itemgubun) + "' and itemid='" + trim(itemid) + "' and itemoption='0000' " + VbCrlf

    'response.write sqlStr & "<BR>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		presubrackcode = rsget("subRackcodeByOption")
	end if
	rsget.Close

    sqlStr = "select isnull(s.rackcodeByOption,'') as rackcodeByOption, isnull(s.subRackcodeByOption,'') as subRackcodeByOption"
    sqlStr = sqlStr & " from [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
    sqlStr = sqlStr & " where itemgubun='" & trim(itemgubun) & "' and itemid='" & trim(itemid) & "' and itemoption='" & trim(itemoption) & "' "

    'response.write sqlStr & "<Br>"
    'response.end
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        presubitemoptionrackcode = rsget("subRackcodeByOption")
    end if
    rsget.Close

    ' 로그작성 이전에 랙코드값이 없을경우 기본값으로 옵션 랙코드를 넣는 로직이 있어서 저장시 무조건 로그가 쌓이는 문제가 있어서
    ' 이전옵션랙코드가 있고, 옵션없음 랙코드 앞자리4와 0000를 합한값이 옵션랙코드와 다른값일 경우에만 로그를 쌓는다.
    if presubitemoptionrackcode<>"" and (left(presubrackcode,4)+"0000"<>presubitemoptionrackcode)then
        ' 업데이트 이전 데이터 로그 저장    ' 2021.03.25 한용민 생성
        sqlStr = "insert into db_log.dbo.tbl_iteminfo_option_history (itemgubun, itemid, itemoption, regadminid, regdate, rackcodeByOption, subRackcodeByOption, comment)"
        sqlStr = sqlStr & " select"
        sqlStr = sqlStr & " '" & trim(itemgubun) & "' as itemgubun, i.itemid, '" & trim(itemoption) & "' as itemoption, '"& session("ssBctId") &"' as regadminid"
        sqlStr = sqlStr & " , getdate() as regdate, s.rackcodeByOption, s.subRackcodeByOption, '보조랙코드삭제' as comment"
        sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)"
        sqlStr = sqlStr & " left join [db_item].[dbo].[tbl_item_option_stock] s with (nolock)"
        sqlStr = sqlStr & "     on s.itemgubun='" & trim(itemgubun) & "'"
        sqlStr = sqlStr & "     and i.itemid = s.itemid"
        sqlStr = sqlStr & "     and s.itemoption = '" & trim(itemoption) & "'"
        sqlStr = sqlStr & " where i.itemid = '" & trim(itemid) & "' "

        'response.write sqlStr & "<BR>"
	    dbget.Execute sqlStr
    end if

	sqlStr = "update [db_item].[dbo].[tbl_item_option_stock]" + VBCrlf
	sqlStr = sqlStr + " set  subrackcodeByOption=NULL" + VBCrlf
	sqlStr = sqlStr + " where itemgubun = '" & itemgubun & "' and itemid=" + CStr(itemid) + " and itemoption = '" & itemoption & "' "
	dbget.Execute sqlStr, affectedRows
end function

function RF_CopyItemRackCodeByBrand(makerid)
    dim sqlStr, affectedRows

    '// -- 데이타 없는 경우 입력(1/2)
    sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption) "
    sqlStr = sqlStr + " select '10', i.itemid, IsNull(o.itemoption, '0000'), IsNull(o.optrackcode, i.itemrackcode) + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = a.itemid "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = o.itemid "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.itemgubun is NULL "
    dbget.Execute sqlStr

    '// -- 데이타 없는 경우 입력(2/2)
    sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption, subRackcodeByOption) "
    sqlStr = sqlStr + " select '10', i.itemid, '0000', i.itemrackcode + '0000', a.subItemRackcode + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = a.itemid "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = '0000' "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.itemgubun is NULL "
    dbget.Execute sqlStr

    '// -- 옵션별 랙코드 입력
    sqlStr = " update os "
    sqlStr = sqlStr + " set os.rackcodeByOption = IsNull(o.optrackcode, i.itemrackcode) + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = a.itemid "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = o.itemid "
    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.rackcodeByOption is NULL "
    dbget.Execute sqlStr

    '// -- 옵션별 랙코드 입력
    sqlStr = " update os "
    sqlStr = sqlStr + " set os.rackcodeByOption = i.itemrackcode + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = '0000' "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.rackcodeByOption is NULL "
    dbget.Execute sqlStr

    '// -- 옵션별 보조랙코드 입력
    sqlStr = " update os "
    sqlStr = sqlStr + " set os.subRackcodeByOption = a.subItemRackcode + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = a.itemid "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = o.itemid "
    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.subRackcodeByOption is NULL "
    sqlStr = sqlStr + " 	and a.subItemRackcode is not NULL "
    ''dbget.Execute sqlStr

    '// -- 옵션별 보조랙코드 입력
    sqlStr = " update os "
    sqlStr = sqlStr + " set os.subRackcodeByOption = a.subItemRackcode + '0000' "
    sqlStr = sqlStr + " from "
    sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		i.itemid = a.itemid "
    sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
    sqlStr = sqlStr + " 	on "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and os.itemgubun = '10' "
    sqlStr = sqlStr + " 		and os.itemid = i.itemid "
    sqlStr = sqlStr + " 		and os.itemoption = '0000' "
    sqlStr = sqlStr + " where "
    sqlStr = sqlStr + " 	1 = 1 "
    sqlStr = sqlStr + " 	and i.makerid = '" & makerid & "' "
    sqlStr = sqlStr + " 	and os.subRackcodeByOption is NULL "
    sqlStr = sqlStr + " 	and a.subItemRackcode is not NULL "
    dbget.Execute sqlStr
end function

function RF_CopyItemRackCodeByItemID(itemgubun, itemid)
    dim sqlStr, affectedRows

    if (itemgubun = "10") then
        '// -- 데이타 없는 경우 입력(1/2)
        sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption) "
        sqlStr = sqlStr + " select '10', i.itemid, IsNull(o.itemoption, '0000'), IsNull(o.optrackcode, i.itemrackcode) + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = a.itemid "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = o.itemid "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.itemgubun is NULL "
        dbget.Execute sqlStr

        '// -- 데이타 없는 경우 입력(2/2)
        sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption, subRackcodeByOption) "
        sqlStr = sqlStr + " select '10', i.itemid, '0000', i.itemrackcode + '0000', a.subItemRackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = a.itemid "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = '0000' "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.itemgubun is NULL "
        dbget.Execute sqlStr

        '// -- 옵션별 랙코드 입력
        sqlStr = " update os "
        sqlStr = sqlStr + " set os.rackcodeByOption = IsNull(o.optrackcode, i.itemrackcode) + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = a.itemid "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = o.itemid "
        sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.rackcodeByOption is NULL "
        dbget.Execute sqlStr

        '// -- 옵션별 랙코드 입력
        sqlStr = " update os "
        sqlStr = sqlStr + " set os.rackcodeByOption = i.itemrackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = '0000' "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.rackcodeByOption is NULL "
        dbget.Execute sqlStr

        '// -- 옵션별 보조랙코드 입력
        sqlStr = " update os "
        sqlStr = sqlStr + " set os.subRackcodeByOption = a.subItemRackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = a.itemid "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option] o "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = o.itemid "
        sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = IsNull(o.itemoption, '0000') "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.subRackcodeByOption is NULL "
        sqlStr = sqlStr + " 	and a.subItemRackcode is not NULL "
        ''dbget.Execute sqlStr

        '// -- 옵션별 보조랙코드 입력
        sqlStr = " update os "
        sqlStr = sqlStr + " set os.subRackcodeByOption = a.subItemRackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	[db_item].[dbo].[tbl_item] i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] a "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		i.itemid = a.itemid "
        sqlStr = sqlStr + " 	join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun = '10' "
        sqlStr = sqlStr + " 		and os.itemid = i.itemid "
        sqlStr = sqlStr + " 		and os.itemoption = '0000' "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.subRackcodeByOption is NULL "
        sqlStr = sqlStr + " 	and a.subItemRackcode is not NULL "
        dbget.Execute sqlStr
    else
        sqlStr = " insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun, itemid, itemoption, rackcodeByOption) "
        sqlStr = sqlStr + " select i.itemgubun, i.shopitemid, i.itemoption, i.offitemrackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun <> '10' "
        sqlStr = sqlStr + " 		and os.itemgubun = i.itemgubun "
        sqlStr = sqlStr + " 		and os.itemid = i.shopitemid "
        sqlStr = sqlStr + " 		and os.itemoption = i.itemoption "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemgubun = '" & itemgubun & "' "
        sqlStr = sqlStr + " 	and i.shopitemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and i.itemoption >= '0000' "
        sqlStr = sqlStr + " 	and os.itemgubun is NULL "
        dbget.Execute sqlStr

        sqlStr = " update os "
        sqlStr = sqlStr + " set os.rackcodeByOption = i.offitemrackcode + '0000' "
        sqlStr = sqlStr + " from "
        sqlStr = sqlStr + " 	db_shop.dbo.tbl_shop_item i "
        sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_option_stock] os "
        sqlStr = sqlStr + " 	on "
        sqlStr = sqlStr + " 		1 = 1 "
        sqlStr = sqlStr + " 		and os.itemgubun <> '10' "
        sqlStr = sqlStr + " 		and os.itemgubun = i.itemgubun "
        sqlStr = sqlStr + " 		and os.itemid = i.shopitemid "
        sqlStr = sqlStr + " 		and os.itemoption = i.itemoption "
        sqlStr = sqlStr + " where "
        sqlStr = sqlStr + " 	1 = 1 "
        sqlStr = sqlStr + " 	and i.itemgubun = '" & itemgubun & "' "
        sqlStr = sqlStr + " 	and i.shopitemid = '" & itemid & "' "
        sqlStr = sqlStr + " 	and os.rackcodeByOption is NULL "
        sqlStr = sqlStr + " 	and i.offitemrackcode is not NULL "
        dbget.Execute sqlStr
   end if
end function

function RF_GetRackCodeBy4By8(rackcode, ByRef rackcode4, ByRef rackcode8)
    if Len(rackcode) <> 4 and Len(rackcode) <> 8 then
        response.write "랙코드는 4자리 또는 8자리만 입력가능합니다."
        response.end
    elseif Len(rackcode) = 4 then
        rackcode4 = rackcode
        rackcode8 = rackcode & "0000"
    elseif Len(rackcode) = 8 then
        rackcode4 = Left(rackcode, 4)
        rackcode8 = rackcode
    end if
end function


%>
