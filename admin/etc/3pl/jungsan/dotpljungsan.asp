<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, tplcompanyid, yyyy1, mm1, differencekey
dim yyyymm, title, st_totalcash, io_totalcash, et_totalcash, finishflag, groupid, itemvatYn, taxtype, company_name, jungsan_hp, jungsan_email
dim masteridx, idx
dim i, j, k, tmpSql, tmpInsSql, gubun
dim itemno, cbmX, cbmY, cbmZ, itemname, unitprice, itypename

mode 			= requestCheckVar(request("mode"),32)
tplcompanyid 	= requestCheckVar(request("tplcompanyid"),32)
yyyy1 			= requestCheckVar(request("yyyy1"),32)
mm1 			= requestCheckVar(request("mm1"),32)
differencekey 	= requestCheckVar(request("differencekey"),32)
masteridx 		= requestCheckVar(request("masteridx"),32)
idx 			= requestCheckVar(request("idx"),32)
gubun 			= requestCheckVar(request("gubun"),32)

itemno 			= requestCheckVar(request("itemno"),32)
cbmX 			= requestCheckVar(request("cbmX"),32)
cbmY 			= requestCheckVar(request("cbmY"),32)
cbmZ 			= requestCheckVar(request("cbmZ"),32)
itemname 		= requestCheckVar(request("itemname"),64)
unitprice 		= requestCheckVar(request("unitprice"),64)
itypename 		= requestCheckVar(request("itypename"),64)

yyyymm = yyyy1 + "-" + mm1

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr

function fnExitIfNotEditState(masteridx, refer)
    dim sqlStr, notstatemodi

	sqlStr = "select top 1 finishflag from [db_threepl].[dbo].[tbl_tpl_jungsan_master] "
	sqlStr = sqlStr + " where idx=" + masteridx + ""
    ''response.write sqlStr
    ''response.end
    dbget_TPL.CursorLocation = adUseClient
    rsget_TPL.Open sqlStr,dbget_TPL,adOpenForwardOnly,adLockReadOnly

    notstatemodi = False

    If not rsget_TPL.EOF Then
        notstatemodi = Not (rsget_TPL("finishflag")="0")
    End If
    rsget_TPL.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('현재 수정중 상태가 아닙니다.');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget_TPL.Close : dbget.close()	:	response.End
	end if
end function

function fnUpdateCbmSummary(masteridx)
    dim sqlStr, yyyymm, tplcompanyid

    sqlStr = "select top 1 yyyymm, tplcompanyid from [db_threepl].[dbo].[tbl_tpl_jungsan_master] "
	sqlStr = sqlStr + " where idx=" & masteridx
	rsget_TPL.Open sqlStr,dbget_TPL,1
		yyyymm = rsget_TPL("yyyymm")
        tplcompanyid = rsget_TPL("tplcompanyid")
	rsget_TPL.Close

    sqlStr = " update "
    sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
    sqlStr = sqlStr + " set currcbm = ( "
    sqlStr = sqlStr + " 	select round(IsNull(sum(1.0 * c.cbmX * c.cbmY * c.cbmZ * c.itemno),0) / 1000000000,2) "
    sqlStr = sqlStr + " 	from "
    sqlStr = sqlStr + " 		[db_threepl].[dbo].[tbl_tpl_jungsan_master] m "
    sqlStr = sqlStr + " 		join [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] c on m.idx = c.masteridx "
    sqlStr = sqlStr + " 	where "
    sqlStr = sqlStr + " 		1 = 1 "
    sqlStr = sqlStr + " 		and m.tplcompanyid = '" & tplcompanyid & "' "
    sqlStr = sqlStr + " 		and m.yyyymm = '" & yyyymm & "' "
    sqlStr = sqlStr + " 		and m.differencekey = 0 "
    sqlStr = sqlStr + " ) "
    sqlStr = sqlStr + " where masteridx = " & masteridx & " and gubundetailname = '상품보관' "
    dbget_TPL.Execute sqlStr

    sqlStr = " update "
	sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	sqlStr = sqlStr + " set prevcbm = ( "
	sqlStr = sqlStr + "		select round(IsNull(sum(1.0 * c.cbmX * c.cbmY * c.cbmZ * c.itemno),0) / 1000000000,2) "
	sqlStr = sqlStr + "		from "
	sqlStr = sqlStr + "			[db_threepl].[dbo].[tbl_tpl_jungsan_master] m "
	sqlStr = sqlStr + "			join [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] c on m.idx = c.masteridx "
	sqlStr = sqlStr + "		where "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and m.tplcompanyid = '" & tplcompanyid & "' "
	sqlStr = sqlStr + "			and m.yyyymm = '" & Left(DateAdd("m", -1, yyyymm + "-01"),7) & "' "
	sqlStr = sqlStr + "			and m.differencekey = 0 "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " where masteridx = " & masteridx & " and gubundetailname = '상품보관' "
	dbget_TPL.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	sqlStr = sqlStr + " set avgcbm = round((currcbm + prevcbm) / 2, 3), itemno = 1 "
	sqlStr = sqlStr + " where masteridx = " & masteridx & " and gubundetailname = '상품보관' "
	dbget_TPL.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	sqlStr = sqlStr + " set totPrice = round(avgcbm * unitPrice * itemno, 0) "
	sqlStr = sqlStr + " where masteridx = " & masteridx & " and gubundetailname = '상품보관' "
	dbget_TPL.Execute sqlStr

	sqlStr = " update a "
	sqlStr = sqlStr + " set a.totPrice = round(b.totPrice * 0.15, 0), a.unitprice = round(b.totPrice * 0.15, 0), a.itemno = 1 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		[db_threepl].[dbo].[tbl_tpl_jungsan_detail] a "
	sqlStr = sqlStr + "		join [db_threepl].[dbo].[tbl_tpl_jungsan_detail] b "
	sqlStr = sqlStr + "		on "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and a.masteridx = " & masteridx
	sqlStr = sqlStr + "			and a.masteridx = b.masteridx "
	sqlStr = sqlStr + "			and a.gubundetailname = '부자재/작업공간' "
	sqlStr = sqlStr + "			and b.gubundetailname = '상품보관' "
	dbget_TPL.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_master] "
	sqlStr = sqlStr + " set st_totalcash = ( "
	sqlStr = sqlStr + "		select sum(totPrice) "
	sqlStr = sqlStr + "		from "
	sqlStr = sqlStr + "		[db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	sqlStr = sqlStr + "		where masteridx = " & masteridx & " and gubuncd = 'cbm' "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " where idx = " & masteridx
	dbget_TPL.Execute sqlStr

end function

function fnUpdateEtcSummary(masteridx)
    dim sqlStr

    sqlStr = " update d "
	sqlStr = sqlStr + " set d.totPrice = T.totPrice, d.itemno = T.totitemno "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		[db_threepl].[dbo].[tbl_tpl_jungsan_detail] d "
	sqlStr = sqlStr + "		join ( "
	sqlStr = sqlStr + "			select gubuncd, gubunname, gubundetailname, typename, unitprice, sum(itemno) as totitemno, sum(totPrice) as totPrice "
	sqlStr = sqlStr + "			from "
	sqlStr = sqlStr + "			[db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
	sqlStr = sqlStr + "			where masteridx = " & masteridx
	sqlStr = sqlStr + "			group by gubuncd, gubunname, gubundetailname, typename, unitprice "
	sqlStr = sqlStr + "		) T "
	sqlStr = sqlStr + "		on "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and d.gubuncd = T.gubuncd "
	sqlStr = sqlStr + "			and d.gubunname = T.gubunname "
	sqlStr = sqlStr + "			and d.gubundetailname = T.gubundetailname "
	sqlStr = sqlStr + "			and d.typename = T.typename "
	sqlStr = sqlStr + "			and d.unitprice = T.unitprice "
	sqlStr = sqlStr + " where masteridx = " & masteridx
	dbget_TPL.Execute sqlStr

	sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_detail](masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, itemno, totPrice) "
	sqlStr = sqlStr + " select T.masteridx, T.gubuncd, T.gubunname, T.gubundetailname, T.typename, T.unitprice, totitemno, T.totPrice "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		( "
	sqlStr = sqlStr + "			select masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, sum(itemno) as totitemno, sum(totPrice) as totPrice "
	sqlStr = sqlStr + "			from "
	sqlStr = sqlStr + "			[db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
	sqlStr = sqlStr + "			where masteridx = " & masteridx
	sqlStr = sqlStr + "			group by masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice "
	sqlStr = sqlStr + "		) T "
	sqlStr = sqlStr + "		left join [db_threepl].[dbo].[tbl_tpl_jungsan_detail] d "
	sqlStr = sqlStr + "		on "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and d.masteridx = T.masteridx "
	sqlStr = sqlStr + "			and d.gubuncd = T.gubuncd "
	sqlStr = sqlStr + "			and d.gubunname = T.gubunname "
	sqlStr = sqlStr + "			and d.gubundetailname = T.gubundetailname "
	sqlStr = sqlStr + "			and d.typename = T.typename "
	sqlStr = sqlStr + "			and d.unitprice = T.unitprice "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + "		d.idx is NULL "
	dbget_TPL.Execute sqlStr

	sqlStr = " update d "
	sqlStr = sqlStr + " set d.totPrice = 0, d.itemno = 0 "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "		[db_threepl].[dbo].[tbl_tpl_jungsan_detail] d "
	sqlStr = sqlStr + "		left join ( "
	sqlStr = sqlStr + "			select gubuncd, gubunname, gubundetailname, typename, unitprice, sum(itemno) as totitemno, sum(totPrice) as totPrice "
	sqlStr = sqlStr + "			from "
	sqlStr = sqlStr + "			[db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
	sqlStr = sqlStr + "			where masteridx = " & masteridx
	sqlStr = sqlStr + "			group by gubuncd, gubunname, gubundetailname, typename, unitprice "
	sqlStr = sqlStr + "		) T "
	sqlStr = sqlStr + "		on "
	sqlStr = sqlStr + "			1 = 1 "
	sqlStr = sqlStr + "			and d.gubuncd = T.gubuncd "
	sqlStr = sqlStr + "			and d.gubunname = T.gubunname "
	sqlStr = sqlStr + "			and d.gubundetailname = T.gubundetailname "
	sqlStr = sqlStr + "			and d.typename = T.typename "
	sqlStr = sqlStr + "			and d.unitprice = T.unitprice "
	sqlStr = sqlStr + " where masteridx = " & masteridx & " and d.gubuncd <> 'cbm' and T.gubuncd is NULL "
	dbget_TPL.Execute sqlStr

	sqlStr = " update "
	sqlStr = sqlStr + " [db_threepl].[dbo].[tbl_tpl_jungsan_master] "
	sqlStr = sqlStr + " set io_totalcash = ( "
	sqlStr = sqlStr + "		select sum(totPrice) "
	sqlStr = sqlStr + "		from "
	sqlStr = sqlStr + "		[db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	sqlStr = sqlStr + "		where masteridx = " & masteridx & " and gubuncd = 'ipchul' "
	sqlStr = sqlStr + " ) "
	sqlStr = sqlStr + " where idx = " & masteridx
	dbget_TPL.Execute sqlStr
end function


select case mode
    case "tplbatchprocess"
        '// 3PL 업체 물류대행비 정산내역 작성
        if ((tplcompanyid="") or (yyyy1="") or (mm1="") or (differencekey="")) then
            response.write "<script>alert('파라미터 오류!!');</script>"
            dbget_TPL.Close : dbget.close()	:	response.End
        end if

        if (Cstr(differencekey)<>"0") then
            title = yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")"
        else
            title = yyyy1 + "년 " + mm1 + "월 정산"
        end if

        finishflag = 0
        taxtype = "01"
        itemvatYn = "Y"

	    sqlStr = " SELECT groupid, company_name, jungsan_hp, jungsan_email "
	    sqlStr = sqlStr + " from db_partner.dbo.tbl_partner "
	    sqlStr = sqlStr + " where id='" & replace(tplcompanyid, "tpl", "3pl") & "' and userdiv = '903' "
        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        If not rsget.EOF Then
            groupid = rsget("groupid")
            company_name = rsget("company_name")
            jungsan_hp = rsget("jungsan_hp")
            jungsan_email = rsget("jungsan_email")
        End If
        rsget.Close

        sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_master](tplcompanyid, yyyymm, title, st_totalcash, io_totalcash, et_totalcash, finishflag, taxtype, differencekey, groupid, itemvatYn, company_name, jungsan_hp, jungsan_email) "
        sqlStr = sqlStr + " values('" & tplcompanyid & "', '" & yyyymm & "', '" & title & "', 0, 0, 0, " & finishflag & ", '" & taxtype & "', '" & differencekey & "', '" & groupid & "', '" & itemvatYn & "', '" & company_name & "', '" & jungsan_hp & "', '" & jungsan_email & "') "
        ''response.write sqlStr
        dbget_TPL.Execute sqlStr

		sqlStr = "select IDENT_CURRENT('[db_threepl].[dbo].[tbl_tpl_jungsan_master]') as idx "
		rsget_TPL.Open sqlStr,dbget_TPL,1
		if Not rsget_TPL.Eof then
			masteridx = rsget_TPL("idx")
		end if
		rsget_TPL.Close

    	if (differencekey = 0) then
IF application("Svr_Info")<>"Dev" THEN
    	    sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_Directorate_SKU_brand_TPL] '" & yyyymm & "', '" & tplcompanyid & "' "
            db3_dbget.CursorLocation = adUseClient
            db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
            i = 0
            tmpSql = ""
            If not db3_rsget.EOF Then
                tmpInsSql = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_cbm](masteridx, itemgubun, itemid, itemoption, itemname, itemoptionname, itemno, cbmX, cbmY, cbmZ) "
                Do until db3_rsget.EOF
                    '// insert value 가 너무 많으면 오류가 발생한다.
                    '// 300개 단위로 insert 문 추가
                    if (i = 0) then
                        tmpSql = tmpInsSql + " values(" & masteridx & ", '" & db3_rsget("itemgubun") & "', '" & db3_rsget("itemid") & "', '" & db3_rsget("itemoption") & "', '" & db3_rsget("itemname") & "', '" & db3_rsget("itemoptionname") & "', '" & db3_rsget("realstock") & "', '" & db3_rsget("cbmX") & "', '" & db3_rsget("cbmY") & "', '" & db3_rsget("cbmZ") & "')"
                    elseif (i mod 300) = 0 then
                        tmpSql = tmpSql + "; " + tmpInsSql + " values(" & masteridx & ", '" & db3_rsget("itemgubun") & "', '" & db3_rsget("itemid") & "', '" & db3_rsget("itemoption") & "', '" & db3_rsget("itemname") & "', '" & db3_rsget("itemoptionname") & "', '" & db3_rsget("realstock") & "', '" & db3_rsget("cbmX") & "', '" & db3_rsget("cbmY") & "', '" & db3_rsget("cbmZ") & "')"
                    else
                        tmpSql = tmpSql + " , (" & masteridx & ", '" & db3_rsget("itemgubun") & "', '" & db3_rsget("itemid") & "', '" & db3_rsget("itemoption") & "', '" & db3_rsget("itemname") & "', '" & db3_rsget("itemoptionname") & "', '" & db3_rsget("realstock") & "', '" & db3_rsget("cbmX") & "', '" & db3_rsget("cbmY") & "', '" & db3_rsget("cbmZ") & "')"
                    end if

	                db3_rsget.MoveNext
				    i = i + 1
			    Loop
            End If
            db3_rsget.Close

            if (tmpSql <> "") then
                ''response.write tmpSql
                dbget_TPL.Execute tmpSql
            end if
END IF
			'// 임대비
        	sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_detail](masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, avgcbm, currcbm, prevcbm, itemno, totPrice, mastercode, comment) "
        	sqlStr = sqlStr + " select " & masteridx & ", c1.comm_cd, c1.comm_name, c2.comm_name, '', IsNull(c3.comm_price, c2.comm_price), 0, 0, 0, 0, 0, '', (case when c2.comm_name = '부자재/작업공간' then '상품적재 임대금액의 15%' else '' end) "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c1 "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c2 on c1.comm_cd = c2.comm_group "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c3 on c2.comm_cd = c3.comm_group "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and c1.comm_isDel = 'N' "
            sqlStr = sqlStr + " 	and c1.dispyn = 'Y' "
            sqlStr = sqlStr + " 	and c1.comm_group = 'gubun' "
            sqlStr = sqlStr + " 	and c1.comm_cd = 'cbm' "
            sqlStr = sqlStr + " order by c1.sortno, c2.sortno, c3.sortno "
            dbget_TPL.Execute sqlStr

            Call fnUpdateCbmSummary(masteridx)

            '// 입출고
        	sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_detail](masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, avgcbm, currcbm, prevcbm, itemno, totPrice, mastercode, comment) "
        	sqlStr = sqlStr + " select " & masteridx & ", c1.comm_cd, c1.comm_name, c2.comm_name, IsNull(c3.comm_name, ''), IsNull(c3.comm_price, c2.comm_price), 0, 0, 0, 0, 0, '', '' "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c1 "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c2 on c1.comm_cd = c2.comm_group "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c3 on c2.comm_cd = c3.comm_group "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and c1.comm_isDel = 'N' "
            sqlStr = sqlStr + " 	and c1.dispyn = 'Y' "
            sqlStr = sqlStr + " 	and c1.comm_group = 'gubun' "
            sqlStr = sqlStr + " 	and c1.comm_cd = 'ipchul' "
            sqlStr = sqlStr + " order by c1.sortno, c2.sortno, c3.sortno "
            dbget_TPL.Execute sqlStr

            '// 입출고 : 입고
    	    sqlStr = " select "
            sqlStr = sqlStr + " 	m.code, count(distinct (d.iitemgubun + convert(varchar, d.itemid) + d.itemoption)) as sku, sum(d.itemno) as pcs "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_acount_storage_master] m "
            sqlStr = sqlStr + " 	join [db_storage].[dbo].[tbl_acount_storage_detail] d on m.code = d.mastercode "
            sqlStr = sqlStr + " 	left join [db_partner].[dbo].[tbl_partner] p on p.id = d.imakerid "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and p.tplcompanyid = '" & tplcompanyid & "' "
            sqlStr = sqlStr + " 	and m.executedt >= '" & yyyymm & "-01' "
            sqlStr = sqlStr + " 	and m.executedt < DateAdd(m, 1, '" & yyyymm & "-01') "
            sqlStr = sqlStr + " 	and m.ipchulflag = 'I' "
            sqlStr = sqlStr + " 	and m.deldt is NULL "
            sqlStr = sqlStr + " 	and d.deldt is NULL "
            sqlStr = sqlStr + " 	and d.itemno > 0 "		'// 반품제외
            sqlStr = sqlStr + " group by "
            sqlStr = sqlStr + " 	m.code "
            sqlStr = sqlStr + " order by "
            sqlStr = sqlStr + " 	m.code "
            dbget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
            i = 0
            tmpSql = ""
            If not rsget.EOF Then
                tmpInsSql = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_etc](masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, itemno, totPrice, mastercode, comment) "
                Do until rsget.EOF
                    '// insert value 가 너무 많으면 오류가 발생한다.
                    '// 300개 단위로 insert 문 추가
                    if (i = 0) then
                        tmpSql = tmpInsSql + " values(" & masteridx & ", 'ipchul', '입출고', '물류입고', '1PLT', '2500', '1', '2500', '" & rsget("code") & "', 'SKU " & rsget("sku") & " PCS " & rsget("pcs") & "')"
                    elseif (i mod 300) = 0 then
                        tmpSql = tmpSql + "; " + tmpInsSql + " values(" & masteridx & ", 'ipchul', '입출고', '물류입고', '1PLT', '2500', '1', '2500', '" & rsget("code") & "', 'SKU " & rsget("sku") & " PCS " & rsget("pcs") & "')"
                    else
                        tmpSql = tmpSql + " , (" & masteridx & ", 'ipchul', '입출고', '물류입고', '1PLT', '2500', '1', '2500', '" & rsget("code") & "', 'SKU " & rsget("sku") & " PCS " & rsget("pcs") & "')"
                    end if

	                rsget.MoveNext
				    i = i + 1
			    Loop
            End If
            rsget.Close

            if (tmpSql <> "") then
                ''response.write tmpSql
                dbget_TPL.Execute tmpSql
            end if

            '// 기타
        	sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_detail](masteridx, gubuncd, gubunname, gubundetailname, typename, unitprice, avgcbm, currcbm, prevcbm, itemno, totPrice, mastercode, comment) "
        	sqlStr = sqlStr + " select " & masteridx & ", c1.comm_cd, c1.comm_name, c2.comm_name, IsNull(c3.comm_name, ''), IsNull(c3.comm_price, c2.comm_price), 0, 0, 0, 0, 0, '', '' "
            sqlStr = sqlStr + " from "
            sqlStr = sqlStr + " 	[db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c1 "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c2 on c1.comm_cd = c2.comm_group "
            sqlStr = sqlStr + " 	left join [db_threepl].[dbo].[tbl_tpl_jungsan_comm_code] c3 on c2.comm_cd = c3.comm_group "
            sqlStr = sqlStr + " where "
            sqlStr = sqlStr + " 	1 = 1 "
            sqlStr = sqlStr + " 	and c1.comm_isDel = 'N' "
            sqlStr = sqlStr + " 	and c1.dispyn = 'Y' "
            sqlStr = sqlStr + " 	and c1.comm_group = 'gubun' "
            sqlStr = sqlStr + " 	and c1.comm_cd = 'etc' "
            sqlStr = sqlStr + " order by c1.sortno, c2.sortno, c3.sortno "
            dbget_TPL.Execute sqlStr

            Call fnUpdateEtcSummary(masteridx)
        end if

        response.write "<script>alert('OK');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"
        dbget_TPL.Close : dbget.close()	:	response.End
    case "dellall"
        Call fnExitIfNotEditState(masteridx, refer)

        sqlStr = "delete from [db_threepl].[dbo].[tbl_tpl_jungsan_master] "
	    sqlStr = sqlStr + " where idx=" + CStr(masteridx)
        dbget_TPL.Execute sqlStr

        sqlStr = "delete from [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] "
	    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
        dbget_TPL.Execute sqlStr

        sqlStr = "delete from [db_threepl].[dbo].[tbl_tpl_jungsan_detail] "
	    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
        dbget_TPL.Execute sqlStr

        sqlStr = "delete from [db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
	    sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
        dbget_TPL.Execute sqlStr

    case "etcadd"
        Call fnExitIfNotEditState(masteridx, refer)

        select case gubun
            case "cbm"
                sqlStr = " insert into [db_threepl].[dbo].[tbl_tpl_jungsan_cbm](masteridx, itemname, itemno, cbmX, cbmY, cbmZ, itemoptionname) "
                sqlStr = sqlStr + " values(" & masteridx & ", '" &itemname & "', '" & itemno & "', '" & cbmX & "', '" & cbmY & "', '" & cbmZ & "', '')"
                dbget_TPL.Execute sqlStr

                Call fnUpdateCbmSummary(masteridx)
            case else
                '//
        end select

    case "deldetail"
        Call fnExitIfNotEditState(masteridx, refer)

        select case gubun
            case "cbm"
                sqlStr = " delete from [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] "
                sqlStr = sqlStr + " where idx = " & idx
                dbget_TPL.Execute sqlStr

                Call fnUpdateCbmSummary(masteridx)
            case else
                sqlStr = " delete from [db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
                sqlStr = sqlStr + " where idx = " & idx
                dbget_TPL.Execute sqlStr

                Call fnUpdateEtcSummary(masteridx)
        end select
    case "modidetail"
        Call fnExitIfNotEditState(masteridx, refer)

        select case gubun
            case "cbm"
                sqlStr = " update [db_threepl].[dbo].[tbl_tpl_jungsan_cbm] "
                sqlStr = sqlStr + " set itemno = '" & itemno & "', cbmX = '" & cbmX & "', cbmY = '" & cbmY & "', cbmZ = '" & cbmZ & "' "
                sqlStr = sqlStr + " where idx = " & idx
                dbget_TPL.Execute sqlStr

                Call fnUpdateCbmSummary(masteridx)
            case else
                sqlStr = " update [db_threepl].[dbo].[tbl_tpl_jungsan_etc] "
                sqlStr = sqlStr + " set itemno = '" & itemno & "', unitprice = '" & unitprice & "', typename = '" & itypename & "', totPrice = " & itemno & "*" & unitprice & " "
                sqlStr = sqlStr + " where idx = " & idx
                dbget_TPL.Execute sqlStr

                Call fnUpdateEtcSummary(masteridx)
        end select
    case else
        response.write "잘못된 접근입니다."
end select

%>
<% if mode="ipkumfinish" then %>
<script language="javascript">
alert('저장 되었습니다.');
window.close();
</script>
<% else %>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
