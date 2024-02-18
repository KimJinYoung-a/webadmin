<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

'==============================================================================
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim masteridx, ordercode, detailidx, baljuno
dim songjangdiv, workgroup, ems, epostmilitary
dim companyid, baljutype, reguserid
dim companygubun, pickingStationCd

dim comment, errorcd



'==============================================================================
mode		= request("mode")

masteridx	= request("masteridx")
ordercode	= request("ordercode")
detailidx	= request("detailidx")
baljuno		= request("baljuno")

comment		= request("comment")
errorcd		= request("errorcd")

songjangdiv		= request("songjangdiv")
workgroup		= request("workgroup")
ems				= request("ems")
epostmilitary	= request("epostmilitary")
pickingStationCd	= request("pickingStationCd")

reguserid	= session("ssBctUId")



'==============================================================================
dim dummymasteridx
dummymasteridx = masteridx

masteridx = split(masteridx,"|")
ordercode = split(ordercode,"|")
detailidx = split(detailidx,"|")
baljuno = split(baljuno,"|")

comment = split(comment,"|")
errorcd = split(errorcd,"|")



'==============================================================================
dim sqlStr,i
dim iid
dim obaljudate
dim errcode
dim masteridxlist
dim differencekey
dim prevmasteridx

dim buf, baljuidarr, baljucodearr, baljunum, baljudate, cnt



if mode="arr" then
	''유효성체크.

	dummymasteridx = Mid(dummymasteridx,2,Len(dummymasteridx))
	dummymasteridx = replace(dummymasteridx,"|",",")

	'==========================================================================
    '마이너스 주문이 있는지 확인
    sqlStr = " select distinct m.baljucode "
    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master m,"
    sqlStr = sqlStr + " [db_storage].[dbo].tbl_ordersheet_detail d "
    sqlStr = sqlStr + " where m.idx = d.masteridx "
    sqlStr = sqlStr + " and m.deldt is null "
    sqlStr = sqlStr + " and d.deldt is null "
    sqlStr = sqlStr + " and m.idx in (" + CStr(dummymasteridx) + ") "
    sqlStr = sqlStr + " and d.baljuitemno < 0 "

	rsget.Open sqlStr, dbget, 1

	buf = ""
	if Not rsget.Eof then
		do until rsget.eof
			if (buf = "") then
			        buf = rsget("baljucode")
			else
			        buf = buf + "," + rsget("baljucode")
			end if
			rsget.movenext
		loop
	end if
	rsget.close

	if (buf <> "") then
	        response.write "<script>alert('주문중에 마이너스 주문이 있는 주문(" + buf + ")이 있습니다.');</script>"
	        response.write "<script>history.back();</script>"
	        dbget.close()	:	response.End
	end if


	On Error Resume Next
	dbget.beginTrans


	'==========================================================================
    '상태변경
    If Err.Number = 0 Then
    	errcode = "000"

	    sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master "
	    sqlStr = sqlStr + " set statecd='1' "
	    sqlStr = sqlStr + " where 1 = 1 "
	    sqlStr = sqlStr + " and statecd = '0' "
	    sqlStr = sqlStr + " and idx in (" + CStr(dummymasteridx) + ") "
	    dbget.Execute sqlStr
	end if

	'==========================================================================
    '실출고수량 0 Reset
    If Err.Number = 0 Then
    	errcode = "001"

	    sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail "
	    sqlStr = sqlStr + " set realitemno=0 "
	    sqlStr = sqlStr + " where 1 = 1 "
	    sqlStr = sqlStr + " and masteridx in (" + CStr(dummymasteridx) + ") "
	    dbget.Execute sqlStr
	end if

	'==========================================================================
	'출고지시입력
    If Err.Number = 0 Then
    	errcode = "002"

	    sqlStr = " select max(isnull(baljunum,0)) as maxbaljunum, convert(varchar,getdate(),109) as baljudate "
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju "
	    ''response.write sqlStr & "<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			baljunum = rsget("maxbaljunum") + 1
			baljudate = rsget("baljudate")
		end if
		rsget.close

		sqlStr = " select (IsNull(max(differencekey), 0) + 1) as differencekey "
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju "
		sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21) "
	    ''response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			differencekey = rsget("differencekey")
		rsget.close

	    sqlStr = " select baljuid, baljucode "
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master "
	    sqlStr = sqlStr + " where 1 = 1 "
	    sqlStr = sqlStr + " and idx in (" + CStr(dummymasteridx) + ") "
	    ''response.write sqlStr & "<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			baljuidarr = ""
			baljucodearr = ""
			do until rsget.eof
				baljuidarr = baljuidarr + rsget("baljuid") + "|"
				baljucodearr = baljucodearr + rsget("baljucode") + "|"
				rsget.movenext
			loop
		end if
		rsget.close

		baljuidarr = split(baljuidarr,"|")
		baljucodearr = split(baljucodearr,"|")

		cnt = ubound(baljuidarr)

		for i = 0 to cnt
	        if (baljuidarr(i) <> "") then
	            sqlStr = " insert into [db_storage].[dbo].tbl_shopbalju(baljunum, baljuid, baljucode, baljudate, differencekey, workgroup, songjangdiv, pickingStationCd) "
	            sqlStr = sqlStr + " values(" + CStr(baljunum) + ", '" + CStr(baljuidarr(i)) + "', '" + CStr(baljucodearr(i)) + "', convert(datetime,'" + CStr(baljudate) + "',109), " + CStr(differencekey) + ", '" + CStr(workgroup) + "', " + CStr(songjangdiv) + ", '" + CStr(pickingStationCd) + "') "
	    		''response.write sqlStr & "<Br>"
	            rsget.Open sqlStr, dbget, 1
	        end if
		next
	end if

	'==========================================================================
	'출고지시수량 입력
	'출고지시 제외 사유입력
    If Err.Number = 0 Then
    	errcode = "003"

		for i=0 to Ubound(detailidx)
			if detailidx(i)<>"" then
				sqlStr = " update d "
				sqlStr = sqlStr + " set "
				sqlStr = sqlStr + " 	d.comment = '" & CStr(comment(i)) & "' "
				sqlStr = sqlStr + " 	, d.realbaljuitemno = " & CStr(baljuno(i)) & " "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_ordersheet_master m "
				sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		m.idx = d.masteridx "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	1 = 1 "
				sqlStr = sqlStr + " 	and d.idx = " & CStr(detailidx(i)) & " "
				'response.write sqlStr + "<br>"

				rsget.Open sqlStr,dbget,1
			end if
		next
	end if

	'==========================================================================
    '재고 변경 (오프 접수 -> 오프상품준비) : 오프접수, 상품준비 수량 재계산
    If Err.Number = 0 Then
    	errcode = "004"

	    sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
	    dbget.Execute sqlStr
	end if

	'==========================================================================
    '출고지시된 내역 한정판매 재설정("실제출고지시수량" 만큼 더해준다.)
    If Err.Number = 0 Then
    	errcode = "005"

	    '옵션코드 없을때
	    sqlstr = " update [db_item].[dbo].tbl_item "
	    sqlstr = sqlstr + " set limitsold=(case when limitno<limitsold + IsNull(T.itemno,0) then limitno else limitsold + IsNull(T.itemno,0) end) "
	    sqlstr = sqlstr + " from ( "
	    sqlstr = sqlstr + "     select sum(d.realbaljuitemno) as itemno, d.itemid "
	    sqlstr = sqlstr + "     from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
	    sqlstr = sqlstr + "     where d.masteridx in (" + CStr(dummymasteridx) + ") "
	    sqlstr = sqlstr + "     and d.itemid=i.itemid "
	    sqlstr = sqlstr + "     and d.deldt is null "
	    sqlstr = sqlstr + "     and d.itemgubun = '10' "
	    sqlstr = sqlstr + "     and d.itemoption = '0000' "
	    sqlstr = sqlstr + "     and i.limityn='Y' "
	    sqlstr = sqlstr + "     and i.mwdiv<>'U'"
	    sqlstr = sqlstr + "     group by d.itemid "
	    sqlstr = sqlstr + " ) as T "
	    sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid "
	    rsget.Open sqlStr, dbget, 1

		'옵션코드 있을때
	    sqlstr = " update [db_item].[dbo].tbl_item_option "
	    sqlstr = sqlstr + " set optlimitsold=(case when optlimitno<optlimitsold+IsNull(T.itemno,0) then optlimitno else optlimitsold+IsNull(T.itemno,0) end) "
	    sqlstr = sqlstr + " from ( "
	    sqlstr = sqlstr + "     select sum(d.realbaljuitemno) as itemno, d.itemid, d.itemoption "
	    sqlstr = sqlstr + "     from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
	    sqlstr = sqlstr + "     where d.masteridx in (" + CStr(dummymasteridx) + ") "
	    sqlstr = sqlstr + "     and d.itemid=i.itemid "
	    sqlstr = sqlstr + "     and d.deldt is null "
	    sqlstr = sqlstr + "     and d.itemgubun = '10' "
	    sqlstr = sqlstr + "     and d.itemoption <> '0000' "
	    sqlstr = sqlstr + "     and i.limityn='Y' "
	    sqlstr = sqlstr + "     and i.mwdiv<>'U'"
	    sqlstr = sqlstr + "     group by d.itemid, d.itemoption "
	    sqlstr = sqlstr + " ) as T "
	    sqlstr = sqlstr + " where [db_item].[dbo].tbl_item_option.itemid=T.itemid and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption "
	    rsget.Open sqlStr, dbget, 1

		'한정수량 재계산
		sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
		sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
		sqlStr = sqlStr + " from (" + VBCrlf
		sqlStr = sqlStr + " 	select itemid, sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
		sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
		sqlStr = sqlStr + " 	where itemid in (select itemid from [db_storage].[dbo].tbl_ordersheet_detail where masteridx in (" + CStr(dummymasteridx) + ") and itemoption <> '0000') " + VBCrlf
		sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
		sqlStr = sqlStr + " 	group by itemid " + VBCrlf
		sqlStr = sqlStr + " ) T" + VBCrlf
		sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid= T.itemid " + VBCrlf
		sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"
		rsget.Open sqlStr, dbget, 1
	end if

	If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n관리자 문의 요망 (에러코드 : " + CStr(errcode) + ")');</script>"
			response.write "데이타를 저장하는 도중에 에러가 발생하였습니다<br>관리자 문의 요망 (에러코드 : " + CStr(errcode) + ")"
	        ''response.write "<script>history.back()</script>"
	        dbget.close()	:	response.End
	End If
	on error Goto 0

end if


%>

<script language="javascript">
alert('출고지시서가 생성 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
