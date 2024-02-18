<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim chargeid, shopid,vatcode,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,mode, idx, midx ,songjangdiv, songjangno
dim sellcash,suplycash,shopbuyprice,itemno, divcode ,statecd ,execdate ,scheduledt
dim i,cnt,sqlStr ,isWaitState ,currState, idxArr
	mode = requestCheckVar(request("mode"),32)
	idx = requestCheckVar(request("idx"),10)
	idxArr = requestCheckVar(request("idx"),5000)
	midx = requestCheckVar(request("midx"),10)
	chargeid = requestCheckVar(request("chargeid"),32)
	shopid = requestCheckVar(request("shopid"),32)
	vatcode = requestCheckVar(request("vatcode"),3)
	divcode = requestCheckVar(request("divcode"),3)
	sellcash = requestCheckVar(request("sellcash"),20)
	suplycash = requestCheckVar(request("suplycash"),20)
	shopbuyprice = requestCheckVar(request("shopbuyprice"),20)
	itemno = requestCheckVar(request("itemno"),10)
	statecd = requestCheckVar(request("statecd"),10)
	scheduledt = requestCheckVar(request("scheduledt"),30)
	songjangdiv = requestCheckVar(request("songjangdiv"),2)
	songjangno  = requestCheckVar(request("songjangno"),32)
	execdate = requestCheckVar(request("execdate"),30)

	''입고예정일
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)

	''실제입고일
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="delmaster" then

	sqlStr = "select top 1 idx, statecd from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		isWaitState = (rsget("statecd")<1)
	end if
	rsget.Close

	if Not isWaitState then
		response.write "<script type='text/javascript'>alert('현재 입고대기 상태가 아닙니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set deleteyn='Y'"  + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr, dbget, 1

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,shopid

	response.write "<script type='text/javascript'>"
	response.write "alert('삭제 되었습니다.');"
	response.write "location.replace('" + refer + "');"
	response.write "</script>"
	dbget.close()	:	response.End

elseif mode="nextstep" then

	''입고 확인 - 가맹점 입고 확인 상태에서만 진행 가능
	sqlStr = "select top 1 idx, statecd from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
		    ''입고대기 에서도 확정 가능.. 변경
			isWaitState = (rsget("statecd")="7") or (rsget("statecd")="0")
		end if
	rsget.Close

	if Not isWaitState then
		response.write "<script type='text/javascript'>alert('현재 가맹점 입고확인 상태가 아닙니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set statecd='8'"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmdate=getdate()"  + vbCrlf
	sqlStr = sqlStr + " ,upcheconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

	response.write "<script type='text/javascript'>"
	response.write "alert('입고 확정 되었습니다.');"
	response.write "location.replace('" + refer + "');"
	response.write "</script>"
	dbget.close()	:	response.End

elseif mode="upchechulgoproc" then

    ''입고 확인 - 가맹점 입고 확인 상태에서만 진행 가능
	sqlStr = "select top 1 idx, statecd from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			currState = rsget("statecd")
		end if
	rsget.Close

	if (currState<>-1) then
		response.write "<script type='text/javascript'>alert('현재 가맹점 입고요청확인 상태가 아닙니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set statecd=0"  + vbCrlf
	sqlStr = sqlStr + " ,songjangdiv='" + songjangdiv + "'"
	sqlStr = sqlStr + " ,songjangno='" + songjangno + "'"
	sqlStr = sqlStr + " ,scheduledate='" + CStr(scheduledt) + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1

''  Join사용
''  sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + VbCrlf
''	sqlStr = sqlStr + " set songjangname=IsNULL(T.divname,'')" + VbCrlf
''	sqlStr = sqlStr + " from [db_order].[10x10].tbl_songjang_div T" + VbCrlf
''	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)
''	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_master.songjangdiv=T.divcd"
''
''	dbget.Execute(sqlStr)

	response.write "<script type='text/javascript'>"
	response.write "alert('발송 처리 되었습니다.');"
	response.write "location.replace('/common/offshop/shop_ipchullist.asp?menupos=504');"
	response.write "</script>"
    dbget.close()	:	response.End

elseif mode="modimaster" then

	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set chargeid='" + chargeid + "'" + vbCrlf
	sqlStr = sqlStr + " ,shopid='" + shopid + "'" + vbCrlf
	sqlStr = sqlStr + " ,divcode='" + divcode + "'" + vbCrlf
	sqlStr = sqlStr + " ,vatcode='" + vatcode + "'" + vbCrlf
	sqlStr = sqlStr + " ,songjangdiv='" + songjangdiv + "'"
	sqlStr = sqlStr + " ,songjangno='" + songjangno + "'"
	sqlStr = sqlStr + " ,scheduledate='" + CStr(scheduledt) + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

''  Join사용
''	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + VbCrlf
''	sqlStr = sqlStr + " set songjangname=IsNULL(T.divname,'')" + VbCrlf
''	sqlStr = sqlStr + " from [db_order].[10x10].tbl_songjang_div T" + VbCrlf
''	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(idx)
''	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_ipchul_master.songjangdiv=T.divcd"
''
''	dbget.Execute(sqlStr)

elseif mode="detailmodi" then

	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(midx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
		if not rsget.Eof then
		    currState = rsget("statecd")

			if currState>0 then
				response.write "<script type='text/javascript'>alert('현재 입고대기 상태가 아닙니다.');</script>"
				response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
				dbget.close()	:	response.End
			end if
		end if
	rsget.Close

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " set sellcash=" + Cstr(sellcash) + "" + vbCrlf
	sqlStr = sqlStr + " ,suplycash=" + Cstr(suplycash) + "" + vbCrlf
	sqlStr = sqlStr + " ,shopbuyprice=" + Cstr(shopbuyprice) + "" + vbCrlf
	sqlStr = sqlStr + " ,itemno=" + Cstr(itemno) + "" + vbCrlf

	if (currState=-2) then
	    ''입고요청에서만 수량 동일시.
	    sqlStr = sqlStr + " ,reqno=" + Cstr(itemno) + "" + vbCrlf
	end if

	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(midx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(midx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

elseif mode="detaildel" then

	sqlStr = " select top 1 statecd from [db_shop].[dbo].tbl_shop_ipchul_master"
	sqlStr = sqlStr + " where idx=" + CStr(midx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
		if rsget("statecd")>0 then
			response.write "<script type='text/javascript'>alert('현재 입고대기 상태가 아닙니다.');</script>"
			response.write "<script type='text/javascript'>location.replace('" + refer + "');</script>"
			dbget.close()	:	response.End
		end if
	end if
	rsget.Close

	''상세내역 삭제.
	sqlStr = " delete from [db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

	sqlStr = " update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totsell,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
	sqlStr = sqlStr + " ,totalshopbuyprice=IsNULL(T.totshopbuy,0)" + vbCrlf
	sqlStr = sqlStr + " from (" + vbCrlf
	sqlStr = sqlStr + " 	select sum(sellcash*itemno) as totsell " + vbCrlf
	sqlStr = sqlStr + " 	,sum(suplycash*itemno) as totsupp " + vbCrlf
	sqlStr = sqlStr + " 	,sum(shopbuyprice*itemno) as totshopbuy " + vbCrlf
	sqlStr = sqlStr + " 	from " + vbCrlf
	sqlStr = sqlStr + " 	[db_shop].[dbo].tbl_shop_ipchul_detail" + vbCrlf
	sqlStr = sqlStr + " 	where masteridx="  + CStr(midx) + vbCrlf
	sqlStr = sqlStr + " 	and deleteyn='N'" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_ipchul_master.idx=" + CStr(midx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

'/입고확정으로 변경
elseif mode="ipgook" then

    sqlStr = "select top 1 idx, statecd from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		currState = rsget("statecd")
	end if
	rsget.Close

    ''입고확정으로 바로진행
	sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " set statecd='8'"  + vbCrlf
	sqlStr = sqlStr + " ,execdt='" + CStr(execdate) + "'" + vbCrlf
	sqlStr = sqlStr + " ,shopconfirmdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " ,shopconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	dbget.Execute(sqlStr)

    if (currState<7) then
        ''신재고 - 업체 입출고
        sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",1"

		'response.write sqlStr &"<Br>"
        dbget.Execute(sqlStr)
    end if

	''샾재고 업데이트 루틴 (OLD - 차후 삭제)
	OffStockUpdateUpcheIpgoByIdx idx

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,shopid

'/관리자 입고 상태 변경
elseif mode="modistate" then

    sqlStr = "select top 1 idx, statecd from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		currState = rsget("statecd")
	end if
	rsget.Close

	select case statecd

	    case "-2" ''입고요청
	        if (currState>=7) then
			    ''신재고 - 업체 입출고
                sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",-1"

                'response.write sqlStr &"<Br>"
                dbget.Execute(sqlStr)
			end if

			sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
			sqlStr = sqlStr + " set statecd='" + statecd + "'"  + vbCrlf
			sqlStr = sqlStr + " ,execdt=null" + vbCrlf
			sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute(sqlStr)

	    case "-1" ''입고요청 확인
	        if (currState>=7) then

			    ''신재고 - 업체 입출고
                sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",-1"

                'response.write sqlStr &"<Br>"
                dbget.Execute(sqlStr)
			end if

			sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
			sqlStr = sqlStr + " set statecd='" + statecd + "'"  + vbCrlf
			sqlStr = sqlStr + " ,execdt=null" + vbCrlf
			sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute(sqlStr)

		case "0" ''입고대기
		    if (currState>=7) then
			    ''신재고 - 업체 입출고
                sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",-1"

                'response.write sqlStr &"<Br>"
                dbget.Execute(sqlStr)
			end if

			sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
			sqlStr = sqlStr + " set statecd='" + statecd + "'"  + vbCrlf
			sqlStr = sqlStr + " ,execdt=null" + vbCrlf
			sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute(sqlStr)

		case "7" ''매장입고확인
			sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
			sqlStr = sqlStr + " set statecd='" + statecd + "'"  + vbCrlf
			sqlStr = sqlStr + " ,execdt='" + CStr(execdate) + "'" + vbCrlf
			sqlStr = sqlStr + " ,shopconfirmdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " ,shopconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
			sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute(sqlStr)

			if (currState<7) then
			    ''신재고 - 업체 입출고
                sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",1"

                'response.write sqlStr &"<Br>"
                dbget.Execute(sqlStr)
			end if

		case "8" ''입고확정(업체확정)
			sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
			sqlStr = sqlStr + " set statecd='" + statecd + "'"  + vbCrlf
			sqlStr = sqlStr + " ,upcheconfirmdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " ,upcheconfirmuserid='" + session("ssBctId") + "'" + vbCrlf
			sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(idx)

			'response.write sqlStr &"<Br>"
			dbget.execute(sqlStr)

			if (currState<7) then
			    ''신재고 - 업체 입출고
                sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idx) & ",1"

                'response.write sqlStr &"<Br>"
                dbget.Execute(sqlStr)
			end if
		case else
	end select

	''샾재고 업데이트 루틴
	OffStockUpdateUpcheIpgoByIdx idx

	'//기주문 업데이트
	PreOrderUpdateByBrand_off idx,chargeid,shopid

elseif (mode = "modistatemulti") then
	idxArr = Split(idxArr, ",")
	for i = 0 to UBound(idxArr)
		sqlStr = "select top 1 idx, statecd, chargeid,shopid from [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(idxArr(i))

		currState = ""
		''response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			currState = rsget("statecd")
			chargeid = rsget("chargeid")
			shopid = rsget("shopid")
		end if
		rsget.Close

		if (currState <> "7") and (currState <> "8") then
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		end if

		if (currState>=7) then
			''신재고 - 업체 입출고
            sqlStr = "exec db_summary.dbo.sp_Ten_Shop_BrandIpchulUpdateByIdx " & CStr(idxArr(i)) & ",-1"

            ''response.write sqlStr &"<Br>"
            dbget.Execute(sqlStr)
		end if

		sqlStr = "update [db_shop].[dbo].tbl_shop_ipchul_master" + vbCrlf
		sqlStr = sqlStr + " set statecd='0'"  + vbCrlf
		sqlStr = sqlStr + " ,execdt=null" + vbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate()" + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(idxArr(i))

		''response.write sqlStr &"<Br>"
		dbget.execute(sqlStr)

		''샾재고 업데이트 루틴
		OffStockUpdateUpcheIpgoByIdx idxArr(i)

		'//기주문 업데이트
		PreOrderUpdateByBrand_off idxArr(i),chargeid,shopid
	next
else
	response.write mode
	dbget.close()	:	response.End

end if
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
