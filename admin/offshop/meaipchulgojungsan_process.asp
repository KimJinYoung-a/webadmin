<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
dim mode, check, idx, topidx
dim shopid, makerid, yyyy, mm, shopdiv, diffKey

mode    = requestCheckVar(request("mode"),32)
check   = requestCheckVar(request("check"),2048)
shopid  = requestCheckVar(request("shopid"),32)
idx     = requestCheckVar(request("idx"),10)
topidx  = requestCheckVar(request("topidx"),10)
makerid = requestCheckVar(request("makerid"),32)
yyyy    = requestCheckVar(request("yyyy1"),10)
mm      = requestCheckVar(request("mm1"),10)
shopdiv = requestCheckVar(request("shopdiv"),10)
diffKey = requestCheckVar(request("diffKey"),10)

dim sqlStr, i, iid, cnt

dim oTax

if (diffKey="") then diffKey=1
if Not IsNumeric(diffKey) then diffKey=1

if mode="chulgo" then

	'' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = ""
	rsget("totalsum") = 0
	rsget("divcode") = "MC"
	rsget("etcstr") = ""
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	rsget.update
	iid = rsget("idx")
	rsget.close

	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ",m.id, m.socid, m.code, s.baljucode, m.executedt,"
	sqlStr = sqlStr + " 0, m.totalsellcash*-1, m.totalbuycash*-1, m.totalsuplycash*-1"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_ordersheet_master s on m.code=s.alinkcode"
	sqlStr = sqlStr + " where m.id in (" + check + ")"

	rsget.Open sqlStr, dbget, 1


	'' insert sub detail
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash)"
	sqlStr = sqlStr + " select  m.idx," + CStr(iid) + ",'',m.code01,"
	sqlStr = sqlStr + " d.id, d.iitemgubun, d.itemid, d.itemoption, d.iitemname, d.iitemoptionname,"
	sqlStr = sqlStr + " d.imakerid, d.itemno*-1,d.sellcash,d.suplycash,d.buycash"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlStr = sqlStr + " where m.masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and m.code01=d.mastercode"

	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1
elseif mode="witsksell" then
    '' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = ""
	rsget("totalsum") = 0
	rsget("divcode") = "WS"
	rsget("etcstr") = ""
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	rsget.update
	iid = rsget("idx")
	rsget.close

	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ", m.idx, d.shopid, convert(varchar(7),m.yyyymm),"
	sqlStr = sqlStr + " m.makerid, convert(varchar(7),m.yyyymm) + '-01', sum(d.itemno) as totitemcnt,"
	sqlStr = sqlStr + " sum(d.realsellprice*d.itemno) as totsum, sum(d.suplyprice*d.itemno) as realjungsansum, 0"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m, "
    sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
    sqlStr = sqlStr + "     where m.idx=d.masteridx "
    sqlStr = sqlStr + "     and m.idx in  (" + check + ")"
    sqlStr = sqlStr + "     and d.gubuncd='B012' "
    sqlStr = sqlStr + "     and d.shopid='" + shopid + "'"
    sqlStr = sqlStr + " group by m.idx, d.shopid, convert(varchar(7),m.yyyymm), m.makerid, convert(varchar(7),m.yyyymm) + '-01'"

	rsget.Open sqlStr, dbget, 1


	'' insert sub detail : sellprice, orgsellprice
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash, orgsellcash)"
	sqlStr = sqlStr + " select m.idx," + CStr(iid) + ",d.gubuncd, d.orderno, d.detailidx,"
	sqlStr = sqlStr + " d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,"
	sqlStr = sqlStr + " d.makerid,d.itemno,d.realsellprice,0,d.suplyprice,d.sellprice"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_detail d"                                                    ''정산을 재작성 할경우 중복으로 잡힐 수 있음. 정산기준==>주문테이블기준 변경 또는 중복정산 안되게 left join 
	sqlStr = sqlStr + " where m.linkidx=d.masteridx"
	sqlStr = sqlStr + " and m.masteridx=" + CStr(iid)
	sqlStr = sqlStr + " and d.gubuncd='B012' "
	sqlStr = sqlStr + " and d.shopid='" + shopid + "'"

	rsget.Open sqlStr, dbget, 1



	'' update Detail shopsuplyprice
	'' 현재 OFF 상품 가격 기준 정산.
	''

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " set suplycash=IsNULL(T.shopbuyprice,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct d.idx, IsNULL(s.defaultsuplymargin,35) as defaultsuplymargin"
	sqlStr = sqlStr + " ,	( case "
	sqlStr = sqlStr + "			when (i.shopbuyprice=0) and (j.discountprice=0) then convert(int,j.sellprice - j.sellprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "  		when (i.shopbuyprice=0) and (j.discountprice<>0) then convert(int,j.discountprice - j.discountprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "    		else i.shopbuyprice "
	sqlStr = sqlStr + "    		end ) as shopbuyprice "

	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail d"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_designer s "
	sqlStr = sqlStr + " 			on d.makerid=s.makerid and s.shopid='" + shopid + "'"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shopjumun_detail j"
	sqlStr = sqlStr + " 			on d.linkmastercode=j.orderno"
	sqlStr = sqlStr + " 			and d.itemgubun=j.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=j.itemid"
	sqlStr = sqlStr + " 			and d.itemoption=j.itemoption"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_item i"
	sqlStr = sqlStr + " 			on d.itemgubun=i.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=i.shopitemid"
	sqlStr = sqlStr + " 			and d.itemoption=i.itemoption"

	sqlStr = sqlStr + " 	where d.topmasteridx=" + CStr(iid)

	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail.idx=T.idx"

	rsget.Open sqlStr, dbget, 1

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

elseif mode="witsksell_old" then
'''이전  OFF 정산 테이블
	'' insert master
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = shopid
	rsget("title") = ""
	rsget("totalsum") = 0
	rsget("divcode") = "WS"
	rsget("etcstr") = ""
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

	rsget.update
	iid = rsget("idx")
	rsget.close


	'' insert sub master
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_submaster "
	sqlStr = sqlStr + " (masteridx,linkidx,shopid,code01,code02,execdate, "
	sqlStr = sqlStr + " totalcount,totalsellcash,totalbuycash,totalsuplycash)"
	sqlStr = sqlStr + " select " + CStr(iid) + ", m.idx, m.shopid, convert(varchar(7),m.yyyymm),"
	sqlStr = sqlStr + " m.jungsanid, convert(varchar(7),m.yyyymm) + '-01', m.totitemcnt,"
	sqlStr = sqlStr + " m.totsum, m.realjungsansum, 0"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_jungsanmaster m"
	sqlStr = sqlStr + " where idx in  (" + check + ")"

	rsget.Open sqlStr, dbget, 1


	'' insert sub detail : sellprice, orgsellprice
	sqlStr = " insert into [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail "
	sqlStr = sqlStr + " (masteridx, topmasteridx, linkbaljucode, linkmastercode, linkdetailidx,"
	sqlStr = sqlStr + " itemgubun, itemid, itemoption, itemname, itemoptionname,"
	sqlStr = sqlStr + " makerid, itemno, sellcash, suplycash, buycash, orgsellcash)"
	sqlStr = sqlStr + " select m.idx," + CStr(iid) + ",d.jungsangubun, d.orderno, d.idx,"
	sqlStr = sqlStr + " d.itemgubun,d.itemid,d.itemoption,d.itemname,d.itemoptionname,"
	sqlStr = sqlStr + " d.makerid,d.itemno,d.realsellprice,0,d.suplyprice,d.sellprice"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_fran_meachuljungsan_submaster m,"
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_jungsandetail d"
	sqlStr = sqlStr + " where m.linkidx=d.masteridx"
	sqlStr = sqlStr + " and m.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1


	'' update Detail shopsuplyprice
	'' 현재 OFF 상품 가격 기준 정산.
	''

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " set suplycash=IsNULL(T.shopbuyprice,0)"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " 	select distinct d.idx, IsNULL(s.defaultsuplymargin,35) as defaultsuplymargin"
	sqlStr = sqlStr + " ,	( case "
	sqlStr = sqlStr + "			when (i.shopbuyprice=0) and (j.discountprice=0) then convert(int,j.sellprice - j.sellprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "  		when (i.shopbuyprice=0) and (j.discountprice<>0) then convert(int,j.discountprice - j.discountprice*IsNULL(s.defaultsuplymargin,35)/100)"
	sqlStr = sqlStr + "    		else i.shopbuyprice "
	sqlStr = sqlStr + "    		end ) as shopbuyprice "

	sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail d"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_designer s "
	sqlStr = sqlStr + " 			on d.makerid=s.makerid and s.shopid='" + shopid + "'"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shopjumun_detail j"
	sqlStr = sqlStr + " 			on d.linkmastercode=j.orderno"
	sqlStr = sqlStr + " 			and d.itemgubun=j.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=j.itemid"
	sqlStr = sqlStr + " 			and d.itemoption=j.itemoption"
	sqlStr = sqlStr + " 		left join [db_shop].[dbo].tbl_shop_item i"
	sqlStr = sqlStr + " 			on d.itemgubun=i.itemgubun"
	sqlStr = sqlStr + " 			and d.itemid=i.shopitemid"
	sqlStr = sqlStr + " 			and d.itemoption=i.itemoption"

	sqlStr = sqlStr + " 	where d.topmasteridx=" + CStr(iid)

	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail.idx=T.idx"

	rsget.Open sqlStr, dbget, 1

	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(iid)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(iid)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(iid)

	rsget.Open sqlStr, dbget, 1

elseif mode="addmaster" then
	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("shopid") = request("shopid")
	rsget("title") = html2db(request("title"))
	rsget("totalsum") =  request("totalsuplycash") ''request("totalsum") =>발행금액 == 공급액
	rsget("totalsellcash") = request("totalsuplycash")
	rsget("totalbuycash") = request("totalbuycash")
	rsget("totalsuplycash") = request("totalsuplycash")

	rsget("divcode") = request("divcode")
	rsget("etcstr") = html2db(request("etcstr"))
	rsget("reguserid") = session("ssBctId")
	rsget("regusername") = session("ssBctCname")

    rsget("shopdiv") = shopdiv
    rsget("diffKey") = diffKey

'	if request("taxdate")<>"" then
'		rsget("taxdate") = request("taxdate")
'	end if

'	if request("ipkumdate")<>"" then
'		rsget("ipkumdate") = request("ipkumdate")
'	end if

	rsget.update
	iid = rsget("idx")
	rsget.close

elseif mode="modimaster" then
    'response.write "사용중지메뉴"
    'response.end

	'재사용 : 제목/기타메모/입금일만 수정한다.(skyer9)
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master" + VbCrlf
	sqlStr = sqlStr + " set title='" + html2db(request("title")) + "'" + VbCrlf
	sqlStr = sqlStr + " ,yyyymm='" + yyyy + "-" + mm + "'"
	sqlStr = sqlStr + " ,shopdiv='" + shopdiv + "'"
    sqlStr = sqlStr + " ,diffKey=" & diffKey
	if request("taxdate")<>"" then
		'sqlStr = sqlStr + " ,taxdate='" + request("taxdate") + "'"
	else
		'sqlStr = sqlStr + " ,taxdate=NULL"
	end if

	if request("ipkumdate")<>"" then
		sqlStr = sqlStr + " ,ipkumdate='" + request("ipkumdate") + "'"
	else
		'sqlStr = sqlStr + " ,ipkumdate=NULL"
	end if

	sqlStr = sqlStr + " ,etcstr='" + html2db(request("etcstr")) + "'"
	sqlStr = sqlStr + " ,finishuserid='" + session("ssBctId") + "'"
	''session("ssBctCname") 값이 없으면 에러가 발생한다.
	sqlStr = sqlStr + " ,finishusername='" + session("ssBctCname") + "'"
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1

elseif mode="changeState" then

	If request("statecd") = "0" Then
		'======================================================================
		'이전 방식(tbl_tax_history_master 를 이용하는 경우)
		sqlStr = " SELECT Count(*) FROM [db_jungsan].[dbo].tbl_tax_history_master " + VbCrlf
		sqlStr = sqlStr + " WHERE jungsanGubun = 'OFFSHOP' "
		sqlStr = sqlStr + " AND jungsanID = '" + request("idx") + "'"
		sqlStr = sqlStr + " AND deleteYN = 'N' "

		rsget.Open sqlStr, dbget, 1
		If Not rsget.EOF Then
			If rsget(0) > 0 Then
				response.write "<script>" & vbCrLf
				response.write "alert('세금계산서발행로그가 존재합니다.\n\n수정하시려면 로그를 먼저 삭제하십시오.')" & vbCrLf
				response.write "history.back();" & vbCrLf
				response.write "</script>" & vbCrLf
				rsget.close
				dbget.close
				response.End
			End If
		End If
		rsget.close


		'======================================================================
		'신규 방식(tbl_taxSheet 이용)
		set oTax = new CTax

		oTax.FRectsearchKey = " t1.orderidx "
		oTax.FRectsearchString = CStr(request("idx"))
        oTax.FRectDelYn="N"
		oTax.GetTaxList

		if oTax.FResultCount > 0 then
			if oTax.FTaxList(0).FisueYn="Y" then
				response.write "<script>alert('이미 발행된 세금계산서가 있습니다.\n\n변경 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다');history.back();</script>"
			else
				response.write "<script>alert('발행대기중인 세금계산서가 있습니다.\n\n변경 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다.');history.back();</script>"
			end if

			response.End
		end if

	End If

	' 정산테이블 상태변경
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master SET " + VbCrlf
	sqlStr = sqlStr + "  statecd='" + request("statecd") + "'"
	sqlStr = sqlStr + " ,etcstr='" + html2db(request("etcstr")) + "'"
	sqlStr = sqlStr + " ,finishuserid='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " ,finishusername='" + session("ssBctCname") + "'"
	' 입금완료시 입금일은 업데이트 하나 타이틀은 수정할 수 없다.
	If request("statecd") = "7" And Len(request("ipkumdate")) = 10 Then
		sqlStr = sqlStr + " ,ipkumDate='" + request("ipkumdate") + "'"	' 입금일
	Else
		sqlStr = sqlStr + " ,title='" + html2db(request("title")) + "'"
	End If
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	dbget.Execute(sqlStr)

	If request("statecd") = "7" And Len(request("ipkumdate")) = 10 Then

		' 주문서 테이블 입금일자 업데이트
		sqlStr = " UPDATE A " + vbCrlf
		sqlStr = sqlStr + " SET a.ipkumDate = '" + request("ipkumdate") + "'"  + vbCrlf
		sqlStr = sqlStr + " FROM db_storage.dbo.tbl_ordersheet_master a " & vbCrLf
		sqlStr = sqlStr + " INNER JOIN [db_shop].[dbo].tbl_fran_meachuljungsan_submaster b " & vbCrLf
		sqlStr = sqlStr + " ON a.baljucode = b.code02 " & vbCrLf
		sqlStr = sqlStr + " WHERE b.masterIdx = " + CStr(idx)

		dbget.Execute(sqlStr)
	End If

elseif mode="delmaster" then
	''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script>alert('수정중 상태에서만 삭제 가능합니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
	sqlStr = sqlStr + " where topmasteridx=" + CStr(idx)

	rsget.Open sqlStr, dbget, 1


	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster" + VbCrlf
	sqlStr = sqlStr + " where masteridx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1

	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_master" + VbCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)  + VbCrlf

	rsget.Open sqlStr, dbget, 1
elseif mode="modidetail" then
    ''현재상태 0 수정중인 경우만 수정 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script>alert('수정중 상태에서만 수정 가능합니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if


	dim ckidx, suplycasharr,itemnoarr
	dim orgsellcasharr,sellcasharr,buycasharr

	ckidx = request.form("ckidx") + ","
	itemnoarr= request.form("itemnoarr")
	suplycasharr = request.form("suplycasharr")
	orgsellcasharr = request.form("orgsellcasharr")
	sellcasharr = request.form("sellcasharr")
	buycasharr = request.form("buycasharr")

	ckidx = split(ckidx,",")
	suplycasharr = split(suplycasharr,",")
	orgsellcasharr = split(orgsellcasharr,",")
	sellcasharr = split(sellcasharr,",")
	buycasharr = split(buycasharr,",")
	itemnoarr = split(itemnoarr,",")

	for i=0 to Ubound(ckidx)
		if trim(ckidx(i))<>"" then
			sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
			sqlStr = sqlStr + " set orgsellcash=" + CStr(orgsellcasharr(i))  + VbCrlf
			sqlStr = sqlStr + " , sellcash=" + CStr(sellcasharr(i))  + VbCrlf
			sqlStr = sqlStr + " , buycash=" + CStr(buycasharr(i))  + VbCrlf
			sqlStr = sqlStr + " , suplycash=" + CStr(suplycasharr(i))  + VbCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(itemnoarr(i))  + VbCrlf
			sqlStr = sqlStr + " where idx=" + trim(ckidx(i))

			rsget.Open sqlStr, dbget, 1
		end if
	next


	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=0"
	sqlStr = sqlStr + " ,totalbuycash=0"
	sqlStr = sqlStr + " ,totalsuplycash=0"
	sqlStr = sqlStr + " ,totalorgsellcash=0"
	sqlStr = sqlStr + " where masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totalbuycash,0)"
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " ,totalorgsellcash=IsNULL(T.totalorgsellcash,0)"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=IsNULL(T.totalsum,0)"
	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totalbuycash,0)"
	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " ,totalorgsellcash=IsNULL(T.totalorgsellcash,0)"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1
elseif mode="deldetail" then
    ''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script>alert('수정중 상태에서만 삭제 가능합니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if


	ckidx = trim(request.form("ckidx") + ",")

	if Right(ckidx,1)="," then
		ckidx = Left(ckidx,Len(ckidx)-1)
	end if

	''response.write ckidx


	sqlStr = " delete from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail" + VbCrlf
	sqlStr = sqlStr + " where idx in (" + trim(ckidx) + ")"
	''response.write sqlStr
	''dbget.close()	:	response.End
	rsget.Open sqlStr, dbget, 1


	'' update Sub master ''전체 삭제시 고려.
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=0"
	sqlStr = sqlStr + " ,totalbuycash=0"
	sqlStr = sqlStr + " ,totalsuplycash=0"
	sqlStr = sqlStr + " ,totalorgsellcash=0"
	sqlStr = sqlStr + " where masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1
elseif mode="etcsubadd" then
    cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script>alert('수정중 상태에서만 추가 가능합니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = topidx
	rsget("linkidx") = 0
	rsget("shopid") = shopid
	rsget("code01") = yyyy + "-" + mm
	rsget("code02") = makerid
	rsget("execdate") = yyyy + "-" + mm + "-01"
	rsget("totalcount") = 0
	rsget("totalsellcash") = 0
	rsget("totalbuycash") = 0
	rsget("totalsuplycash") = 0
	rsget("totalorgsellcash") = 0
	rsget.update
	iid = rsget("idx")
	rsget.close
elseif mode="etcsubdetailadd" then
    ''현재상태 0 수정중인 경우만 삭제 가능
	cnt=0
	sqlStr = " select count(idx) as cnt from [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " where idx=" + CStr(topidx)  + VbCrlf
    sqlStr = sqlStr + " and statecd=0"

	rsget.Open sqlStr, dbget, 1
		cnt = rsget("cnt")
	rsget.Close

	if cnt<1 then
		response.write "<script>alert('수정중 상태에서만 추가 가능합니다.');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	end if


	sqlStr = " select * from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew

	rsget("masteridx") = idx
	rsget("topmasteridx") = topidx
	rsget("linkbaljucode") = request("linkbaljucode")
	rsget("linkmastercode") = "0"
	rsget("linkdetailidx") = 0
	rsget("itemgubun") = request("itemgubun")
	rsget("itemid") = request("itemid")
	rsget("itemoption") = request("itemoption")
	rsget("itemname") = html2Db(request("itemname"))
	rsget("itemoptionname") = html2Db(request("itemoptionname"))
	rsget("makerid") = request("makerid")
	rsget("itemno") = request("itemno")
	rsget("sellcash") = request("sellcash")
	rsget("suplycash") = request("suplycash")
	rsget("buycash") = request("buycash")
	rsget("orgsellcash") = request("orgsellcash")

	rsget.update
	iid = rsget("idx")
	rsget.close


	'' update Sub master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " set totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select masteridx, sum(sellcash*itemno) as totalsellcash, "
	sqlStr = sqlStr + "			sum(buycash*itemno) as totalbuycash, sum(suplycash*itemno) as totalsuplycash,"
	sqlStr = sqlStr + "			sum(orgsellcash*itemno) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_subdetail"
	sqlStr = sqlStr + " 		where topmasteridx=" + CStr(topidx)
	sqlStr = sqlStr + "		    group by masteridx"
	sqlStr = sqlStr + " ) as T "
	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.idx=T.masteridx"
	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_fran_meachuljungsan_submaster.masteridx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1

	'' update master
	sqlStr = " update [db_shop].[dbo].tbl_fran_meachuljungsan_master"
	sqlStr = sqlStr + " set totalsum=T.totalsum"
	sqlStr = sqlStr + " ,totalsellcash=T.totalsellcash"
	sqlStr = sqlStr + " ,totalbuycash=T.totalbuycash"
	sqlStr = sqlStr + " ,totalsuplycash=T.totalsuplycash"
	sqlStr = sqlStr + " ,totalorgsellcash=T.totalorgsellcash"
	sqlStr = sqlStr + " from (select sum(totalsuplycash) as totalsum, sum(totalsellcash) as totalsellcash, "
	sqlStr = sqlStr + "			sum(totalbuycash) as totalbuycash, sum(totalsuplycash) as totalsuplycash, "
	sqlStr = sqlStr + "			sum(totalorgsellcash) as totalorgsellcash "
	sqlStr = sqlStr + " 		from [db_shop].[dbo].tbl_fran_meachuljungsan_submaster"
	sqlStr = sqlStr + " 		where masteridx=" + CStr(topidx)
	sqlStr = sqlStr + " 	) as T "
	sqlStr = sqlStr + " where idx=" + CStr(topidx)

	rsget.Open sqlStr, dbget, 1
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
<% if (mode="chulgo") or (mode="modimaster") or (mode="witsksell")  then %>
alert('저장 되었습니다.');
opener.popMasterEdit('<%= iid %>');
opener.location.reload();
window.close();
<% elseif mode="changeState" then %>
alert('변경 되었습니다.');
opener.location.reload();
window.close();
<% elseif mode="delmaster" then %>
alert('삭제 되었습니다.');
opener.location.reload();
window.close();
<% elseif (mode="etcsubadd") or (mode="etcsubdetailadd") then %>
alert('저장 되었습니다.');
opener.location.reload();
window.close();
<% elseif mode="modidetail" or mode="deldetail" then %>
alert('수정 되었습니다.');
location.replace('<%= refer %>');
<% else %>
alert('저장 되었습니다.');
opener.popMasterEdit('<%= iid %>');
opener.location.reload();
window.close();
<% end if %>
</script>


<!-- #include virtual="/lib/db/dbclose.asp" -->