<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim idx, mode, gubun, designer, yyyy1, mm1, yyyymm, tx_memo, rd_state, jgubun
dim itemno,sellcash,suplycash, reducedprice, commission, itemvatyn
dim itemname

dim itemid, itemoption, prejaego
dim ipgono, chulgono, sellno, ocha, realjaego
dim jungsanno, isdelete
dim detailidx
dim midx, code
dim taxregdate,ipkumregdate

dim preidx,curridx
dim premastercode,currmastercode

dim taxtype, differencekey
dim groupid, availneoport
dim neotaxno, taxlinkidx, billsiteCode, eseroEvalSeq
Dim AssignedRow
dim iCheExists,ipfileNo
dim jacctcd
dim targetGbn, idxarr

midx    = request("midx")
idx     = request.form("idx")
idxarr     = request.form("idxarr")
mode    = request("mode")
gubun   = request("gubun")
designer = request("designer")
yyyy1   = request("yyyy1")
mm1     = request("mm1")
tx_memo = html2db(request("tx_memo"))
rd_state = request("rd_state")

itemno          = replace(request("itemno"),",","")
sellcash        = replace(request("sellcash"),",","")
suplycash       = replace(request("suplycash"),",","")
reducedprice    = replace(request("reducedprice"),",","")
commission      = replace(request("commission"),",","")
itemvatyn       = requestCheckvar(request("itemvatyn"),10)

itemname = html2db(request("itemname"))

yyyymm = yyyy1 + "-" + mm1

itemid = request("itemid")
itemoption = request("itemoption")
prejaego = request("prejaego")
ipgono = request("ipgono")
chulgono = request("chulgono")
sellno = request("sellno")
ocha        = request("ocha")
realjaego   = request("realjaego")
jungsanno   = request("jungsanno")
isdelete    = request("isdelete")
detailidx   = request("detailidx")


taxregdate = request("taxregdate")
ipkumregdate = request("ipkumregdate")


preidx  = request("preidx")
curridx = request("curridx")
premastercode = request("premastercode")
currmastercode = request("currmastercode")

taxtype = request("taxtype")
differencekey = request("differencekey")
groupid = request("groupid")
availneoport = request("availneoport")
neotaxno = RequestCheckvar(request("neotaxno"),32)
taxlinkidx = RequestCheckvar(request("taxlinkidx"),10)

billsiteCode= RequestCheckvar(request("billsiteCode"),10)
eseroEvalSeq= RequestCheckvar(request("eseroEvalSeq"),24)

jacctcd = Trim(requestCheckVar(request("jacctcd"),10))

if (availneoport="on") then
    availneoport="1"
else
    availneoport="0"
end if

''//taxtype 01:세금계산서 02:계산서 03:영세율
'if taxtype="" then taxtype="01"
if differencekey="" then differencekey="0"


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr
dim masterExists
dim notstatemodi
dim masteridx

dim i,cnt

dim bufsellcash,bufsuplycash, bufprejaego, bufipgono, bufchulgono
dim bufsellno, bufocha, bufrealjaego, bufjungsanno, bufdetailidx
dim AssignRow, mastercode

if mode="arrsave" then
	if gubun="upche" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write notstatemodi
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.idx, d.orderserial,"
			sqlStr = sqlStr + " m.buyname, m.reqname, d.itemid, d.itemoption,"
			sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
			sqlStr = sqlStr + " d.itemcost, d.buycash"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_master m on m.orderserial=d.orderserial"
			sqlStr = sqlStr + " where d.idx in (" + idx + ")"
			sqlStr = sqlStr + " and d.idx not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "')"
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set ub_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,ub_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,ub_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		rsget.Open sqlStr,dbget,1

        ''groupid 추가.
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlStr = sqlStr + " set groupid=p.groupid"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
        sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
        rsget.Open sqlStr,dbget,1

	elseif gubun="lecture" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write notstatemodi
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", 'upche', d.detailidx, d.orderserial,"
			sqlStr = sqlStr + " m.buyname, d.entryname, d.itemid, d.itemoption,"
			sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
			sqlStr = sqlStr + " d.itemcost, d.buycash"
			sqlStr = sqlStr + " from [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_detail d"
			sqlStr = sqlStr + " left join [ACADEMYDB].[db_academy].[dbo].tbl_academy_order_master m on m.orderserial=d.orderserial"
			sqlStr = sqlStr + " where d.detailidx in (" + idx + ")"
			'sqlStr = sqlStr + " and d.detailidx not in ( "
			'sqlStr = sqlStr + " 	select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			'sqlStr = sqlStr + " where gubuncd='upche')"

			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set ub_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,ub_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,ub_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and gubuncd='upche') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		rsget.Open sqlStr,dbget,1


        ''groupid 추가.
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlStr = sqlStr + " set groupid=p.groupid"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
        sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
        rsget.Open sqlStr,dbget,1

	elseif gubun="witaksell" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then


			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.idx, d.orderserial,"
			sqlStr = sqlStr + " m.buyname, m.reqname, d.itemid, d.itemoption,"
			sqlStr = sqlStr + " d.itemname, d.itemoptionname, d.itemno,"
			sqlStr = sqlStr + " d.itemcost, d.buycash"
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_master m on m.orderserial=d.orderserial"

			'sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_detail_2003 d"
			'sqlStr = sqlStr + " left join [db_log].[dbo].tbl_old_order_master_2003 m on m.orderserial=d.orderserial"

			sqlStr = sqlStr + " where d.idx in (" + idx + ")"
			sqlStr = sqlStr + " and d.idx not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "' )"
			'response.write sqlStr
			'dbget.close()	:	response.End
			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1


        ''groupid 추가.
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlStr = sqlStr + " set groupid=p.groupid"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
        sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
        rsget.Open sqlStr,dbget,1


	elseif gubun="maeip" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.id, d.mastercode,"
			sqlStr = sqlStr + " d.itemid, d.itemoption,"
			sqlStr = sqlStr + " i.itemname, d.iitemoptionname as itemoptionname,"
			sqlStr = sqlStr + " d.itemno,"
			sqlStr = sqlStr + " d.sellcash, d.suplycash"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_storage].[dbo].tbl_acount_storage_detail d"
			'sqlStr = sqlStr + " left join [db_item].[dbo].vw_all_option v on d.itemoption=v.optioncode"
			sqlStr = sqlStr + " where d.id in (" + idx + ")"
			sqlStr = sqlStr + " and i.itemid=d.itemid"
			sqlStr = sqlStr + " and d.id not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "')"
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set me_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,me_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,me_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1


        ''groupid 추가.
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlStr = sqlStr + " set groupid=p.groupid"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
        sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
        rsget.Open sqlStr,dbget,1

	elseif gubun="maeipchulgo" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,reqname,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.id, d.mastercode,m.socid,"
			sqlStr = sqlStr + " d.itemid, d.itemoption,"
			sqlStr = sqlStr + " i.itemname, d.iitemoptionname as itemoptionname,"
			sqlStr = sqlStr + " d.itemno,"
			sqlStr = sqlStr + " d.sellcash, d.buycash"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
			'sqlStr = sqlStr + " left join [db_item].[dbo].vw_all_option v on d.itemoption=v.optioncode"
			sqlStr = sqlStr + " where d.id in (" + idx + ")"
			sqlStr = sqlStr + " and m.code=d.mastercode"
			sqlStr = sqlStr + " and i.itemid=d.itemid"
			sqlStr = sqlStr + " and d.id not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "')"
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if
	elseif gubun="witak" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.id, d.mastercode,"
			sqlStr = sqlStr + " d.itemid, d.itemoption,"
			sqlStr = sqlStr + " i.itemname, d.iitemoptionname as itemoptionname,"
			sqlStr = sqlStr + " d.itemno,"
			sqlStr = sqlStr + " d.sellcash, d.suplycash"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_storage].[dbo].tbl_acount_storage_detail d"
			'sqlStr = sqlStr + " left join [db_item].[dbo].vw_all_option v on d.itemoption=v.optioncode"
			sqlStr = sqlStr + " where d.id in (" + idx + ")"
			sqlStr = sqlStr + " and i.itemid=d.itemid"
			sqlStr = sqlStr + " and d.id not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "')"
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		'sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		'sqlStr = sqlStr + " set wi_cnt=T.cnt"
		'sqlStr = sqlStr + " ,wi_totalsellcash=T.totalsellcash"
		'sqlStr = sqlStr + " ,wi_totalsuplycash=T.totalsuplycash"
		'sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		'sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		'sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + ") as T"
		'sqlStr = sqlStr + " where id=" + CStr(masteridx)
		'response.write sqlStr
		'rsget.Open sqlStr,dbget,1
	elseif gubun="witakchulgo" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where designerid='" + designer + "'"
		sqlStr = sqlStr + " and yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + " and differencekey=" + CStr(differencekey)
		sqlStr = sqlStr + " and taxtype='" + taxtype + "'"

		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if Not masterExists then
			''Insert Master
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " (designerid,yyyymm,title,taxtype,differencekey)"
			sqlStr = sqlStr + " values('" + designer + "'"
			sqlStr = sqlStr + " ,'" + yyyymm + "'"
			if (Cstr(differencekey)<>"0") then
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산(" + CStr(differencekey) + ")'"
			else
				sqlStr = sqlStr + " ,'" + yyyy1 + "년 " + mm1 + "월 정산'"
			end if
			sqlStr = sqlStr + " ,'" + taxtype + "'"
			sqlStr = sqlStr + " ,'" + CStr(differencekey) + "'"
			sqlStr = sqlStr + " )"

			rsget.Open sqlStr,dbget,1

			sqlStr = "select IDENT_CURRENT('[db_jungsan].[dbo].tbl_designer_jungsan_master') as id"
			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				masteridx = rsget("id")
			end if
			rsget.Close
		end if

		if Right(idx,1)="," then
			idx = Left(idx,Len(idx)-1)
		end if

		''Insert Detail
		if idx<>"" then
			sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,reqname,"
			sqlStr = sqlStr + " itemid,itemoption,itemname,itemoptionname,itemno,"
			sqlStr = sqlStr + " sellcash,suplycash)"
			sqlStr = sqlStr + " select " + CStr(masteridx) + ", '" + gubun + "', d.id, d.mastercode, m.socid,"
			sqlStr = sqlStr + " d.itemid, d.itemoption,"
			sqlStr = sqlStr + " i.itemname, d.iitemoptionname as itemoptionname,"
			sqlStr = sqlStr + " d.itemno*-1,"
			sqlStr = sqlStr + " d.sellcash, d.buycash"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i, [db_storage].[dbo].tbl_acount_storage_master m, [db_storage].[dbo].tbl_acount_storage_detail d"
			'sqlStr = sqlStr + " left join [db_item].[dbo].vw_all_option v on d.itemoption=v.optioncode"
			sqlStr = sqlStr + " where d.id in (" + idx + ")"
			sqlStr = sqlStr + " and m.code=d.mastercode"
			sqlStr = sqlStr + " and i.itemid=d.itemid"
			sqlStr = sqlStr + " and d.id not in ( select detailidx from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
			sqlStr = sqlStr + " where gubuncd='" + gubun + "')"
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1


        ''groupid 추가.
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
        sqlStr = sqlStr + " set groupid=p.groupid"
        sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
        sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
        rsget.Open sqlStr,dbget,1

	elseif gubun="witakjungsan" then
		sqlStr = "select top 1 id, finishflag,yyyymm from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where id=" + idx + ""
		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
			yyyymm = rsget("yyyymm")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		if len(itemid)>1 then
			detailidx = split(detailidx,"|")
			itemid = split(itemid,"|")
			itemoption = split(itemoption,"|")
			sellcash = split(sellcash,"|")
			suplycash = split(suplycash,"|")
			prejaego = split(prejaego,"|")
			ipgono   = split(ipgono,"|")
			chulgono = split(chulgono,"|")
			sellno	 = split(sellno,"|")
			ocha	= split(ocha,"|")
			realjaego = split(realjaego,"|")
			jungsanno = split(jungsanno,"|")
			isdelete  = split(isdelete,"|")

			cnt = UBound(itemid)
			for i=0 to cnt-1
				if Not IsNumeric(sellcash(i)) then
					bufsellcash ="0"
				else
					bufsellcash = sellcash(i)
				end if
				if Not IsNumeric(suplycash(i)) then
					bufsuplycash ="0"
				else
					bufsuplycash = suplycash(i)
				end if
				if Not IsNumeric(prejaego(i)) then
					bufprejaego ="0"
				else
					bufprejaego =prejaego(i)
				end if
				if Not IsNumeric(ipgono(i)) then
					bufipgono ="0"
				else
					bufipgono =ipgono(i)
				end if
				if Not IsNumeric(chulgono(i)) then
					bufchulgono ="0"
				else
					bufchulgono =chulgono(i)
				end if
				if Not IsNumeric(sellno(i)) then
					bufsellno ="0"
				else
					bufsellno =sellno(i)
				end if
				if Not IsNumeric(ocha(i)) then
					bufocha ="0"
				else
					bufocha =ocha(i)
				end if
				if Not IsNumeric(realjaego(i)) then
					bufrealjaego ="0"
				else
					bufrealjaego =realjaego(i)
				end if
				if Not IsNumeric(jungsanno(i)) then
					bufjungsanno ="0"
				else
					bufjungsanno =jungsanno(i)
				end if

				if (detailidx(i)<>"") then
					sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_witak"
					sqlStr = sqlStr + " set sellcash=" + bufsellcash + ","
					sqlStr = sqlStr + " suplycash=" + bufsuplycash + ","
					sqlStr = sqlStr + " prejaego=" + bufprejaego + ","
					sqlStr = sqlStr + " ipgono=" + bufipgono + ","
					sqlStr = sqlStr + " chulgono=" + bufchulgono + ","
					sqlStr = sqlStr + " sellno=" + bufsellno + ","
					sqlStr = sqlStr + " ocha=" + bufocha + ","
					sqlStr = sqlStr + " realjaego=" + bufrealjaego + ","
					sqlStr = sqlStr + " jungsanno=" + bufjungsanno + ","
					sqlStr = sqlStr + " deleteyn='" + isdelete(i) + "'"
					sqlStr = sqlStr + " where id=" + detailidx(i)

					rsget.Open sqlStr,dbget,1
				else
					sqlStr = " insert into [db_jungsan].[dbo].tbl_designer_jungsan_witak"
					sqlStr = sqlStr + " (masterid,itemid,itemoption,sellcash,suplycash,prejaego,"
					sqlStr = sqlStr + " ipgono, chulgono, sellno, ocha, realjaego, jungsanno, deleteyn)"
					sqlStr = sqlStr + " values(" + idx + ","
					sqlStr = sqlStr + " " + itemid(i) + ","
					sqlStr = sqlStr + " '" + itemoption(i) + "',"
					sqlStr = sqlStr + " " + bufsellcash + ","
					sqlStr = sqlStr + " " + bufsuplycash + ","
					sqlStr = sqlStr + " " + bufprejaego + ","
					sqlStr = sqlStr + " " + bufipgono + ","
					sqlStr = sqlStr + " " + bufchulgono + ","
					sqlStr = sqlStr + " " + bufsellno + ","
					sqlStr = sqlStr + " " + bufocha + ","
					sqlStr = sqlStr + " " + bufrealjaego + ","
					sqlStr = sqlStr + " " + bufjungsanno + ","
					sqlStr = sqlStr + " '" + isdelete(i) + "')"

					rsget.Open sqlStr,dbget,1
				end if

			next

			''Update Master
			sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
			sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
			sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
			sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
			sqlStr = sqlStr + " from (select count(d.id) as cnt, sum(d.jungsanno*d.sellcash) as totalsellcash,sum(d.jungsanno*d.suplycash) as totalsuplycash"
			sqlStr = sqlStr + " from  [db_jungsan].[dbo].tbl_designer_jungsan_witak d"
			sqlStr = sqlStr + " where masterid=" + CStr(masteridx)
			sqlStr = sqlStr + " and deleteyn='N') as T"
			sqlStr = sqlStr + " where id=" + CStr(masteridx)
			'response.write sqlStr
			rsget.Open sqlStr,dbget,1

            ''groupid 추가.
            sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
            sqlStr = sqlStr + " set groupid=p.groupid"
            sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p"
            sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=" + CStr(masteridx)
            sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_designer_jungsan_master.designerid=p.id"
            rsget.Open sqlStr,dbget,1

			''update Storage
			'yyyymm
			sqlStr = " select code from [db_storage].[dbo].tbl_acount_storage_master"
			sqlStr = sqlStr + " where Left(code,2)='ME'"
			sqlStr = sqlStr + " and deldt is null"
			sqlStr = sqlStr + " and convert(varchar(7),executedt,21)='" + yyyymm + "'"

			rsget.Open sqlStr,dbget,1
				code = rsget("code")
			rsget.close

			if IsNull(code) or (code="") then
				response.write "<script language=javascript>"
				response.write "alert(" + yyyymm + "'월 월말재고가 입력되지 않았습니다. \r\n관리자에게 문의 하세요.');"
				response.write "location.replace('" + refer + "');"
				response.write "</script>"
				dbget.close()	:	response.End
			end if

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
			sqlStr = sqlStr + " set itemno=T.realjaego,"
			sqlStr = sqlStr + " updt=getdate()"
			sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " (select w.itemid, w.itemoption, sum(w.realjaego) as realjaego "
				sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_witak w"
				sqlStr = sqlStr + " where w.masterid=" + CStr(idx)
				sqlStr = sqlStr + " and w.deleteyn='N'"
				sqlStr = sqlStr + " group by w.itemid, w.itemoption"
				sqlStr = sqlStr + " ) as T"
			sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
			sqlStr = sqlStr + " and deldt is null"
			sqlStr = sqlStr + " and T.itemid=[db_storage].[dbo].tbl_acount_storage_detail.itemid"
			sqlStr = sqlStr + " and T.itemoption=[db_storage].[dbo].tbl_acount_storage_detail.itemoption"

			'response.write sqlStr
			rsget.Open sqlStr,dbget,1
		end if
	elseif gubun="witakjungsan_del" then
		sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " where id=" + idx + ""
		rsget.Open sqlStr,dbget,1
		masterExists = Not rsget.Eof
		if masterExists then
			masteridx = rsget("id")
			notstatemodi = Not (rsget("finishflag")="0")
		end if
		rsget.Close

		if notstatemodi then
			response.write "<script language=javascript>"
			response.write "alert('현재 수정중 상태가 아닙니다.');"
			response.write "location.replace('" + refer + "');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if

		sqlStr = "delete from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
		sqlStr = sqlStr + " where masterid=" + CStr(masteridx)

		rsget.Open sqlStr,dbget,1

		''Update Master
		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
		sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
		sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
		sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
		sqlStr = sqlStr + " from (select count(d.id) as cnt, sum(d.jungsanno*d.sellcash) as totalsellcash,sum(d.jungsanno*d.suplycash) as totalsuplycash"
		sqlStr = sqlStr + " from  [db_jungsan].[dbo].tbl_designer_jungsan_witak d"
		sqlStr = sqlStr + " where masterid=" + CStr(masteridx)
		sqlStr = sqlStr + " and deleteyn='N') as T"
		sqlStr = sqlStr + " where id=" + CStr(masteridx)
		'response.write sqlStr
		rsget.Open sqlStr,dbget,1
	end if
elseif mode="dellall" then
	sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + idx + ""
	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
	end if
	rsget.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('현재 수정중 상태가 아닙니다.');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

    ''입출금File에 내역이 있으면 삭제 불가 //2012/12/12
    iCheExists = FALSE
    sqlstr = " select ipfileNo from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
    sqlstr = sqlstr + "     where targetGbn='ON'"
    sqlstr = sqlstr + "     and targetIdx=" + CStr(idx) + VbCrlf
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ipfileNo = rsget("ipfileNo")
        iCheExists = true
    end if
    rsget.close

    if (iCheExists) then
        response.write "<script>alert('이체 파일 내역이 존재합니다.(파일번호:"&ipfileNo&") 내역을 삭제할 수 없습니다..');</script>"
        response.write "<script>location.replace('" + refer + "');</script>"
		dbget.close()	:	response.End
    end if

	sqlStr = "delete from [db_jungsan].[dbo].tbl_designer_jungsan_witak"
	sqlStr = sqlStr + " where masterid=" + CStr(masteridx)

	rsget.Open sqlStr,dbget,1

	sqlStr = "delete from [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)

	rsget.Open sqlStr,dbget,1

	sqlStr = "delete from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + CStr(masteridx)

	rsget.Open sqlStr,dbget,1
elseif mode="memoedit" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	if gubun="maeip" then
		sqlStr = sqlStr + " set me_comment='" + tx_memo + "'"
	elseif gubun="witak" then
		sqlStr = sqlStr + " set wi_comment='" + tx_memo + "'"
	else
		sqlStr = sqlStr + " set ub_comment='" + tx_memo + "'"
	end if

	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="statechange" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set finishflag='" + rd_state + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="multistatechange" then
	idxarr = left(idxarr,len(idxarr)-1)
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set finishflag='3'"
	sqlStr = sqlStr + " where id in (" + CStr(idxarr) + ")"

	rsget.Open sqlStr,dbget,1
elseif mode="editAvailNeo" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set availneo="&availneoport
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="delTaxInfo" then
	sqlstr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set taxlinkidx=NULL"
    sqlstr = sqlstr + " ,neotaxno=NULL"
    sqlstr = sqlstr + " ,eseroevalseq=NULL"
    sqlstr = sqlstr + " ,taxregdate=NULL"
    sqlstr = sqlstr + " ,taxinputdate=NULL"
    sqlstr = sqlstr + " ,billsitecode=NULL"
    sqlstr = sqlstr + " where id=" + CStr(idx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('0','1','2')"  + VbCrlf

	rsget.Open sqlStr,dbget,1

elseif mode="editGroupid" then
    sqlStr = "select top 1 id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + idx + ""
	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
	end if
	rsget.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('현재 업체확인중 이전 상태가 아닙니다.');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set groupid='" + groupid + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif (mode="editJAcctCd") then
    if (jacctcd="") then
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	    sqlStr = sqlStr + " set jacctcd=NULL"
	    sqlStr = sqlStr + " where id=" + CStr(idx)

	    rsget.Open sqlStr,dbget,1  
    else
        sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	    sqlStr = sqlStr + " set jacctcd='" + jacctcd + "'"
	    sqlStr = sqlStr + " where id=" + CStr(idx)

	    rsget.Open sqlStr,dbget,1
    end if

elseif mode="taxregchange" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set taxregdate='" + taxregdate + "'"
	sqlStr = sqlStr + " ,taxinputdate=getdate()"
	sqlStr = sqlStr + " ,finishflag=(CASE WHEN finishflag='1' THEN '3' ELSE finishflag END)"
	IF (neotaxno<>"") or (taxlinkidx="") then
	    sqlStr = sqlStr + " ,neotaxno='"&neotaxno&"'"
	    sqlStr = sqlStr + " ,billsiteCode='"&billsiteCode&"'"+ VbCrlf
    end if
    sqlStr = sqlStr + " ,eseroEvalSeq='"&eseroEvalSeq&"'"+ VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(idx)
	rsget.Open sqlStr,dbget,1

	if (taxlinkidx="") then
	    sqlStr = " exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne] '"&eseroEvalSeq&"',1,"&idx&""
	    dbget.Execute sqlStr,AssignedRow
	    ''if (AssignedRow<1) then AssignedRow=0
	    ''response.write "<script>alert('Tax 매핑 : "&AssignedRow&" 건');</script>"
	end if
elseif mode="ipkumregchange" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ipkumdate='" + ipkumregdate + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="ipkumfinish" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set ipkumdate='" + ipkumregdate + "'"
	sqlStr = sqlStr + " , finishflag='" + rd_state + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="taxtypechange" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set taxtype='" + taxtype + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)
response.write sqlStr
	rsget.Open sqlStr,dbget,1
elseif mode="differencekeychange" then
	''기존 차수가 존재하는지 체크...


	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set differencekey='" + differencekey + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
elseif mode="modidetail" then
    sqlStr = "select top 1 id, finishflag, jgubun, targetGbn from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + midx + ""
	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
		jgubun    = rsget("jgubun")
		targetGbn = rsget("targetGbn")
	end if
	rsget.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('수정중상태 에서만 수정 가능 합니다.');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
''response.end

	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_detail" &VbCRLF
	sqlStr = sqlStr + " set itemno=" + CStr(itemno) &VbCRLF
	sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash) &VbCRLF
	sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash) &VbCRLF
	if (jgubun="CC") then
	    sqlStr = sqlStr + " ,reducedprice=" + CStr(reducedprice) &VbCRLF
	    if (gubun="upche") or (gubun="witaksell") then
	        if (targetGbn="AC") and (masteridx>=307970) and (masteridx<>312482) then
	            sqlStr = sqlStr + " ,commission=" + CStr(reducedprice-suplycash)&"-(CASE WHEN "&suplycash&"=0 THEN 0 ELSE floor(reducedprice*0.033) END)" &VbCRLF  
	            sqlStr = sqlStr + " ,pgcommission=(CASE WHEN "&suplycash&"=0 THEN 0 ELSE floor(reducedprice*0.033) END)"
	        else
	            sqlStr = sqlStr + " ,commission=" + CStr(reducedprice-suplycash) &VbCRLF  '' 자동계산되게 수정
	        end if
	    else
	        sqlStr = sqlStr + " ,commission=" + CStr(commission) &VbCRLF  '' 자동계산되게 수정
	    end if
    end if
	sqlStr = sqlStr + " where id=" + CStr(idx)

	''rsget.Open sqlStr,dbget,1
	dbget.Execute sqlStr,AssignRow

    if (AssignRow>0) then
       if (gubun="upche") or (gubun="witaksell") then
            sqlStr = " select detailidx, mastercode from [db_jungsan].[dbo].tbl_designer_jungsan_detail" &VbCRLF
            sqlStr = sqlStr + " where id=" + CStr(idx)
            rsget.Open sqlStr,dbget,1
            if Not rsget.Eof then
				detailidx = rsget("detailidx")
				mastercode = rsget("mastercode")
			end if
			rsget.Close

            if (detailidx>0) then
                sqlStr = "update db_order.dbo.tbl_order_detail"&VbCRLF
                sqlStr = sqlStr + " SET buycash="&CStr(suplycash)&VbCRLF
                sqlStr = sqlStr + " where idx="&detailidx&VbCRLF
                sqlStr = sqlStr + " and orderserial='"&mastercode&"'"&VbCRLF
                sqlStr = sqlStr + " and makerid<>'ithinkso'" ''아이띵소는 위탁
                dbget.Execute sqlStr
            end if
       end if
    end if

    sqlStr = " exec [db_jungsan].[dbo].sp_Ten_jungsanMasterSummaryUpdateON "&CStr(midx)
    dbget.Execute sqlStr

'	if gubun="upche" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set ub_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,ub_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,ub_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='upche') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="maeip" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set me_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,me_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,me_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='maeip') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witaksell" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witaksell') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakchulgo" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witakchulgo') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakoffshop" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set sh_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,sh_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,sh_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witakoffshop') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	end if
elseif mode="deldetail" then
    sqlStr = "select top 1 id, finishflag, jgubun from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + midx + ""
	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
		jgubun    = rsget("jgubun")
	end if
	rsget.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('현재 업체확인중 이전 상태가 아닙니다.');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
''response.end

	sqlStr = "delete from [db_jungsan].[dbo].tbl_designer_jungsan_detail"&VbCRLF
	sqlStr = sqlStr + " where id=" + CStr(idx) &VbCRLF

	rsget.Open sqlStr,dbget,1

    sqlStr = " exec [db_jungsan].[dbo].sp_Ten_jungsanMasterSummaryUpdateON "&CStr(midx)
    dbget.Execute sqlStr

'	if gubun="upche" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set ub_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,ub_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,ub_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='upche') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="maeip" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set me_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,me_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,me_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='maeip') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witaksell" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witaksell') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakchulgo" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witakchulgo') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakoffshop" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set sh_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,sh_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,sh_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(midx)
'		sqlStr = sqlStr + " and gubuncd='witakoffshop') as T"
'		sqlStr = sqlStr + " where id=" + CStr(midx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	end if

elseif mode="etcadd" then
    sqlStr = "select top 1 id, finishflag, jgubun from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where id=" + idx + ""
	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
		jgubun    = rsget("jgubun")
	end if
	rsget.Close

	if notstatemodi then
		response.write "<script language=javascript>"
		response.write "alert('현재 업체확인중 이전 상태가 아닙니다.\n\n!');"
		response.write "location.replace('" + refer + "');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,itemid,itemoption,itemname,"
    sqlStr = sqlStr + " itemno,sellcash,suplycash,reducedprice, commission, vatyn)"
	sqlStr = sqlStr + " values(" + CStr(idx) + ","
	sqlStr = sqlStr + " '" + gubun + "',"
	sqlStr = sqlStr + " 0,"
	sqlStr = sqlStr + " '0',"
	sqlStr = sqlStr + " 0,"
	sqlStr = sqlStr + " '0000',"
	sqlStr = sqlStr + " '" + itemname + "',"
	sqlStr = sqlStr + " " + itemno + ","
	sqlStr = sqlStr + " " + sellcash + ","
	sqlStr = sqlStr + " " + suplycash + ","
    sqlStr = sqlStr + " " + CHKIIF(reducedprice="","0",reducedprice) + ","
    sqlStr = sqlStr + " " + CHKIIF(commission="","0",commission) + ","
    sqlStr = sqlStr + " '" + itemvatyn + "'"
    sqlStr = sqlStr + ")"

	rsget.Open sqlStr,dbget,1

    sqlStr = " exec [db_jungsan].[dbo].sp_Ten_jungsanMasterSummaryUpdateON "&CStr(idx)
    dbget.Execute sqlStr

'	''Update Master
'	if gubun="upche" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set ub_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,ub_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,ub_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(idx)
'		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
'		sqlStr = sqlStr + " where id=" + CStr(idx)
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="maeip" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set me_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,me_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,me_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(idx)
'		sqlStr = sqlStr + " and gubuncd='" + gubun + "') as T"
'		sqlStr = sqlStr + " where id=" + CStr(idx)
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witaksell" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set wi_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,wi_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,wi_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(idx)
'		sqlStr = sqlStr + " and gubuncd='witaksell') as T"
'		sqlStr = sqlStr + " where id=" + CStr(idx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakchulgo" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(idx)
'		sqlStr = sqlStr + " and gubuncd='witakchulgo') as T"
'		sqlStr = sqlStr + " where id=" + CStr(idx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	elseif gubun="witakoffshop" then
'		sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
'		sqlStr = sqlStr + " set sh_cnt=IsNULL(T.cnt,0)"
'		sqlStr = sqlStr + " ,sh_totalsellcash=IsNULL(T.totalsellcash,0)"
'		sqlStr = sqlStr + " ,sh_totalsuplycash=IsNULL(T.totalsuplycash,0)"
'		sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
'		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
'		sqlStr = sqlStr + " where masteridx=" + CStr(idx)
'		sqlStr = sqlStr + " and gubuncd='witakoffshop') as T"
'		sqlStr = sqlStr + " where id=" + CStr(idx)
'		'response.write sqlStr
'		rsget.Open sqlStr,dbget,1
'	end if


	response.write "<script language=javascript>"
	response.write "alert('저장 되었습니다.');"
	response.write "opener.location.reload();"
	response.write "self.close();"
	response.write "</script>"
elseif (mode="etcbeasongpayadd") then

    if (yyyymm="") or (request("makerid")="") then
		response.write "<script language=javascript>"
		response.write "alert('올바른 접속이 아닙니다.');"
		response.write "history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

'   정산내역 등록 가능 체크


    sqlStr = "select id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and designerid='" + request("makerid") + "'"
	sqlStr = sqlStr + " and taxtype ='01'" '' 과세로 넣음. or order by

	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
	end if
	rsget.Close

	if (notstatemodi) then
		response.write "<script language=javascript>"
		response.write "alert('현재 수정중 상태가 아닙니다. \n내역을 추가 하실 수 없습니다.');"
		response.write "history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

    sqlStr = "insert into [db_jungsan].[dbo].tbl_designer_jungsan_detail"
	sqlStr = sqlStr + " (masteridx,gubuncd,detailidx,mastercode,buyname,reqname,itemid,itemoption,itemname,"
	sqlStr = sqlStr + " itemno,sellcash,suplycash)"
	sqlStr = sqlStr + " values(" + CStr(masteridx) + ","
	sqlStr = sqlStr + " '" + gubun + "',"
	sqlStr = sqlStr + " 0,"
	sqlStr = sqlStr + " '" + request("orderserial") + "',"
	sqlStr = sqlStr + " '" + request("buyname") + "',"
	sqlStr = sqlStr + " '" + request("reqname") + "',"
	sqlStr = sqlStr + " 0,"
	sqlStr = sqlStr + " '0000',"
	sqlStr = sqlStr + " '" + itemname + "',"
	sqlStr = sqlStr + " " + itemno + ","
	sqlStr = sqlStr + " " + sellcash + ","
	sqlStr = sqlStr + " " + suplycash + ")"

    'response.write sqlStr
	rsget.Open sqlStr,dbget,1


    sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
	sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and gubuncd='witakchulgo') as T"
	sqlStr = sqlStr + " where id=" + CStr(masteridx)

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1
elseif (mode="etcbeasongpayedit") then

    if (yyyymm="") or (request("makerid")="") then
		response.write "<script language=javascript>"
		response.write "alert('올바른 접속이 아닙니다.');"
		response.write "history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

'   정산내역 등록 가능 체크
    sqlStr = "select id, finishflag from [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " where yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + " and designerid='" + request("makerid") + "'"
	sqlStr = sqlStr + " and taxtype ='01'" '' 과세로 넣음. or order by

	rsget.Open sqlStr,dbget,1
	masterExists = Not rsget.Eof
	if masterExists then
		masteridx = rsget("id")
		notstatemodi = Not (rsget("finishflag")="0")
	end if
	rsget.Close

	if (notstatemodi) then
		response.write "<script language=javascript>"
		response.write "alert('현재 수정중 상태가 아닙니다. \n내역을 추가 하실 수 없습니다.');"
		response.write "history.back();"
		response.write "</script>"
		dbget.close()	:	response.End
	end if

    sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_detail" & VbCrlf
    sqlStr = sqlStr + " set buyname='" & request("buyname") & "'" & VbCrlf
    sqlStr = sqlStr + " ,reqname='" & request("reqname") & "'" & VbCrlf
    sqlStr = sqlStr + " ,itemname='" & itemname & "'" & VbCrlf
    sqlStr = sqlStr + " ,itemno=" & itemno & "" & VbCrlf
    sqlStr = sqlStr + " ,sellcash=" & sellcash & "" & VbCrlf
    sqlStr = sqlStr + " ,suplycash=" & suplycash & "" & VbCrlf
    sqlStr = sqlStr + " where id=" & request("detailid")

	rsget.Open sqlStr,dbget,1


    sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set et_cnt=IsNULL(T.cnt,0)"
	sqlStr = sqlStr + " ,et_totalsellcash=IsNULL(T.totalsellcash,0)"
	sqlStr = sqlStr + " ,et_totalsuplycash=IsNULL(T.totalsuplycash,0)"
	sqlStr = sqlStr + " from (select count(d.mastercode) as cnt, sum(d.itemno*d.sellcash) as totalsellcash,sum(d.itemno*d.suplycash) as totalsuplycash"
	sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
	sqlStr = sqlStr + " where masteridx=" + CStr(masteridx)
	sqlStr = sqlStr + " and gubuncd='witakchulgo') as T"
	sqlStr = sqlStr + " where id=" + CStr(masteridx)

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1
elseif (mode="beasongpayaddArr") then
    dim ckidx, refunddeliverypayArr
    ckidx                = Trim(request("ckidx"))
    refunddeliverypayArr = Trim(request("refunddeliverypayArr"))


    if (Right(ckidx,1)=",") then ckidx=Left(ckidx,Len(ckidx)-1)
    if (Right(refunddeliverypayArr,1)=",") then refunddeliverypayArr=Left(refunddeliverypayArr,Len(refunddeliverypayArr)-1)

    'response.write yyyymm
    'response.write ckidx
    'dbget.close()	:	response.End

''response.end

    sqlStr = " Declare @TmpListTable TABLE ( " + VbCrlf
    sqlStr = sqlStr + " [masteridx] INT , " + VbCrlf
    sqlStr = sqlStr + " [gubuncd] varchar(16), " + VbCrlf
    sqlStr = sqlStr + " [detailidx] INT, " + VbCrlf
    sqlStr = sqlStr + " [mastercode] varchar(32), " + VbCrlf
    sqlStr = sqlStr + " [buyname] varchar(32), " + VbCrlf
    sqlStr = sqlStr + " [reqname] varchar(32), " + VbCrlf
    sqlStr = sqlStr + " [itemid] INT, " + VbCrlf
    sqlStr = sqlStr + " [itemoption] varchar(4), " + VbCrlf
    sqlStr = sqlStr + " [itemname] varchar(128), " + VbCrlf
    sqlStr = sqlStr + " [itemoptionname] varchar(64), " + VbCrlf
    sqlStr = sqlStr + " [itemno] INT, " + VbCrlf
    sqlStr = sqlStr + " [sellcash] money, " + VbCrlf
    sqlStr = sqlStr + " [suplycash] money, " + VbCrlf
    sqlStr = sqlStr + " [sitename] varchar(32), " + VbCrlf
    sqlStr = sqlStr + " [beasongdate] varchar(10), " + VbCrlf
    sqlStr = sqlStr + " [vatyn] char(1)" + VbCrlf
    sqlStr = sqlStr + " )  ;" + VbCrlf

    '' 여러 차수중 1건만 등록되게 수정;; Max(J.id) => MayBe 과세 ==>과세로 변경
    sqlStr = sqlStr + " INSERT INTO @TmpListTable " + VbCrlf
    sqlStr = sqlStr + " select Max(J.id) as masteridx"
    if (request("jgubunMM")="on") then
        sqlStr = sqlStr + " , 'DL' as gubuncd " + VbCrlf
    else
        sqlStr = sqlStr + " , 'DT' as gubuncd " + VbCrlf
    end if
    sqlStr = sqlStr + " , A.id, A.orderserial, " + VbCrlf
    sqlStr = sqlStr + " m.buyname, m.reqname, 0,'0000',"
    'sqlStr = sqlStr + " Case When IsNULL(U.add_upchejungsandeliverypay,0)=0 then '반품 배송비 정산' "
    'sqlStr = sqlStr + "      When IsNULL(U.add_upchejungsandeliverypay,0)<>0 and IsNULL(R.refunddeliverypay,0)=0 then IsNULL(U.add_upchejungsancause,'') "
    'sqlStr = sqlStr + "      Else Convert(varchar(128),'반품 배송비 정산,' + IsNULL(U.add_upchejungsancause,'')) End, "
    'sqlStr = sqlStr + "  NULL, 1, 0,  IsNULL(R.refunddeliverypay,0)*-1 + IsNULL(U.add_upchejungsandeliverypay,0) " + VbCrlf
    sqlStr = sqlStr + " IsNULL(U.add_upchejungsancause,'기타'),"
    sqlStr = sqlStr + "  NULL, 1, 0, IsNULL(U.add_upchejungsandeliverypay,0)  " + VbCrlf
    sqlStr = sqlStr + "  ,m.sitename, convert(varchar(10),A.finishdate,21),'Y'"
    sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list A " + VbCrlf
   sqlStr = sqlStr + " 	left join [db_cs].[dbo].tbl_as_refund_info R  " + VbCrlf
    sqlStr = sqlStr + " 	on  A.id=R.asid " + VbCrlf
    sqlStr = sqlStr + " 	left join db_cs.dbo.tbl_as_upcheAddjungsan U" + VbCrlf
    sqlStr = sqlStr + " 	on  A.id=U.asid" + VbCrlf
    sqlStr = sqlStr + " 	left join [db_jungsan].[dbo].tbl_designer_jungsan_master J " + VbCrlf
    sqlStr = sqlStr + " 	on J.yyyymm='" + yyyymm + "' and J.designerid=A.makerid " + VbCrlf

    sqlStr = sqlStr + " 	left join [db_jungsan].[dbo].tbl_designer_jungsan_detail D " + VbCrlf
	sqlStr = sqlStr + " 	on D.itemid=0 and A.id=D.detailidx " + VbCrlf

	if (request("jgubunMM")="on") then
	    sqlStr = sqlStr + " 	and D.gubuncd='DL'  " + VbCrlf
	else
	    sqlStr = sqlStr + " 	and D.gubuncd='DT'  " + VbCrlf
    end if

    sqlStr = sqlStr + " 	left join [db_order].[dbo].tbl_order_master m " + VbCrlf
    sqlStr = sqlStr + " 	on A.orderserial=m.orderserial " + VbCrlf
    sqlStr = sqlStr + " where A.divcd in ('A004','A700','A000','A001','A100','A002','A200','A999') " + VbCrlf
    sqlStr = sqlStr + " and A.id in (" + ckidx + ") " + VbCrlf
    sqlStr = sqlStr + " and A.currstate='B007' " + VbCrlf
    sqlStr = sqlStr + " and A.deleteyn='N' " + VbCrlf
    sqlStr = sqlStr + " and ((A.divcd in ('A004','A000','A001','A100','A002','A200','A999') and A.requireupche='Y') or (A.divcd='A700'))" + VbCrlf
    sqlStr = sqlStr + " and (U.add_upchejungsandeliverypay<>0)" + VbCrlf


    ''sqlStr = sqlStr + " and A.requireupche='Y' " + VbCrlf
    ''sqlStr = sqlStr + " and R.refunddeliverypay <>0 " + VbCrlf
    ''sqlStr = sqlStr + " and ((R.refunddeliverypay <>0) or (U.add_upchejungsandeliverypay<>0))" + VbCrlf

    sqlStr = sqlStr + " and J.finishflag=0 " + VbCrlf
	IF (request("itemnotax")="on") then
    sqlStr = sqlStr + " and J.itemvatYn ='N'" + VbCrlf  ''상품이 면세인 계산서로 등록
    ELSE
    sqlStr = sqlStr + " and J.itemvatYn ='Y'" + VbCrlf  ''상품이 과세인 계산서로 등록
    END IF

    IF (request("notax")="on") then
    sqlStr = sqlStr + " and J.taxtype ='02'" + VbCrlf  ''면세로 등록
    ELSE
    sqlStr = sqlStr + " and J.taxtype ='01'" + VbCrlf  ''과세로 등록
    END IF

    if (request("jgubunMM")="on") then  '''수수료로 넣음
        sqlStr = sqlStr + " and J.jgubun='MM'"
    else
        sqlStr = sqlStr + " and j.jgubun='CC'"  + VbCrlf
    end if
    sqlStr = sqlStr + " and J.differencekey=0" '' 2011-10 추가
    sqlStr = sqlStr + " and D.id is NULL " + VbCrlf
    sqlStr = sqlStr + " group by A.id, A.orderserial,m.buyname, m.reqname,IsNULL(U.add_upchejungsancause,'기타'),IsNULL(U.add_upchejungsandeliverypay,0),m.sitename, convert(varchar(10),A.finishdate,21)" + VbCrlf

    sqlStr = sqlStr + " ;" + VbCrlf

   sqlStr = sqlStr + " insert into  [db_jungsan].[dbo].tbl_designer_jungsan_detail " + VbCrlf
    sqlStr = sqlStr + " (masteridx, " + VbCrlf
    sqlStr = sqlStr + " gubuncd, " + VbCrlf
    sqlStr = sqlStr + " detailidx, " + VbCrlf
    sqlStr = sqlStr + " mastercode, " + VbCrlf
    sqlStr = sqlStr + " buyname, " + VbCrlf
    sqlStr = sqlStr + " reqname, " + VbCrlf
    sqlStr = sqlStr + " itemid, " + VbCrlf
    sqlStr = sqlStr + " itemoption, " + VbCrlf
    sqlStr = sqlStr + " itemname, " + VbCrlf
    sqlStr = sqlStr + " itemoptionname, " + VbCrlf
    sqlStr = sqlStr + " itemno, " + VbCrlf
    sqlStr = sqlStr + " sellcash, " + VbCrlf
    sqlStr = sqlStr + " suplycash, " + VbCrlf
    sqlStr = sqlStr + " reducedprice, " + VbCrlf
    sqlStr = sqlStr + " sitename, " + VbCrlf
    sqlStr = sqlStr + " commission, " + VbCrlf
    sqlStr = sqlStr + " beasongdate, " + VbCrlf
    sqlStr = sqlStr + " vatyn, " + VbCrlf
    sqlStr = sqlStr + " CpnNotAppliedPrice " + VbCrlf
    sqlStr = sqlStr + " ) " + VbCrlf
    sqlStr = sqlStr + " select masteridx, " + VbCrlf
    sqlStr = sqlStr + " gubuncd, " + VbCrlf
    sqlStr = sqlStr + " detailidx, " + VbCrlf
    sqlStr = sqlStr + " mastercode, " + VbCrlf
    sqlStr = sqlStr + " buyname, " + VbCrlf
    sqlStr = sqlStr + " reqname, " + VbCrlf
    sqlStr = sqlStr + " itemid, " + VbCrlf
    sqlStr = sqlStr + " itemoption, " + VbCrlf
    sqlStr = sqlStr + " itemname, " + VbCrlf
    sqlStr = sqlStr + " itemoptionname, " + VbCrlf
    sqlStr = sqlStr + " itemno, " + VbCrlf
    sqlStr = sqlStr + " sellcash, " + VbCrlf
    sqlStr = sqlStr + " suplycash,  " + VbCrlf
    sqlStr = sqlStr + " sellcash as reducedprice,  " + VbCrlf
    sqlStr = sqlStr + " sitename, " + VbCrlf
    sqlStr = sqlStr + " 0 as commission,  " + VbCrlf
    sqlStr = sqlStr + " beasongdate, " + VbCrlf
    IF (request("notax")="on") then
    sqlStr = sqlStr + " 'N', " + VbCrlf
    else
    sqlStr = sqlStr + " 'Y', " + VbCrlf
    end if
    sqlStr = sqlStr + " sellcash as CpnNotAppliedPrice " + VbCrlf
    sqlStr = sqlStr + "  from @TmpListTable ;" + VbCrlf


    sqlStr = sqlStr + " update [db_jungsan].[dbo].tbl_designer_jungsan_master " + VbCrlf
    sqlStr = sqlStr + " set dlv_totalsellcash=isNULL(T.dlv_totalsellcash,0)" + VbCrlf
	sqlStr = sqlStr + "  ,dlv_totalreducedprice=isNULL(T.dlv_totalreducedprice,0)" + VbCrlf
	sqlStr = sqlStr + "  ,dlv_totalsuplycash=isNULL(T.dlv_totalsuplycash,0)" + VbCrlf
    sqlStr = sqlStr + " from ( " + VbCrlf
    sqlStr = sqlStr + "     select masteridx, sum(d.itemno*d.sellcash) as dlv_totalsellcash,sum(d.itemno*d.suplycash) as dlv_totalsuplycash,sum(d.itemno*d.reducedprice) as dlv_totalreducedprice" + VbCrlf
    sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_designer_jungsan_detail d " + VbCrlf
    sqlStr = sqlStr + "     where masteridx in ( " + VbCrlf
    sqlStr = sqlStr + "         select distinct masteridx from @TmpListTable " + VbCrlf
    sqlStr = sqlStr + "     ) " + VbCrlf
    sqlStr = sqlStr + "     and gubuncd in ('DL','DT') " + VbCrlf
    sqlStr = sqlStr + "     group by masteridx " + VbCrlf
    sqlStr = sqlStr + " ) as T " + VbCrlf
    sqlStr = sqlStr + " where [db_jungsan].[dbo].tbl_designer_jungsan_master.id=T.masteridx ;" + VbCrlf

''response.write sqlStr
    dbget.execute sqlStr

elseif (mode="monthjaegobojung") then
	'response.write "preidx" + preidx + "<br>"

	preidx     	= Left(preidx,Len(preidx))
	curridx		= Left(curridx,Len(curridx))
	itemid		= Left(itemid,Len(itemid))
	itemoption	= Left(itemoption,Len(itemoption))
	sellcash	= Left(sellcash,Len(sellcash))
	suplycash	= Left(suplycash,Len(suplycash))
	premastercode	= Left(premastercode,Len(premastercode))
	currmastercode	= Left(currmastercode,Len(currmastercode))
	prejaego = Left(prejaego,Len(prejaego))
	realjaego = Left(realjaego,Len(realjaego))

	'response.write "preidx" + preidx + "<br>"

	preidx = split(preidx,"|")
	curridx = split(curridx,"|")
	itemid = split(itemid,"|")
	itemoption = split(itemoption,"|")
	sellcash = split(sellcash,"|")
	suplycash = split(suplycash,"|")
	premastercode = split(premastercode,"|")
	currmastercode = split(currmastercode,"|")

	prejaego = split(prejaego,"|")
	realjaego  = split(realjaego,"|")
	cnt = UBound(preidx)-1



	for i=0 to cnt
		if preidx(i)<>"" then
			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
			sqlStr = sqlStr + " set itemno=" + prejaego(i)
			sqlStr = sqlStr + " where id=" + preidx(i)

			response.write sqlStr + "<br>"
			rsget.Open sqlStr,dbget,1
		else
			if premastercode(i)<>"" then
				sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail"
				sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,itemno,indt,updt)"
				sqlStr = sqlStr + " values('" + premastercode(i) + "',"
				sqlStr = sqlStr + " " + itemid(i) + ","
				sqlStr = sqlStr + " '" + itemoption(i) + "',"
				sqlStr = sqlStr + " " + sellcash(i) + ","
				sqlStr = sqlStr + " " + suplycash(i) + ","
				sqlStr = sqlStr + " " + prejaego(i) + ","
				sqlStr = sqlStr + " getdate(),"
				sqlStr = sqlStr + " getdate()"
				sqlStr = sqlStr + " )"

				response.write sqlStr + "<br>"
				rsget.Open sqlStr,dbget,1
			else
				response.write premastercode(i) + "," + itemid(i) + "," + itemoption(i)
			end if
		end if

		'response.write sqlStr + "<br>"

		if curridx(i)<>"" then
			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
			sqlStr = sqlStr + " set itemno=" + realjaego(i)
			sqlStr = sqlStr + " where id=" + curridx(i)

			'response.write sqlStr + "<br>"
			rsget.Open sqlStr,dbget,1
		else
			if currmastercode(i)<>"" then
				sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail"
				sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,itemno,indt,updt)"
				sqlStr = sqlStr + " values('" + currmastercode(i) + "',"
				sqlStr = sqlStr + " " + itemid(i) + ","
				sqlStr = sqlStr + " '" + itemoption(i) + "',"
				sqlStr = sqlStr + " " + sellcash(i) + ","
				sqlStr = sqlStr + " " + suplycash(i) + ","
				sqlStr = sqlStr + " " + realjaego(i) + ","
				sqlStr = sqlStr + " getdate(),"
				sqlStr = sqlStr + " getdate()"
				sqlStr = sqlStr + " )"

				response.write sqlStr + "<br>"
				rsget.Open sqlStr,dbget,1
			else
				response.write currmastercode(i) + "," + itemid(i) + "," + itemoption(i)
			end if
		end if

		'response.write sqlStr + "<br>"

    next
elseif (mode="brandbatchprocess") then

    jgubun = request("jgubun")
    if ((jgubun="") or (designer="") or (yyyy1="") or (mm1="") or (differencekey="") or (itemvatYN="")) then
        response.write "<script>alert('Not Valid Params key ');</script>"
        dbget.close()	:	response.End
    end if

    ''2014
    sqlStr = " exec db_jungsan.dbo.sp_Ten_jungsanMakeByBrandONN '"&jgubun&"','"&designer&"','"&yyyy1+"-"+mm1&"','"&itemvatYN&"','"&differencekey&"'"
    dbget.Execute sqlStr
    response.write "<script>alert('OK');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
else
    response.write "<script>alert('정의 되지 않았습니다. - " & mode & "');</script>"
end if


%>

<% if mode="ipkumfinish" then %>
<script language="javascript">
alert('저장 되었습니다.');
window.close();
//location.replace('/admin/upchejungsan/jungsanfinish.asp?menupos=353&ipkumregdate=<%= ipkumregdate %>');
</script>
<% else %>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->