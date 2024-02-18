<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 출고
' History : 이상구 생성
'			2019.05.23 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim masterid
dim code, mode, scheduledt, executedt, chargeid, chargename, storeid, divcode, vatcode, comment
dim itemgubunarr, itemidarr, itemoptionarr, itemnamearr, itemoptionnamearr, sellcasharr, buycasharr, suplycasharr, itemnoarr, designerarr, mwdivarr
dim itemgubun, itemid, itemoption, itemname, itemoptionname, sellcash, buycash, suplycash, itemno, designer, mwdiv
dim statecd
dim menupos
Dim finishid, finishname
dim shopid, alinkcode, currencyUnit, currencyUnit_Pos, priceChange, masteridx, sitename

masterid = request("masterid")
code = request("code")
mode = request("mode")
scheduledt = request("scheduledt")
executedt = request("executedt")
chargeid = session("ssBctid")
chargename = session("ssBctCname")
storeid = request("storeid")
socname = request("socname")
divcode = request("divcode")
vatcode = request("vatcode")
comment = html2db(request("comment"))
statecd = request("statecd")

itemgubunarr = request("itemgubunarr")
itemidarr = request("itemidarr")
itemoptionarr = request("itemoptionarr")
itemnamearr = html2db(request("itemnamearr"))
itemoptionnamearr = html2db(request("itemoptionnamearr"))
sellcasharr = request("sellcasharr")
buycasharr = request("buycasharr")
suplycasharr = request("suplycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
mwdivarr = request("mwdivarr")
menupos =  request("menupos")

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnamearr = split(itemnamearr, "|")
itemoptionnamearr = split(itemoptionnamearr, "|")
sellcasharr = split(sellcasharr, "|")
buycasharr = split(buycasharr, "|")
suplycasharr = split(suplycasharr, "|")
itemnoarr = split(itemnoarr, "|")
designerarr = split(designerarr, "|")
mwdivarr = split(mwdivarr, "|")

finishid = session("ssBctid")
finishname = html2db(session("ssBctCname"))

dim i,cnt,sqlStr, chk, didx
dim AssignedRows

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim socname, iid

dim STOCKBASEDATE, yyyymmdd
dim tmpitemid, tmpitemgubun, tmpitemoption, tmpitemno, isdeleted, ipchulflag
dim puserdiv
isdeleted = false

response.write "mode=" + mode

if mode="write" then '주문접수
	puserdiv=""
	sqlStr = " select top 1 isNULL(userdiv,'') as puserdiv" & vbcrlf	' 출고처 체크추가	' 2019.05.23 한용민 추가
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner" & vbcrlf
	sqlStr = sqlStr & " where id='" + storeid + "'" & vbcrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		puserdiv = rsget("puserdiv")
	else
		response.write "<script type='text/javascript'>alert('해당 되는 출고처가 없습니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('"& refer &"')</script>"
		rsget.close : dbget.close() : response.End
	end if
	rsget.close

	if masterid = "" or masterid ="0" then   ' 임시저장 없을 경우

			'업체명 검색
			sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
			sqlStr = sqlStr + " where userid='" + storeid + "'"
			rsget.Open sqlStr, dbget, 1
			if Not rsget.Eof then
				socname = rsget("socname_kor")
			end if
			rsget.close

				sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
				rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("code") = ""
				rsget("socid") = storeid
				rsget("socname") = socname
				rsget("chargeid") = chargeid
				rsget("chargename") = chargename
				rsget("divcode") = divcode
				rsget("vatcode") = "008"
				rsget("comment") = comment
				rsget("scheduledt") = scheduledt
				'rsget("executedt") = executedt
				rsget("statecd") = 1

				if (Left(storeid,10)="streetshop") or (Left(storeid,9)="wholesale") or (puserdiv="501") or (puserdiv="502") or (puserdiv="503") then
					rsget("ipchulflag") = "S"
				else
					rsget("ipchulflag") = "E"
				end if


				rsget.update
				iid = rsget("id")
				rsget.close
				code = "SO" + Format00(6,Right(CStr(iid),6))

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
			sqlStr = sqlStr + " set code='" + code + "'" + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(iid)
			rsget.Open sqlStr,dbget,1

			sqlStr = " delete from [db_storage].[dbo].tbl_acount_storage_detail where mastercode= '" + CStr(code) + "'"
			dbget.execute sqlStr
	ELSE
			iid = masterid
	END IF

			'''2.온라인 입고 디테일 입력
			for i=0 to UBound(itemgubunarr) - 1
				if (trim(itemgubunarr(i)) <> "") then
					itemgubun = trim(itemgubunarr(i))
					itemid = trim(itemidarr(i))
					itemoption = trim(itemoptionarr(i))
					itemname = trim(itemnamearr(i))
					itemoptionname = trim(itemoptionnamearr(i))
					sellcash = trim(sellcasharr(i))
					suplycash = trim(suplycasharr(i))
					buycash = trim(buycasharr(i))

					'출고는 수량이 줄어드는 것이다.
					itemno = -1 * CInt(trim(itemnoarr(i)))

					designer = trim(designerarr(i))
					mwdiv = trim(mwdivarr(i))
					itemname = ""
					itemoptionname = ""

					sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
					sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
					sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
					sqlStr = sqlStr + " values('" + code + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
					rsget.Open sqlStr,dbget,1
				end if
			next

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,socid='"+storeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,socname='"+socname+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargeid='"+chargeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargename='"+chargename+"'" + VBCrlf
	sqlStr = sqlStr + " ,divcode='"+divcode+"'" + VBCrlf
	sqlStr = sqlStr + " ,vatcode='008'" + VBCrlf
	sqlStr = sqlStr + " ,comment='"+comment+"'" + VBCrlf
	sqlStr = sqlStr + " ,scheduledt='"+scheduledt+"'" + VBCrlf
	sqlStr = sqlStr + " ,statecd=1" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(code) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('출고 내역이 입력되었습니다. - 출고 확정하셔야 출고처리됩니다.');</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=" + storeid + "&statecd=1&menupos="+menupos+"')</script>"
	dbget.close()	:	response.End

elseif mode ="temp" then  '임시저장

		if storeid = "" then storeid = ""
		if socname = "" then socname = ""
		if chargeid = "" then chargeid = ""
		if chargename = "" then chargename = ""
		if comment = "" then comment = ""
		if scheduledt = "" then scheduledt =  NULL

	puserdiv=""
	sqlStr = " select top 1 isNULL(userdiv,'') as puserdiv" & vbcrlf	' 출고처 체크추가	' 2019.05.23 한용민 추가
	sqlStr = sqlStr & " from [db_partner].[dbo].tbl_partner" & vbcrlf
	sqlStr = sqlStr & " where id='" + storeid + "'" & vbcrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		puserdiv = rsget("puserdiv")
	else
		response.write "<script type='text/javascript'>alert('해당 되는 출고처가 없습니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('"& refer &"')</script>"
		rsget.close : dbget.close() : response.End
	end if
	rsget.close

	if masterid = "" or masterid = "0" then
		'업체명 검색
		sqlStr = " select top 1 socname_kor from [db_user].[dbo].tbl_user_c"
		sqlStr = sqlStr + " where userid='" + storeid + "'"
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			socname = rsget("socname_kor")
		end if
		rsget.close

		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("code") = ""
		rsget("socid") = storeid
		rsget("socname") = socname
		rsget("chargeid") = chargeid
		rsget("chargename") = chargename
		rsget("divcode") = divcode
		rsget("vatcode") = "008"
		rsget("comment") = comment
		rsget("scheduledt") = scheduledt
		rsget("statecd") = 0
		'rsget("executedt") = executedt

		if (Left(storeid,10)="streetshop") or (Left(storeid,9)="wholesale") or (puserdiv="501") or (puserdiv="502") or (puserdiv="503") then
			rsget("ipchulflag") = "S"
		else
			rsget("ipchulflag") = "E"
		end if


		rsget.update
		iid = rsget("id")
		rsget.close

		code = "SO" + Format00(6,Right(CStr(iid),6))

		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
		sqlStr = sqlStr + " set code='" + code + "'" + VBCrlf
		sqlStr = sqlStr + " where id=" + CStr(iid)
		rsget.Open sqlStr,dbget,1

		sqlStr = " delete from [db_storage].[dbo].tbl_acount_storage_detail where mastercode= '" + CStr(code) + "'"
		dbget.execute sqlStr
	ELSE
			iid = masterid
	end if

	'''2.온라인 입고 디테일 입력
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemname = trim(itemnamearr(i))
			itemoptionname = trim(itemoptionnamearr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))

			'출고는 수량이 줄어드는 것이다.
			itemno = -1 * CInt(trim(itemnoarr(i)))

			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + code + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1
		end if
	next

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(code) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,socid='"+storeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,socname='"+socname+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargeid='"+chargeid+"'" + VBCrlf
	sqlStr = sqlStr + " ,chargename='"+chargename+"'" + VBCrlf
	sqlStr = sqlStr + " ,divcode='"+divcode+"'" + VBCrlf
	sqlStr = sqlStr + " ,vatcode='008'" + VBCrlf
	sqlStr = sqlStr + " ,comment='"+comment+"'" + VBCrlf
	IF not isNull(scheduledt) then
	sqlStr = sqlStr + " ,scheduledt='"+scheduledt+"'" + VBCrlf
	end if
	sqlStr = sqlStr + " ,statecd=0" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(code) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T" + vbCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('출고 내역이 임시저장되었습니다.');</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=" + storeid + "&statecd=1&menupos="+menupos+"')</script>"
	dbget.close()	:	response.End

elseif mode="delete" then
	''출고일 - 최근 2달 내역만 가능 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  m.code, m.executedt, m.ipchulflag, m.deldt, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)
	rsget.Open sqlStr, dbget, 1

	if statecd<>"0" and not(isnull(statecd)) then
		' 수정로그저장		' 2021.03.09 한용민
		chulgo_edit_log masterid,finishid,"전체삭제"
	end if

	if (not isdeleted) then
		''QuickUpdateNewIpgoDetailSummary code, true

		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
		dbget.Execute sqlStr, AssignedRows

		if (AssignedRows>0) then
		    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
		end if

	end if

	response.write "<script>alert('삭제 되었습니다.')</script>"
	response.write "<script>location.replace('culgolist.asp?research=on&code=&designer=&statecd=1')</script>"
	dbget.close()	:	response.End

elseif mode="chulgo" then

	''출고일 - 출고된건은 재출고 하지 않는다.
	sqlStr = "select top 1  m.code, m.executedt, socid as shopid "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		shopid = rsget("shopid")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set executedt='" + executedt + "' , statecd = 7, finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)

	dbget.Execute(sqlStr)

	' 수정로그저장		' 2021.03.09 한용민
	chulgo_edit_log masterid,finishid,"출고확정"

	if (yyyymmdd<>"") then
		response.write "<script>alert('Err - 이미 출고처리 된 내역입니다.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	end if

	''출고된 내역 한정판매설정

	'' item
	sqlstr = " update [db_item].[dbo].tbl_item"
	sqlstr = sqlstr + " set limitsold=limitsold - T.chulno"
	sqlstr = sqlstr + " from "
	sqlstr = sqlstr + " ("
	sqlstr = sqlstr + " 	select d.itemid, sum(d.itemno) as chulno"
	sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlstr = sqlstr + " 	where d.mastercode = '" + code + "'"
	sqlstr = sqlstr + " 	and d.deldt is NULL"
	sqlstr = sqlstr + " 	and d.itemno<0"
	sqlstr = sqlstr + " 	and d.iitemgubun='10'"
	sqlstr = sqlstr + " 	group by d.itemid"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid"
	sqlstr = sqlstr + " and [db_item].[dbo].tbl_item.limityn='Y'"

	dbget.Execute(sqlStr)

	''옵션있는상품
	sqlStr = "update [db_item].[dbo].tbl_item_option" + vbCrlf
	sqlStr = sqlStr + " set optlimitsold=optlimitsold - T.chulno" + vbCrlf
	sqlStr = sqlStr + " from " + vbCrlf
	sqlstr = sqlstr + " ("
	sqlstr = sqlstr + " 	select d.itemid, d.itemoption, sum(d.itemno) as chulno"
	sqlstr = sqlstr + " 	from [db_storage].[dbo].tbl_acount_storage_detail d"
	sqlstr = sqlstr + " 	where d.mastercode = '" + code + "'"
	sqlstr = sqlstr + " 	and d.deldt is NULL"
	sqlstr = sqlstr + " 	and d.itemno<0"
	sqlstr = sqlstr + " 	and d.iitemgubun='10'"
	sqlstr = sqlstr + " 	group by d.itemid, d.itemoption"
	sqlstr = sqlstr + " ) as T"
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=T.Itemid"
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=T.itemoption"
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.optlimityn='Y'"

	dbget.Execute(sqlStr)

	'// 평균매입가 => 출고내역매입가 '' 주석처리..2017/04/07 기타출고에만 반영해야함?
	'// 기타출고+매입재고 만 반영, 프로시저에서 처리, skyer9
	sqlStr = " exec [db_storage].[dbo].[usp_Ten_AvgIpgoPriceToAccoundStorageBuycash] '" & code & "' "
	dbget.Execute sqlStr

	''재고반영
	''QuickUpdateNewIpgoDetailSummary code,false
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	'// 매장재고 반영
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & code & "', 'N' "
	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	if (AssignedRows>0) then
	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
	end if

	response.write "<script>alert('출고처리 되었습니다..')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"
	dbget.close()	:	response.End

elseif mode="chulgo2jupsu" Then
	'// 출고완료인지 체크(샵아이디 구하기)
	sqlStr = "select top 1  m.code, m.executedt, socid as shopid "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		shopid = rsget("shopid")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""

	if (yyyymmdd = "") then
		response.write "<script>alert('Err - 출고이전 내역입니다.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	else
		'// 접수상태로 전환, 출고일자 삭제
		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
		sqlStr = sqlStr + " set executedt=NULL , statecd = 1, finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
		sqlStr = sqlStr + " where id=" + CStr(masterid)
		dbget.Execute(sqlStr)
	end if

	'// 재고반영(물류) : 출고수량 삭제
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	'// 재고반영(매장) : 출고수량 삭제
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & shopid & "', '" & code & "', 'N' "
	dbget.Execute sqlStr

	' 수정로그저장		' 2021.03.09 한용민
	chulgo_edit_log masterid,finishid,"접수전환"

	response.write "<script>alert('출고이전 전환 되었습니다..')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"
	dbget.close()	:	response.End

elseif mode="chchulgodate" Then
	''출고일 - 기출고내역 출고일자 변경
	sqlStr = "select top 1  m.code, m.executedt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" +  CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	if (yyyymmdd = "") then
		response.write "<script>alert('Err - 출고이전 내역입니다.')</script>"
		response.write "<script>location.replace('" + refer + "')</script>"
		dbget.close()	:	response.End
	end If

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set executedt='" + executedt + "', finishid = '" & finishid & "', finishname = '" & finishname & "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)
	dbget.Execute(sqlStr)

	sqlStr = " update "
	sqlStr = sqlStr + " [db_storage].[dbo].[tbl_ordersheet_master] "
	sqlStr = sqlStr + " set beasongdate = '" + executedt + "', ipgodate = '" + executedt + "', finishuser = '" & finishid & "', finishname = '" & finishname & "' "
	sqlStr = sqlStr + " where alinkcode = '" + code + "' and statecd = '7' "
	dbget.Execute(sqlStr)

	''재고반영
	''QuickUpdateNewIpgoDetailSummary code,false
    ''sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	sqlStr = " exec db_summary.dbo.[sp_Ten_recentIpChul_changeChulgoDate] '" & code & "','" & yyyymmdd & "','" & executedt & "' "
	dbget.Execute sqlStr

	' 수정로그저장		' 2021.03.09 한용민
	chulgo_edit_log masterid,finishid,"출고일자변경"

    response.write "<script>alert('재고디비에 반영되었습니다.')</script>"
	response.write "<script>location.replace('/admin/newstorage/culgolist.asp?menupos=540')</script>"

elseif mode="editmaster" then

	''출고일 - 최근 2달 내역만 가능 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  m.code, m.executedt, m.ipchulflag, m.deldt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + CStr(masterid) + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	if (yyyymmdd<>CStr(executedt)) and (not isdeleted) then
		QuickUpdateNewIpgoDetailSummary code, true
	end if

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set updt = getdate() " + VbCrlf
	if (executedt <> "") then
		sqlStr = sqlStr + " ,executedt='" + executedt + "' " + VbCrlf
	end if
	sqlStr = sqlStr + " ,comment='" + comment + "' " + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)

	dbget.Execute(sqlStr)

	' 수정로그저장		' 2021.03.09 한용민
	chulgo_edit_log masterid,finishid,"출고일,코맨트변경"

	''재고반영
	if (yyyymmdd<>CStr(executedt)) and (code<>"") and (not isdeleted)  then
		''QuickUpdateNewIpgoDetailSummary code, false
		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'','" & yyyymmdd & "'"
    	dbget.Execute sqlStr, AssignedRows

    	if (AssignedRows>0) then
    	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
    	end if

	end if

	response.write "<script>alert('수정 되었습니다.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="deldetail" then
	iid = request("masterid")
	chk= request("chk") + ",,"
	chk = split(chk, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	''출고일 - 최근 2달 내역만 가능 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	for i=0 to UBound(chk) - 1
		tmpitemgubun = ""
		tmpitemid = ""
		tmpitemoption = ""
		tmpitemno = 0

		if (trim(chk(i)) <> "") then

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set deldt=getdate()" + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))
			dbget.Execute(sqlStr)

			''재고 반영
''			if (tmpitemgubun<>"") then
''				if ipchulflag="S" then
''					QuickUpdateItemChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, tmpitemno*-1,(tmpitemno>0)
''				elseif ipchulflag="E" then
''					QuickUpdateItemEtcChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, tmpitemno*-1,(tmpitemno>0)
''				else
''					'' Nothing
''				end if
''			end if
		end if
	next

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute(sqlStr)

	if statecd<>"0" and not(isnull(statecd)) then
		' 수정로그저장		' 2021.03.09 한용민
		chulgo_edit_log iid,finishid,"상품삭제"
	end if

    '' 재고 반영
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
	end if

	response.write "<script>alert('삭제 되었습니다.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="editdetail" then
	iid = request("masterid")
	alinkcode = request("alinkcode")
	currencyUnit = request("currencyUnit")
	currencyUnit_Pos = request("currencyUnit_Pos")
	priceChange = 0

	chk= request("chk") + ",,"
	chk = split(chk, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	sellcash = request("sellcash") + ",,"
	sellcash = split(sellcash, ",")

	buycash = request("buycash") + ",,"
	buycash = split(buycash, ",")

	suplycash = request("suplycash") + ",,"
	suplycash = split(suplycash, ",")

	''출고일 - 최근 2달 내역만 가능 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if

	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	'주문서가 있을 경우 주문서 가격 정보도 같이 변경
	if alinkcode<>"" then
		if currencyUnit ="KRW" then
			priceChange=1
		end if
		if currencyUnit_Pos ="KRW" then
			priceChange=2
		end if
	end if
	if priceChange > 0 then
		''주문코드로 masteridx값 찾기
		sqlStr = "select top 1 idx, isnull(sitename,'') as sitename"
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_master"
		sqlStr = sqlStr + " where baljucode='" & Cstr(alinkcode) & "'"

		rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			masteridx = rsget("idx")
			sitename = rsget("sitename")
		end if
		rsget.close
		if sitename<>"WSLWEB" then priceChange=0
	end if

	for i=0 to UBound(chk) - 1
		'response.write  trim(chk(i))
		if (trim(chk(i)) <> "") then
			tmpitemgubun = ""
			tmpitemid = ""
			tmpitemoption = ""
			tmpitemno = 0

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close


			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(itemno(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))
			dbget.Execute(sqlStr)

			if priceChange>0 then

				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_detail" + VBCrlf
				sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
				sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,realitemno=" + CStr(itemno(CInt(chk(i)))*-1) + " " + VBCrlf
				if priceChange>1 then
				sqlStr = sqlStr + " ,foreign_sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
				sqlStr = sqlStr + " ,foreign_suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
				end if
				sqlStr = sqlStr + " where masteridx=" + CStr(masteridx) + " " + VBCrlf
				sqlStr = sqlStr + " and itemid=" + CStr(tmpitemid) + " " + VBCrlf
				sqlStr = sqlStr + " and itemoption='" + CStr(tmpitemoption) + "'"

				'response.write sqlStr & "<Br>"
				'response.end
				dbget.Execute(sqlStr)
			end if

			''재고 반영
''			if (yyyymmdd<>"") and (tmpitemgubun<>"") and (CStr(tmpitemno)<>CStr(itemno(CInt(chk(i))))) then
''				if ipchulflag="S" then
''					QuickUpdateItemChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, (itemno(CInt(chk(i)))-tmpitemno),(tmpitemno>0)
''				elseif ipchulflag="E" then
''					QuickUpdateItemEtcChulgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, (itemno(CInt(chk(i)))-tmpitemno),(tmpitemno>0)
''				else
''					'' Nothing
''				end if
''			end if
		end if
	next

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute(sqlStr)

	'주문 마스터 합계 금액 수정
	if priceChange>0 then
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
		sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
		sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrlf		'/발주시 해외 소비자가
		sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrlf			'/발주시 해외 공급가
		sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrlf			'/확정 해외 소비자가
		sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrlf		'/확정 해외 공급가
		sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
		sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
		sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
		sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*baljuitemno) as totforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*baljuitemno) as totforeign_suplycash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_sellcash*realitemno) as realforeign_sellcash" + vbCrlf
		sqlStr = sqlStr + " 	,sum(foreign_suplycash*realitemno) as realforeign_suplycash" + vbCrlf
		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
		sqlStr = sqlStr + " where masteridx="  + CStr(masteridx) + vbCrlf
		sqlStr = sqlStr + " and deldt is null" + vbCrlf
		sqlStr = sqlStr + " ) as T" + vbCrlf
		sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(masteridx)
		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr, dbget, 1
	end if

	if statecd<>"0" and not(isnull(statecd)) then
		' 수정로그저장		' 2021.03.09 한용민
		chulgo_edit_log iid,finishid,"상품수정"
	end if

    '' 재고 반영
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
	end if

	response.write "<script>alert('수정되었습니다.');</script>"
	response.write "<script>location.replace('" + refer + "')</script>"
	dbget.close()	:	response.End

elseif mode="adddetail" then
	iid = request("masterid")

	''출고일 - 최근 2달 내역만 가능 함. 입고날짜 변경건은 무시.
	sqlStr = "select top 1  m.executedt, m.ipchulflag, m.statecd"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		ipchulflag = rsget("ipchulflag")
		yyyymmdd = rsget("executedt")
		statecd = rsget("statecd")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	'''2.온라인 입고 디테일 추가
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))
			itemno = CInt(trim(itemnoarr(i))) * -1
			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = ""
			itemoptionname = ""

			''sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			''sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			''sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			''sqlStr = sqlStr + " values('" + request("code") + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			''rsget.Open sqlStr,dbget,1

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + request("code") + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + CStr(itemno) + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1

			''재고 반영
''			if ipchulflag="S" then
''				QuickUpdateItemChulgoSummary  yyyymmdd, itemgubun, itemid, itemoption, itemno,(itemno>0)
''			elseif ipchulflag="E" then
''				QuickUpdateItemEtcChulgoSummary  yyyymmdd, itemgubun, itemid, itemoption, itemno,(itemno>0)
''			else
''				'' Nothing
''			end if

		end if
	next

	'// 상품명이 없는 내역만 업데이트, skyer9, 2016-11-28
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemname=''"
	rsget.Open sqlStr, dbget, 1

	'''2.온라인 입고 마스타 업데이트
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=IsNull(T.totsell,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.totsupp,0)" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.totbuy,0)" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	if statecd<>"0" and not(isnull(statecd)) then
		' 수정로그저장		' 2021.03.09 한용민
		chulgo_edit_log masterid,finishid,"상품추가"
	end if

    '' 재고 반영
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('재고디비에 " & AssignedRows & "열 반영되었습니다.')</script>"
	end if

	response.write "<scr" + "ipt>location.replace('" + refer + "')</sc" + "ript>"
	dbget.close()	:	response.End
elseif mode="wichulgoconv" then
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set mwgubun='H'" + vbCrlf
	sqlStr = sqlStr + " where id=" + CStr(request("id"))
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('OK')</script>"
	response.write "<script>window.close()</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
