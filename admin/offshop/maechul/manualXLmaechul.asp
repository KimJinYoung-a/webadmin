<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ����
' History : 2012.12.11 �̻� ����
'			2013.04.23 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offmanualmeachulcls.asp"-->

<%
dim mode, nowdate, jungsandate, errbarcodearr, errselldayarr, errshopidarr
dim orgdata, dataarr, onedataarr, isdatavalidarr, failtypearr ,adddataarr, oneadddataarr
dim sellday, shopid, barcode, sellprice, itemno, NumExcelCols ,totcnt, failcnt, failtype
dim itemgubun, itemid, itemoption, shopname, itemname, itemoptionname, yyyymmddarr, yyyymmdd
dim shopidarr, ordermasteridx, datainserted, orderno, posid, result, i, j, sqlStr
dim temporderidx, temporderidxarr, ErrStr
dim imaechulgubun, payMethod
	mode 			= requestCheckVar(request("mode"),32)
	orgdata 		= request("orgdata")

NumExcelCols	= 5		'�׸��
posid			= 99
nowdate = now()
jungsandate = year(nowdate) & "-" & Format00(2,month(nowdate)) & "-" & "10"

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

'// ============================================================================
'' uploadorder		�����ڷ� ���� �� ���ε�
'' regtemporder		���ε� �ڷ� ������
'' deltemporder		���ε� ����
'// ============================================================================
'' if (mode = "") then
'' 	mode = "chkorder"
'' end if

'// �����ڷ� ���� �� ���ε�
if (mode = "uploadorder") then

	'// ���� ���г��� ����
	sqlStr = " update db_temp.dbo.tbl_shopjumun_ordertemp set isusing = 'N' where ordertempstatus <> '9' and regadminid = '" + CStr(session("ssBctId")) + "' and isusing = 'Y' "
	rsget.Open sqlStr,dbget,1

	dataarr = split(orgdata, vbCrLf)

	totcnt = 0
	failcnt = 0

	for i = LBound(dataarr) to UBound(dataarr)

		onedataarr = split(dataarr(i),chr(9))

		if UBound(onedataarr) >= NumExcelCols then
			totcnt = totcnt + 1

			sellday		= Trim(onedataarr(0))
			shopid		= Trim(onedataarr(1))
			barcode		= Trim(onedataarr(2))
			sellprice	= Trim(onedataarr(3))
			itemno		= Trim(onedataarr(4))
			payMethod	= Trim(onedataarr(5))

			'// 2012-12-01 -> 20121201
			sellday = Replace(sellday, "-", "")

			'//�ǸŰ��� ���� �Ѵ� ���̳ʽ� �ϰ��..���ϸ� �÷����� ����.. �ðܳ�
			if left(trim(sellprice),1)="-" and left(trim(itemno),1)="-" then
				failcnt = failcnt + 1
                ErrStr = ErrStr + CStr(i+1) + "�� ���� : " + dataarr(i) + " \n�ǸŰ��� ���� �Ѵ� ���̳ʽ� ���� �ɼ� �����ϴ�.\n���̳ʽ� �ֹ� �Է½� ������ ���̳ʽ��� �Է����ּ���\n"
			end if

			if (sellday = "") or (shopid = "") or (barcode = "") or (sellprice = "") or (itemno = "") or (payMethod = "") then
				failcnt = failcnt + 1
                ErrStr = ErrStr + CStr(i+1) + "�� ���� : " + dataarr(i) + "\n"
            elseif (Len(sellday) <> 8) or (Len(barcode) < 4) or (Not IsNumeric(sellprice)) or (Not IsNumeric(itemno)) then
				failcnt = failcnt + 1
                ErrStr = ErrStr + CStr(i+1) + "�� ���� : " + dataarr(i) + "\n"
			else
				sqlStr = "insert into db_temp.dbo.tbl_shopjumun_ordertemp (" + vbcrlf
				sqlStr = sqlStr + " yyyymmdd, shopid, barcode, ordertempstatus" + vbcrlf
				sqlStr = sqlStr + " , sellprice, itemno, payMethod, regdate, isusing, regadminid)" + vbcrlf
				sqlStr = sqlStr + " select" + vbcrlf
				sqlStr = sqlStr + " 	'" + CStr(sellday) + "' " + vbcrlf
				sqlStr = sqlStr + " 	, '" + CStr(shopid) + "' " + vbcrlf
				sqlStr = sqlStr + " 	, '" + CStr(barcode) + "', 0 " + vbcrlf
				sqlStr = sqlStr + " 	, '" + CStr(sellprice) + "' " + vbcrlf
				sqlStr = sqlStr + " 	, '" + CStr(itemno) + "' " + vbcrlf
				sqlStr = sqlStr + " 	, '" + CStr(payMethod) + "' " + vbcrlf
				sqlStr = sqlStr + " 	, getdate(), 'Y', '" + CStr(session("ssBctId")) + "' " + vbcrlf
				''response.write sqlStr & "<br>"

				rsget.Open sqlStr,dbget,1
			end if
		end if
	next

	if (failcnt <> 0) then
		sqlStr = " update db_temp.dbo.tbl_shopjumun_ordertemp set isusing = 'N' where ordertempstatus = 0 and regadminid = '" + CStr(session("ssBctId")) + "' and isusing = 'Y' "
		rsget.Open sqlStr,dbget,1
	end if

	if (failcnt = 0) then
		ErrStr = ""

		'// ������ڵ�
		sqlStr = " update t "
		sqlStr = sqlStr + " set t.itemgubun = s.itemgubun, t.itemid = s.itemid, t.itemoption = s.itemoption "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_option_stock s "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		t.barcode = s.barcode "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	and t.itemid is NULL "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		sqlStr = " select t.barcode "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	and t.itemid is NULL "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		errbarcodearr = ""
		if not rsget.EOF  then
			do until rsget.eof
				errbarcodearr = errbarcodearr + "|" + rsget("barcode")
				rsget.MoveNext
			loop
		end if
		rsget.close

		if (errbarcodearr <> "") then
			errbarcodearr = Split(errbarcodearr, "|")

			for i = 0 to UBound(errbarcodearr)
				barcode = Trim(errbarcodearr(i))

				itemgubun	= ""
				itemid		= ""
				itemoption	= ""


				if (barcode <> "") then
					'// �ٹ����� ���ڵ�(101111110000)
					if (BF_IsMaybeTenBarcode(barcode) = True) then

						sqlStr = "SELECT top 1 itemgubun, shopitemid as itemid, itemoption "
						sqlStr = sqlStr & " FROM [db_shop].[dbo].[tbl_shop_item]"
						sqlStr = sqlStr & " WHERE itemgubun = '" & BF_GetItemGubun(barcode) & "'"
						sqlStr = sqlStr & " AND shopitemid = '" & BF_GetItemId(barcode) & "'"
						sqlStr = sqlStr & " AND itemoption = '" & BF_GetItemOption(barcode) & "'"

						''response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1
						If Not rsget.Eof Then
							itemgubun = rsget("itemgubun")
							itemid = rsget("itemid")
							itemoption = rsget("itemoption")
						End If
						rsget.close()

						if (itemid <> "") then
							sqlStr = " update t "
							sqlStr = sqlStr + " set t.itemgubun = '" + CStr(itemgubun) + "', t.itemid = '" + CStr(itemid) + "', t.itemoption = '" + CStr(itemoption) + "' "
							sqlStr = sqlStr + " from "
							sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
							sqlStr = sqlStr + " where "
							sqlStr = sqlStr + " 	1 = 1 "
							sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
							sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
							sqlStr = sqlStr + " 	and t.itemid is NULL "
							sqlStr = sqlStr + " 	and t.barcode = '" + CStr(barcode) + "' "
							sqlStr = sqlStr + " 	and t.isusing = 'Y' "

							''response.write sqlStr & "<br>"
							rsget.Open sqlStr,dbget,1

							errbarcodearr(i) = ""
						end if
					end if
				end if
			next
		end if

		sqlStr = " select t.yyyymmdd, t.shopid, t.barcode "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	and t.itemid is NULL "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		if not rsget.EOF  then
			do until rsget.eof
				ErrStr = ErrStr + "���� : " + CStr(rsget("yyyymmdd")) + "," + CStr(rsget("shopid")) + "," + CStr(rsget("barcode")) + " ���ڵ� ��� �ʵ�\n"
				failcnt = failcnt + 1
				rsget.MoveNext
			loop
		end if
		rsget.close

		sqlStr = " update t "
		sqlStr = sqlStr + " set t.failtype = 'B' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	and t.itemid is NULL "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
	end if

	if (failcnt = 0) then
		'// ���� ��ü �ߺ�
		sqlStr = " select t.yyyymmdd, t.shopid, t.barcode "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " left join db_temp.dbo.tbl_shopjumun_ordertemp tt "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.yyyymmdd = tt.yyyymmdd "
		sqlStr = sqlStr + " 	and t.shopid = tt.shopid "
		sqlStr = sqlStr + " 	and t.barcode = tt.barcode "
		sqlStr = sqlStr + " 	and t.regadminid = tt.regadminid "
		sqlStr = sqlStr + " 	and t.idx <> tt.idx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = '0' "
		sqlStr = sqlStr + " 	and tt.ordertempstatus not in ('0', '9') "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "
		sqlStr = sqlStr + " 	and tt.isusing = 'Y' "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		if not rsget.EOF  then
			do until rsget.eof
				ErrStr = ErrStr + "���� : " + CStr(rsget("yyyymmdd")) + "," + CStr(rsget("shopid")) + "," + CStr(rsget("barcode")) + " ���� ��ü �ߺ�\n"
				failcnt = failcnt + 1
				rsget.MoveNext
			loop
		end if
		rsget.close

		sqlStr = " update t set t.failtype = 'U' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " left join db_temp.dbo.tbl_shopjumun_ordertemp tt "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.yyyymmdd = tt.yyyymmdd "
		sqlStr = sqlStr + " 	and t.shopid = tt.shopid "
		sqlStr = sqlStr + " 	and t.barcode = tt.barcode "
		sqlStr = sqlStr + " 	and t.regadminid = tt.regadminid "
		sqlStr = sqlStr + " 	and t.idx <> tt.idx "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and t.ordertempstatus = '0' "
		sqlStr = sqlStr + " 	and tt.ordertempstatus not in ('0', '9') "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "
		sqlStr = sqlStr + " 	and tt.isusing = 'Y' "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1


		'// �����ߺ�üũ
		sqlStr = " select t.yyyymmdd, t.shopid, t.barcode "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m "
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.idx = d.masteridx "
		sqlStr = sqlStr + " join db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and convert(datetime, t.yyyymmdd, 112) = m.IXyyyymmdd "
		sqlStr = sqlStr + " 	and m.shopid = t.shopid "
		sqlStr = sqlStr + " 	and d.itemgubun = t.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid = t.itemid "
		sqlStr = sqlStr + " 	and d.itemoption = t.itemoption "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and maechulgubun <> 'POS' "		'// ������ ���⸸

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		if not rsget.EOF  then
			do until rsget.eof
				ErrStr = ErrStr + "���� : " + CStr(rsget("yyyymmdd")) + "," + CStr(rsget("shopid")) + "," + CStr(rsget("barcode")) + " ���� �ߺ�\n"
				failcnt = failcnt + 1
				rsget.MoveNext
			loop
		end if
		rsget.close

		sqlStr = " update t set failtype = 'D'  "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m "
		sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	m.idx = d.masteridx "
		sqlStr = sqlStr + " join db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " on "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and convert(datetime, t.yyyymmdd, 112) = m.IXyyyymmdd "
		sqlStr = sqlStr + " 	and m.shopid = t.shopid "
		sqlStr = sqlStr + " 	and d.itemgubun = t.itemgubun "
		sqlStr = sqlStr + " 	and d.itemid = t.itemid "
		sqlStr = sqlStr + " 	and d.itemoption = t.itemoption "
		sqlStr = sqlStr + " 	and t.isusing = 'Y' "
		sqlStr = sqlStr + " 	and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 	and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and m.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and maechulgubun <> 'POS' "		'// ������ ���⸸

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1


		sqlStr = " select t.yyyymmdd, t.shopid, t.barcode "
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	join db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and i.itemgubun = t.itemgubun "
		sqlStr = sqlStr + " 		and i.shopitemid = t.itemid "
		sqlStr = sqlStr + " 		and i.itemoption = t.itemoption "
		sqlStr = sqlStr + " 		and t.isusing = 'Y' "
		sqlStr = sqlStr + " 		and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 		and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid=t.shopid"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and s.shopid is NULL "
		sqlStr = sqlStr + " 		and t.isusing = 'Y' "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		if not rsget.EOF  then
			do until rsget.eof
				ErrStr = ErrStr + "���� : " + CStr(rsget("yyyymmdd")) + "," + CStr(rsget("shopid")) + "," + CStr(rsget("barcode")) + " ������\n"
				failcnt = failcnt + 1
				rsget.MoveNext
			loop
		end if
		rsget.close

		sqlStr = " 	update t set t.failtype = 'J' " + vbcrlf
		sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
		sqlStr = sqlStr + " 	join db_temp.dbo.tbl_shopjumun_ordertemp t "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and i.itemgubun = t.itemgubun "
		sqlStr = sqlStr + " 		and i.shopitemid = t.itemid "
		sqlStr = sqlStr + " 		and i.itemoption = t.itemoption "
		sqlStr = sqlStr + " 		and t.isusing = 'Y' "
		sqlStr = sqlStr + " 		and t.ordertempstatus = 0 "
		sqlStr = sqlStr + " 		and t.regadminid = '" + CStr(session("ssBctId")) + "' "
		sqlStr = sqlStr + " 	left join db_shop.dbo.tbl_shop_designer s" & VbCRLF
		sqlStr = sqlStr + " 		on s.shopid=t.shopid"
		sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
		sqlStr = sqlStr + " 	where "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and s.shopid is NULL "

		''response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1
	end if

	if (failcnt = 0) then
		sqlStr = " update db_temp.dbo.tbl_shopjumun_ordertemp set ordertempstatus = '1' where ordertempstatus = 0 and regadminid = '" + CStr(session("ssBctId")) + "' and isusing = 'Y' "
		rsget.Open sqlStr,dbget,1
	end if

'// �����ڷ� ���
elseif (mode = "regtemporder") then

	dataarr = split(orgdata, ",")

	totcnt = 0
	failcnt = 0

	'// �����ߺ�üũ
	sqlStr = " select count(*) as cnt  "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m "
	sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	m.idx = d.masteridx "
	sqlStr = sqlStr + " join db_temp.dbo.tbl_shopjumun_ordertemp t "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and convert(datetime, t.yyyymmdd, 112) = m.IXyyyymmdd "
	sqlStr = sqlStr + " 	and m.shopid = t.shopid "
	sqlStr = sqlStr + " 	and d.itemgubun = t.itemgubun "
	sqlStr = sqlStr + " 	and d.itemid = t.itemid "
	sqlStr = sqlStr + " 	and d.itemoption = t.itemoption "
	sqlStr = sqlStr + " 	and t.isusing = 'Y' "
	''sqlStr = sqlStr + " 	and t.ordertempstatus = 1 "
	sqlStr = sqlStr + " 	and t.idx in (" + CStr(orgdata) + ") "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and maechulgubun <> 'POS' "		'// ������ ���⸸

	'response.write sqlStr & "<br>"
	rsget.Open sqlStr,dbget,1
	If Not rsget.Eof Then
		failcnt = failcnt + rsget("cnt")
	End If
	rsget.close()

	sqlStr = " update t set failtype = 'D'  "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m "
	sqlStr = sqlStr + " join [db_shop].[dbo].tbl_shopjumun_detail d "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	m.idx = d.masteridx "
	sqlStr = sqlStr + " join db_temp.dbo.tbl_shopjumun_ordertemp t "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and convert(datetime, t.yyyymmdd, 112) = m.IXyyyymmdd "
	sqlStr = sqlStr + " 	and m.shopid = t.shopid "
	sqlStr = sqlStr + " 	and d.itemgubun = t.itemgubun "
	sqlStr = sqlStr + " 	and d.itemid = t.itemid "
	sqlStr = sqlStr + " 	and d.itemoption = t.itemoption "
	sqlStr = sqlStr + " 	and t.isusing = 'Y' "
	''sqlStr = sqlStr + " 	and t.ordertempstatus = 1 "
	sqlStr = sqlStr + " 	and t.idx in (" + CStr(orgdata) + ") "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and m.cancelyn <> 'Y' "
	sqlStr = sqlStr + " 	and maechulgubun <> 'POS' "		'// ������ ���⸸

	''response.write sqlStr & "<br>"
	rsget.Open sqlStr,dbget,1

	if (failcnt = 0) then
		yyyymmddarr = ""
		shopidarr = ""

		'// �Է��� ����-������
		sqlStr = " select distinct shopid, convert(datetime, yyyymmdd, 112) as yyyymmdd from db_temp.dbo.tbl_shopjumun_ordertemp "
		sqlStr = sqlStr + " where ordertempstatus = 1 and idx in (" + CStr(orgdata) + ") and isusing = 'Y' "
		sqlStr = sqlStr + " order by shopid, convert(datetime, yyyymmdd, 112) "
		rsget.Open sqlStr,dbget,1
		''response.write sqlStr &"<br>"

		if not rsget.EOF  then
			do until rsget.eof
				yyyymmddarr = yyyymmddarr & "|" + CStr(rsget("yyyymmdd"))
				shopidarr = shopidarr & "|" + CStr(rsget("shopid"))
				rsget.MoveNext
			loop
		end if
		rsget.close

		yyyymmddarr = Split(yyyymmddarr, "|")
		shopidarr = Split(shopidarr, "|")

		for i = 0 to UBound(yyyymmddarr)
			yyyymmdd = Trim(yyyymmddarr(i))
			shopid = Trim(shopidarr(i))

			if (yyyymmdd <> "") then
                ''���ⱸ��  /2013/12/17 �߰�
                imaechulgubun=""

                sqlStr = "select isNULL(tplcompanyid,'MANUAL') as maechulgubun"
                sqlStr = sqlStr&" from db_partner.dbo.tbl_partner"
                sqlStr = sqlStr&" where id='"&shopid&"'"
                rsget.Open sqlStr,dbget,1
				if Not rsget.Eof then
					imaechulgubun=rsget("maechulgubun")
				end if
				rsget.close

                if (imaechulgubun="") then
                    imaechulgubun="MANUAL"
                end if

				'// ================================================================

				'//�δ� ���� ���� �ϰ��
				if datediff("m", Left(yyyymmdd,10) , nowdate) >= 2 then
					response.write "<script language='javascript'>"
					response.write "	alert('�δ� ���������� �Է� �ϽǼ� �����ϴ�.');"
					response.write "	location.href='"&refer&"';"
					response.write "</script>"
					response.end	:	dbget.close()
				end if

				'//������ ���� �Է½�
				if datediff("m", Left(yyyymmdd,10) , nowdate) < 0 then
					response.write "<script language='javascript'>"
					response.write "	alert('������ ������ �Է��� �Ұ��� �մϴ�.');"
					response.write "	location.href='"&refer&"';"
					response.write "</script>"
					response.end	:	dbget.close()
				end if

'				'//������ ���� �Է½�
'				if datediff("m", Left(yyyymmdd,10) , nowdate) = 1 then
'					if datediff("d", jungsandate , nowdate) > 0 then
'						response.write "<script language='javascript'>"
'						response.write "	alert('�������� ������ ������ �� ��¥ �Դϴ�.');"
'						response.write "</script>"
'						if Not C_ADMIN_AUTH then
'							response.end	:	dbget.close()
'						else
'							response.write "<script type='text/javascript'>"
'							response.write "	alert('[�����ڱ���]\n\n��������մϴ�.');"
'							response.write "</script>"
'						end if
'					end if
'				end if

				' ��ǰ���� ��� �ֹ���ȣ�� �����Ѵ�. �ֹ� �ϳ��� �� ���� ����. ���� ����̶� ��ǰ��ȣ ���� �ǸŰ��� �������ϰ� ���� ������ �ؼ� ������ �ȵ�.
				if isarray(dataarr) then
					for j = 0 to ubound(dataarr)
					'// �ֹ� ������ ����
					orderno = manualordernomake_off(shopid, posid)

					'/�̹������ϴ� �ֹ���ȣ���� üũ
					sqlStr = "select count(idx) as cnt"
					sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shopjumun_master"
					sqlStr = sqlStr + " where orderno='"&orderno&"'"

					'response.write sqlStr &"<br>"
					rsget.Open sqlStr,dbget,1

					if Not rsget.Eof then
						if (rsget("cnt")>0) then result = "Y"
					end if
					rsget.close

					if result = "Y" then
						response.write "<script language='javascript'>"
						response.write "	alert('�ֹ���ȣ�� �̹� ���� �մϴ�. ������ ���ǿ��.');"
						response.write "</script>"
						response.end	:	dbget.close()
					end if
					result = ""

					sqlStr = "select * from [db_shop].[dbo].tbl_shopjumun_master where 1=0"
					rsget.Open sqlStr,dbget,1,3
					rsget.AddNew

					rsget("orderno")    = orderno
					rsget("shopid")     = shopid
					rsget("totalsum")   = 0
					rsget("realsum")    = 0
					rsget("jumundiv")   = "00"
					rsget("jumunmethod") = "01"
					rsget("shopregdate") = Left(yyyymmdd,10)
					rsget("cancelyn")   = "N"
					rsget("shopidx")    = "0"
					rsget("spendmile")  = "0"
					rsget("pointuserno") = ""
					rsget("gainmile") = "0"
					rsget("cashsum")    = 0
					rsget("cardsum")    = "0"
					rsget("casherid")   = session("ssBctId")
					rsget("GiftCardPaySum") = "0"
					rsget("CardAppNo")      = ""
					rsget("CashReceiptNo")  = ""
					rsget("CashreceiptGubun") = ""
					rsget("CardInstallment")  = ""
					rsget("IXyyyymmdd") = Left(yyyymmdd,10)
					rsget("tableno")  = "0"
					rsget("TenGiftCardPaySum")  = "0"
					rsget("TenGiftCardMatchCode")  = ""
					rsget("refOrderNo")  = ""
					rsget("maechulgubun")  = imaechulgubun

					rsget.update
						ordermasteridx = rsget("idx")
					rsget.close

					'// zoneidx �� ���ε峻�� �ε����� �žҴ� �ֹ�������IDX ������ �����.

					'//������ ���̺� ���
					sqlStr = "insert into [db_shop].[dbo].tbl_shopjumun_detail" + vbcrlf
					sqlStr = sqlStr + " ( masteridx, orderno, itemgubun, itemid, itemoption" + vbcrlf
					sqlStr = sqlStr + " , itemno, itemname, itemoptionname, sellprice, realsellprice" + vbcrlf
					sqlStr = sqlStr + " , suplyprice" + vbcrlf
					sqlStr = sqlStr + " , shopbuyprice" + vbcrlf
					sqlStr = sqlStr + " , makerid, jungsanid, cancelyn" + vbcrlf
					sqlStr = sqlStr + " , shopidx, itempoint, discountKind, Iorgsellprice, Ishopitemprice" + vbcrlf
					sqlStr = sqlStr + " , jcomm_cd, zoneidx, addtaxcharge, vatinclude, payMethod)" + vbcrlf
					sqlStr = sqlStr + " 	select" + vbcrlf
					sqlStr = sqlStr + " 	'"& ordermasteridx &"','"&orderno&"',i.itemgubun ,i.shopitemid ,i.itemoption" + vbcrlf
					sqlStr = sqlStr + " 	,t.itemno, i.shopitemname, i.shopitemoptionname,t.sellprice,t.sellprice" + vbcrlf
					sqlStr = sqlStr + " 	,(CASE" & VbCRLF
					sqlStr = sqlStr + " 		when isnull(ii.mwdiv,'')='M' and s.comm_cd not in ('B012')" & VbCRLF		'//�¶��θ����̰�, ��ü��Ź�� �ƴϸ� �¶��θ��԰���
					sqlStr = sqlStr + " 			THEN isnull(ii.buycash,0)" & VbCRLF
					'sqlStr = sqlStr + " 		when i.shopsuplycash = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/���԰��� 0 ,������Ź, ��ü��Ź ,�����Ź
					sqlStr = sqlStr + " 		when i.shopsuplycash = 0" & VbCRLF		'���԰� �� ������ ������
					sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultmargin,100))/100)" & VbCRLF
					sqlStr = sqlStr + " 		else i.shopsuplycash" & VbCRLF
					sqlStr = sqlStr + "			end) as shopsuplycash" & VbCRLF
					sqlStr = sqlStr + " 	,(CASE" & VbCRLF
					'sqlStr = sqlStr + " 		when i.shopbuyprice = 0 and s.comm_cd in ('B011','B012','B013')" & VbCRLF		'/������� 0 ,������Ź, ��ü��Ź ,�����Ź
					sqlStr = sqlStr + " 		when i.shopbuyprice = 0" & VbCRLF		'������� �� ������ ������
					sqlStr = sqlStr + " 			then convert(int,i.shopitemprice*(100-IsNULL(s.defaultsuplymargin,100))/100)" & VbCRLF
					sqlStr = sqlStr + "			else i.shopbuyprice" & VbCRLF
					sqlStr = sqlStr + "			end) as shopbuyprice" & VbCRLF
					sqlStr = sqlStr + " 	, i.makerid, i.makerid, 'N'" + vbcrlf
					sqlStr = sqlStr + " 	,'0','0','0', i.orgsellprice, i.shopitemprice" + vbcrlf
					sqlStr = sqlStr + " 	, s.comm_cd, t.idx, 0, i.vatinclude" + vbcrlf
					sqlStr = sqlStr + " 	, isNull(t.payMethod,'C') " + vbcrlf
					sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shop_item i" + vbcrlf
					sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shop_designer s" & VbCRLF
					sqlStr = sqlStr + " 		on s.shopid='"&shopid&"'"
					sqlStr = sqlStr + " 		and i.makerid=s.makerid" & VbCRLF
					sqlStr = sqlStr + " 	join db_temp.dbo.tbl_shopjumun_ordertemp t "
					sqlStr = sqlStr + " 	on "
					sqlStr = sqlStr + " 		1 = 1 "
					sqlStr = sqlStr + " 		and convert(datetime, t.yyyymmdd, 112) = '" + CStr(yyyymmdd) + "' "
					sqlStr = sqlStr + " 		and t.shopid = '" + CStr(shopid) + "' "
					sqlStr = sqlStr + " 		and i.itemgubun = t.itemgubun "
					sqlStr = sqlStr + " 		and i.shopitemid = t.itemid "
					sqlStr = sqlStr + " 		and i.itemoption = t.itemoption "
					sqlStr = sqlStr + " 		and t.isusing = 'Y' "
					sqlStr = sqlStr + " 		and t.ordertempstatus = 1 "
					'sqlStr = sqlStr + " 		and t.idx in (" + CStr(orgdata) + ") "
					sqlStr = sqlStr + " 		and t.idx in ("& dataarr(j) &") "
					sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item ii" & VbCRLF
					sqlStr = sqlStr + " 		on i.shopitemid = ii.itemid" & VbCRLF
					sqlStr = sqlStr + " 		and i.itemgubun = '10'" & VbCRLF

					'response.write sqlStr & "<Br>"
					dbget.Execute sqlStr

					sqlStr = " update t "
					sqlStr = sqlStr + " set t.shopjumundetailidx = d.idx "
					sqlStr = sqlStr + " from "
					sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
					sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shopjumun_detail d "
					sqlStr = sqlStr + " 	on "
					sqlStr = sqlStr + " 		t.idx = d.zoneidx "
					sqlStr = sqlStr + " where "
					sqlStr = sqlStr + " 	1 = 1 "
					'sqlStr = sqlStr + " 	and t.idx in (" + CStr(orgdata) + ") "
					sqlStr = sqlStr + " 	and t.idx in ("& dataarr(j) &") "

					'response.write sqlStr & "<Br>"
					dbget.Execute sqlStr

					sqlStr = " update d "
					sqlStr = sqlStr + " set d.zoneidx = NULL "
					sqlStr = sqlStr + " from "
					sqlStr = sqlStr + " 	db_temp.dbo.tbl_shopjumun_ordertemp t "
					sqlStr = sqlStr + " 	join [db_shop].[dbo].tbl_shopjumun_detail d "
					sqlStr = sqlStr + " 	on "
					sqlStr = sqlStr + " 		t.idx = d.zoneidx "
					sqlStr = sqlStr + " where "
					sqlStr = sqlStr + " 	1 = 1 "
					'sqlStr = sqlStr + " 	and t.idx in (" + CStr(orgdata) + ") "
					sqlStr = sqlStr + " 	and t.idx in ("& dataarr(j) &") "

					'response.write sqlStr & "<Br>"
					dbget.Execute sqlStr

					'//������ ���̺� �ջ�
					sqlStr = "update m" + vbcrlf
					sqlStr = sqlStr + " set m.totalsum = t.sellprice" + vbcrlf
					sqlStr = sqlStr + " ,m.realsum = t.realsellprice" + vbcrlf
					sqlStr = sqlStr + " ,m.cashsum = t.cashSellprice" + vbcrlf
					sqlStr = sqlStr + " ,m.extPaySum = t.extSellprice" + vbcrlf
					sqlStr = sqlStr + " ,m.jumunMethod = Case When t.extSellprice>0 then '09' else '01' end " + vbcrlf
					sqlStr = sqlStr + " from db_shop.dbo.tbl_shopjumun_master m" + vbcrlf
					sqlStr = sqlStr + " join (" + vbcrlf
					sqlStr = sqlStr + " 	select" + vbcrlf
					sqlStr = sqlStr + " 	orderno ,sum((d.sellprice+addtaxcharge) * d.itemno) as sellprice" + vbcrlf
					sqlStr = sqlStr + " 	,sum((d.realsellprice+addtaxcharge) * d.itemno) as realsellprice" + vbcrlf
					sqlStr = sqlStr + " 	,sum(Case When payMethod='C' then (d.realsellprice+addtaxcharge) * d.itemno else 0 end) as cashSellprice" + vbcrlf
					sqlStr = sqlStr + " 	,sum(Case When payMethod='E' then (d.realsellprice+addtaxcharge) * d.itemno else 0 end) as extSellprice" + vbcrlf
					sqlStr = sqlStr + " 	,sum((d.suplyprice+addtaxcharge) * d.itemno) as suplyprice" + vbcrlf
					sqlStr = sqlStr + " 	from [db_shop].[dbo].tbl_shopjumun_detail d" + vbcrlf
					sqlStr = sqlStr + " 	where d.cancelyn = 'N'" + vbcrlf
					sqlStr = sqlStr + " 	and d.orderno = '"&orderno&"'" + vbcrlf
					sqlStr = sqlStr + " 	group by orderno" + vbcrlf
					sqlStr = sqlStr + " ) as t" + vbcrlf
					sqlStr = sqlStr + " 	on m.orderno = t.orderno" + vbcrlf
					sqlStr = sqlStr + " 	and m.cancelyn = 'N'" + vbcrlf
					sqlStr = sqlStr + " where m.orderno = '"&orderno&"'"

					'response.write sqlStr & "<Br>"
					dbget.Execute sqlStr

					'// �ߺ��Է� ����
					sqlStr = "[db_shop].[dbo].[usp_TEN_Shop_ManualOrder_DuppRemove] '" & orderno & "'"
					dbget.Execute sqlStr

					''��� ������Ʈ(No tran)
					sqlStr = "exec db_summary.dbo.sp_Ten_Shop_Stock_RegOrder '" & orderno & "'"
					dbget.Execute sqlStr
					''response.write "TODO : ���ݿ�<br><br><br><br>"

					sqlStr = " update db_temp.dbo.tbl_shopjumun_ordertemp set ordertempstatus = 9" & vbcrlf
					sqlStr = sqlStr & " where convert(datetime, yyyymmdd, 112) = '" + CStr(yyyymmdd) + "' and shopid = '" + CStr(shopid) + "' and ordertempstatus = 1" & vbcrlf
					'sqlStr = sqlStr & " and idx in (" + CStr(orgdata) + ")" & vbcrlf
					sqlStr = sqlStr & " and idx in ("& dataarr(j) &")" & vbcrlf
					sqlStr = sqlStr & " and isusing = 'Y' and shopjumundetailidx is not NULL" & vbcrlf

					'response.write sqlStr & "<br>"
					dbget.execute sqlStr
					next
				end if

				response.write "<script>alert('" + CStr(shopid) + " ���� - " + CStr(yyyymmdd) + " �� ���� ��ϿϷ�');</script>"

			end if
		next

	end if

elseif (mode = "deltemporder") then
	'// ���ε� ����

	sqlStr = " update db_temp.dbo.tbl_shopjumun_ordertemp set isusing = 'N' "
	sqlStr = sqlStr + " where ordertempstatus <> 9 and idx in (" + CStr(orgdata) + ") and isusing = 'Y' "
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('�����Ǿ����ϴ�.');</script>"
end if

'// =============================================================================
dim oCManualMeachul
set oCManualMeachul = new CManualMeachul
	oCManualMeachul.FPageSize = 100
	oCManualMeachul.FCurrPage = 1

	if (mode = "regtemporder") and (failcnt = 0) then
		oCManualMeachul.FRectCurrentInsertOnly = "Y"
		oCManualMeachul.FRectIdxArr = orgdata
	else
		oCManualMeachul.FRectExcludeRegFinish = "Y"
	end if

	oCManualMeachul.GetList

dim oCFailManualMeachul
set oCFailManualMeachul = new CManualMeachul
	oCFailManualMeachul.FPageSize = 100
	oCFailManualMeachul.FCurrPage = 1
	oCFailManualMeachul.FRectRegAdminID = session("ssBctId")
	oCFailManualMeachul.GetFailList

%>

<script language='javascript'>

function checkClick() {
	var frm = document.frm;

	if (confirm('���� �ڷḦ üũ�մϴ�.\n\n�����Ͻðڽ��ϱ�?')) {
		frm.mode.value="uploadorder";
		frm.submit();
	}
}

function saveClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("������ ������ �����ϴ�.");
		return;
	}

	if (confirm("������ �Ͻðڽ��ϱ�?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="regtemporder";
		frm.submit();
	}
}

function delClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("������ ����� �����ϴ�.");
		return;
	}

	if (confirm("���� �Ͻðڽ��ϱ�?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="deltemporder";
		frm.submit();
	}
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

function clearData() {
	var frm = document.frm;
	frm.orgdata.value = "";
}

</script>

<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name=frm method=post onSubmit="return false;">
<input type="hidden" name="mode" value="">
<tr>
	<td>
		<p><span style="color:red;">������ �и�</span></p>
		<p>���������, ���óID, ���ڵ�, �ǸŰ�, ����, �������<span style="color:green;">(����:C, ��Ÿ:E)</span><br>
		sellday, shopid, barcode, sellprice, itemno, payMethod</p>
		<p><span style="color:red;">���� ������ ����� �ȵ˴ϴ�.</span></p>
		<% if (ErrStr <> "") then %>
			<p><br /><span style="color:red;font-weight:bold;"><%= Replace(ErrStr, "\n", "<br>") %></span></p>
		<% end if %>
	</td>
	<td align="right" valign="bottom">
		<a href="�������θ�����FORM.xlsx" target="_blank">[����FORM]</a>
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="orgdata" cols=100 rows=5></textarea>
	</td>
</tr>
<tr>
	<td>
	<input type= button class="button" value="Clear" onClick="clearData();">
	</td>
	<td>
		<input type= button class="button" value="���ε�" onclick="checkClick()">
	</td>
</tr>
</form>
</table>

<p>

<% if oCFailManualMeachul.FResultCount <> 0 then %>

[���ε����]
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name=frmfail method=post onSubmit="return false;">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkall" disabled ></td>
	<td width="55">������</td>
	<td width="100">���óID</td>
	<td>���ó</td>
	<td width="90">���ڵ�</td>
	<td width="25">����</td>
	<td width="50">��ǰ�ڵ�</td>
	<td width="30">�ɼ�</td>
	<td>��ǰ��[<font color="blue">�ɼǸ�</font>]</td>
	<td width="60">�ǸŰ�</td>
	<td width="40">����</td>
	<td width="60">�����</td>
	<td width="70">����</td>
	<td width="70">���</td>
</tr>

	<% For i = 0 To oCFailManualMeachul.FResultCount - 1 %>

			<% if IsNull(oCFailManualMeachul.FItemList(i).Ffailtype) then %>
			<tr align="center" bgcolor="#FFFFFF">
			<% else %>
			<tr align="center" bgcolor="#CCCCCC">
			<% end if %>
				<td><input type="checkbox" disabled ></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fsellday %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fshopid %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fshopname %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fbarcode %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fitemgubun %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fitemid %></td>
				<td><%= oCFailManualMeachul.FItemList(i).Fitemoption %></td>
				<td align="left"><%= oCFailManualMeachul.FItemList(i).Fitemname %>[<font color="blue"><%= oCFailManualMeachul.FItemList(i).Fitemoptionname %></font>]</td>
				<td align="right"><%= FormatNumber(oCFailManualMeachul.FItemList(i).Fsellprice, 0) %>&nbsp;</td>
				<td align="right"><%= oCFailManualMeachul.FItemList(i).Fitemno %>&nbsp;</td>
				<td><%= oCFailManualMeachul.FItemList(i).Fregadminid %></td>
				<td><%= oCFailManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
				<td>
					<%= oCFailManualMeachul.FItemList(i).GetFailTypeName %>
				</td>
			</tr>

	<% next %>

<% end if %>

</form>
</table>

<p><br><br>

[���ε峻��]
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name=frmdetail method=post onSubmit="return false;">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkall" onClick="CheckAll(this)" ></td>
	<td width="55">������</td>
	<td width="60">����IDX</td>
	<td width="100">���óID</td>
	<td>���ó</td>
	<td width="90">���ڵ�</td>
	<td width="25">����</td>
	<td width="50">��ǰ�ڵ�</td>
	<td width="30">�ɼ�</td>
	<td>��ǰ��[<font color="blue">�ɼǸ�</font>]</td>
	<td width="60">�ǸŰ�</td>
	<td width="40">����</td>
	<td width="60">�����</td>
	<td width="70">����</td>
	<td width="70">���</td>
</tr>

<% if oCManualMeachul.FResultCount = 0 then %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="15" height="35">
			���ε峻�� ����
		</td>
	</tr>

<% else %>

<% For i = 0 To oCManualMeachul.FResultCount - 1 %>

<% if IsNull(oCManualMeachul.FItemList(i).Ffailtype) then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="#CCCCCC">
<% end if %>
	<td><input type="checkbox" id="chk_<%= i %>" name="chk_<%= i %>" value="<%= oCManualMeachul.FItemList(i).Fidx %>" <% if (oCManualMeachul.FItemList(i).Fordertempstatus <> "1") or (Not IsNull(oCManualMeachul.FItemList(i).Ffailtype)) then %>disabled<% end if %> ></td>
	<td><%= oCManualMeachul.FItemList(i).Fsellday %></td>
	<td><%= oCManualMeachul.FItemList(i).Fshopjumundetailidx %></td>
	<td><%= oCManualMeachul.FItemList(i).Fshopid %></td>
	<td><%= oCManualMeachul.FItemList(i).Fshopname %></td>
	<td><%= oCManualMeachul.FItemList(i).Fbarcode %></td>
	<td><%= oCManualMeachul.FItemList(i).Fitemgubun %></td>
	<td><%= oCManualMeachul.FItemList(i).Fitemid %></td>
	<td><%= oCManualMeachul.FItemList(i).Fitemoption %></td>
	<td align="left"><%= oCManualMeachul.FItemList(i).Fitemname %>[<font color="blue"><%= oCManualMeachul.FItemList(i).Fitemoptionname %></font>]</td>
	<td align="right"><%= FormatNumber(oCManualMeachul.FItemList(i).Fsellprice, 0) %>&nbsp;</td>
	<td align="right"><%= oCManualMeachul.FItemList(i).Fitemno %>&nbsp;</td>
	<td><%= oCManualMeachul.FItemList(i).Fregadminid %></td>
	<td><%= oCManualMeachul.FItemList(i).GetOrderTempStatusName %></td>
	<td>
		<%= oCManualMeachul.FItemList(i).GetFailTypeName %>
	</td>
</tr>
<% next %>
<% end if %>

</form>
</table>

<div align="center">
	<input type= button class="button" value="������" onclick="saveClick()">
	&nbsp;&nbsp;&nbsp;
	<input type= button class="button" value="���ε� �����ϱ�" onclick="delClick()">
</div>

<script language='javascript'>

	var totcnt = "<%= totcnt %>";
	var failcnt = "<%= failcnt %>";
	var failtype = "<%= failtype %>";

	if ((failcnt != "") && (failcnt != "0")) {
		alert("\n\n" + failcnt + "�� �Է½���!!\n\n");
	} else {
		/*
		alert(<%= totcnt %> + " �� ����Ǿ����ϴ�.");
		opener.location.reload();
		window.close();
		*/
	}

	<% if ErrStr <> "" then %>
		alert('<%= ErrStr %>');
	<% end if %>

	<% if (mode = "uploadorder") then %>
		<% if ErrStr = "" and failcnt = "0" then %>
			alert(totcnt + '��. ok');
		<% end if %>
	<% else %>

	<% end if %>

</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
