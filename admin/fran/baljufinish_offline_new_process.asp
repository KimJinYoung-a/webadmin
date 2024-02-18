<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ��� ó��
' Hieditor : 2011.03.09 �̻� ����
'			 2012.08.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/logisticsbaljuofflinecls.asp"-->
<%

dim lastPageTime, pageElapsedTime
lastPageTime = Timer

'// Call checkAndWriteElapsedTime("001")
function checkAndWriteElapsedTime(memo)
	pageElapsedTime = Timer - lastPageTime
	lastPageTime = Timer
	response.write "<!-- Page Execute Time Check : " & FormatNumber(pageElapsedTime, 4) & " : " & memo & " -->" & vbCrLf
end function

dim mode , baljuid, baljudate, itemgubun, itemid, itemoption, comment
dim i,cnt,sqlStr, errstring ,masteridxlist, baljuname, baljucodelist ,songjangdiv
dim divcode, vatinclude, targetid, targetname, baljucode, brandlist, obaljucode
dim siteseq, companyid ,masteridx, ordercode ,isorgorder, orgordercode ,remasteridx, reordercode
dim itemexists, iid, newbaljucode, itemAlreadyExists, tmp
dim currencyUnit, IsFinished, isWait, loginsite
Dim baljuKey, siteBaljuKey
Dim ordercodelistNoBeasongDate, ordercodelistNoSiteInsertDate

	mode        = RequestCheckVar(request("mode"),32)
	baljukey    = RequestCheckVar(request("baljunum"),32)
	baljuid     = RequestCheckVar(request("baljuid"),32)
	itemgubun   = RequestCheckVar(request("itemgubun"),200)
	itemid      = RequestCheckVar(request("itemid"),840)
	itemoption  = RequestCheckVar(request("itemoption"),440)
	comment     = RequestCheckVar(request("comment"),1280)
	isWait      = RequestCheckVar(request("isWait"),32)


dim IsWriteReOrderSheet : IsWriteReOrderSheet = False			'// ���ֹ��� �ۼ�
dim IsWriteChulgoSheet : IsWriteChulgoSheet = False				'// ����� �ۼ�
dim errMsg :errMsg = ""

sqlStr = " select IsFinished, siteBaljuid as siteBaljukey, songjangdiv " + VbCrLf
sqlStr = sqlStr + " from db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
sqlStr = sqlStr + " where baljuKey = " & baljuKey & " " + VbCrLf
'response.write sqlStr & "<Br>"
rsget_Logistics.Open sqlStr,dbget_Logistics,1

if  not rsget_Logistics.EOF  then
	IsFinished 		= rsget_Logistics("IsFinished")
	siteBaljuKey 	= rsget_Logistics("siteBaljuKey")
	songjangdiv 	= rsget_Logistics("songjangdiv")
end if
rsget_Logistics.Close

Select Case IsFinished
	Case "N"
		'// ����۾���
		IsWriteReOrderSheet = True
		if (isWait = "N") then
			IsWriteChulgoSheet = True
		end if
	Case "W"
		'// �����
		if (isWait = "N") then
			IsWriteChulgoSheet = True
		else
			errMsg = "���� : ����۾��� ���°� �ƴմϴ�."
		end if
	Case "Y"
		'// ����
		errMsg = "���� : �̹� ���Ϸ�� �����Դϴ�."
	Case Else
		errMsg = "���� : �� �� ���� ����"
End Select

if (errMsg <> "") then
	response.write errMsg
	dbget_Logistics.Close
	dbget.Close
	response.end
end If


companyid = session("ssBctID")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

siteseq = GetLogicsSiteSeq		'/lib/classes/order/logisticsbaljuofflinecls.asp

function GetFromWhere(siteseq, baljuKey, baljuid)
	dim tmpsql

    tmpsql = " FROM " + VbCrLf
    tmpsql = tmpsql + " 	db_aLogistics.dbo.tbl_Logistics_offline_baljumaster bm " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_baljudetail b " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_order_master m " + VbCrLf
    tmpsql = tmpsql + " 	, [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail d " + VbCrLf
    tmpsql = tmpsql + " 	LEFT JOIN [db_aLogistics].[dbo].tbl_Logistics_offline_item i " + VbCrLf
    tmpsql = tmpsql + " 	ON " + VbCrLf
    tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 		and d.siteseq = i.siteseq " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemgubun = i.siteitemgubun " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemid = i.siteitemid " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemoption = i.siteitemoption " + VbCrLf
    tmpsql = tmpsql + " 	LEFT JOIN [db_aLogistics].[dbo].tbl_Logistics_offline_tmppacking p " + VbCrLf
    tmpsql = tmpsql + " 	ON " + VbCrLf
    tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 		and d.siteseq = p.siteseq " + VbCrLf
    tmpsql = tmpsql + " 		and d.ordercode = p.ordercode " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemgubun = p.itemgubun " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemid = p.itemid " + VbCrLf
    tmpsql = tmpsql + " 		and d.itemoption = p.itemoption " + VbCrLf
    tmpsql = tmpsql + " WHERE " + VbCrLf
    tmpsql = tmpsql + " 	1 = 1 " + VbCrLf
    tmpsql = tmpsql + " 	and bm.baljukey = b.baljukey " + VbCrLf
    tmpsql = tmpsql + " 	and m.siteseq = d.siteseq " + VbCrLf
    tmpsql = tmpsql + " 	and m.ordercode = d.ordercode " + VbCrLf
    tmpsql = tmpsql + " 	and b.ordercode = m.ordercode " + VbCrLf
    tmpsql = tmpsql + " 	and d.cancelyn <> 'Y' " + VbCrLf
    tmpsql = tmpsql + " 	and m.SiteSeq = " & siteseq & " " + VbCrLf
    tmpsql = tmpsql + "     and b.baljukey = '" + CStr(baljuKey) + "' " + VbCrLf
    tmpsql = tmpsql + " 	and m.shopid = '" & baljuid & "' " + VbCrLf

    GetFromWhere = tmpsql
end Function

Call checkAndWriteElapsedTime("001")
''dbget.close() : response.end

if mode="chulgoproc" Then

	'// ========================================================================
    '����üũ : �߸��� �Է�(�ڽ���ȣ�� 0 �̸鼭 �����ȣ�� �ִ°�� or realitemno �� �����鼭, �ڽ���ȣ�� ���°��)üũ
    sqlStr = " select d.itemname,d.itemoptionname " + VbCrLf
    sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid)
    sqlStr = sqlStr + " 	and m.beasongdate is null " + VbCrLf
    sqlStr = sqlStr + " 	and (((isnull(d.packingstate,0) = 0) and (isnull(d.songjangno,'0') <> '0')) or ((d.fixedno > 0) and (isnull(d.songjangno,'0') = '0'))) " + VbCrLf

	'response.write sqlStr & "<Br>"
    rsget_Logistics.Open sqlStr, dbget_Logistics, 1
    if  not rsget_Logistics.EOF  then
        do until rsget_Logistics.eof
            if (trim(errstring) = "") then
				errstring = rsget_Logistics("itemname") + "(" + rsget_Logistics("itemoptionname") + ")"
            else
                errstring = errstring + ", " + rsget_Logistics("itemname") + "(" + rsget_Logistics("itemoptionname") + ")"
            end if

            rsget_Logistics.MoveNext
        loop
    else
		errstring = ""
    end if
    rsget_Logistics.close

    if (errstring <> "") then
        response.write "<script>alert('�߸��� �Է��� �ֽ��ϴ�. �����ȣ �Ǵ� ��ǰ�ڵ带 ������ �ٽ� �Է��ϼ���.\n\n" + errstring + "');</script>"
        response.write "<script>history.back();</script>"

        dbget.close()
        dbget_Logistics.close()
        response.End
    end If


	'// ========================================================================
	ordercodelistNoSiteInsertDate = ""
	ordercodelistNoBeasongDate = ""

	'// siteinsertdate �� ���� �ƴϸ� ���ֹ��� �ۼ��Ϸ�
	sqlStr = " select distinct d.ordercode " + VbCrLf
	sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid) + VbCrLf
	sqlStr = sqlStr + " 	and m.siteinsertdate is null " + VbCrLf
	'response.write sqlStr & "<Br>"
	rsget_Logistics.Open sqlStr,dbget_Logistics,1

	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
			if (ordercodelistNoSiteInsertDate <> "") then
				ordercodelistNoSiteInsertDate = ordercodelistNoSiteInsertDate + ",'" + CStr(rsget_Logistics("ordercode")) & "'"
			else
				ordercodelistNoSiteInsertDate = "'" & CStr(rsget_Logistics("ordercode")) & "'"
			end if

			rsget_Logistics.MoveNext
		loop
	end if
	rsget_Logistics.close

	'// beasongdate �� ���� �ƴϸ� ���Ϸ�
	sqlStr = " select distinct d.ordercode "
	sqlStr = sqlStr + GetFromWhere(siteseq, baljuKey, baljuid)
	sqlStr = sqlStr + " 	and m.beasongdate is null " + VbCrLf
	'response.write sqlStr & "<Br>"
	rsget_Logistics.Open sqlStr,dbget_Logistics,1

	if  not rsget_Logistics.EOF  then
		do until rsget_Logistics.eof
			if (ordercodelistNoBeasongDate <> "") then
				ordercodelistNoBeasongDate = ordercodelistNoBeasongDate + ",'" + CStr(rsget_Logistics("ordercode")) & "'"
			else
				ordercodelistNoBeasongDate = "'" & CStr(rsget_Logistics("ordercode")) & "'"
			end if

			rsget_Logistics.MoveNext
		loop
	end if
	rsget_Logistics.Close

	'�ش� ��������ڵ�/������þ��̵� ���� masteridx �� ���Ѵ�.
	sqlStr = " select distinct d.masteridx "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_shopbalju b "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		b.baljucode = m.baljucode "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.idx = d.masteridx "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	b.baljunum = " & siteBaljuKey
	sqlStr = sqlStr + " 	and m.deldt is null "
	sqlStr = sqlStr + " 	and d.deldt is null "
	sqlStr = sqlStr + " 	and m.divcode in ('501','502','503') "
	sqlStr = sqlStr + " 	and m.statecd < '7' "
	'response.write sqlStr & "<Br>"
	rsget.Open sqlStr,dbget,1

	masteridxlist = ""
	if  not rsget.EOF  then
		do until rsget.eof
			if (masteridxlist <> "") then
				masteridxlist = masteridxlist + "," + CStr(rsget("masteridx"))
			else
				masteridxlist = CStr(rsget("masteridx"))
			end if

			rsget.MoveNext
		loop
	end if
	rsget.Close

	IF (masteridxlist="") then masteridxlist="-1"       ''2011-07-04 �߰�

	Call checkAndWriteElapsedTime("002")
	''dbget.close() : response.end

	'// ========================================================================
	'// ������ �۾�����(�̹�� ����, ���� ��ǰ��, �������� ��) �Է� �� ����
	if IsWriteReOrderSheet And (ordercodelistNoSiteInsertDate <> "") Then
		'������ �̹�� �ֹ���ǰ�� ���� �ڸ�Ʈ �Է�
		if (Trim(comment) <> "") then
			itemgubun = split(itemgubun,"|")
			itemid = split(itemid,"|")
			itemoption = split(itemoption,"|")
			comment = split(comment,"|")
			cnt = ubound(itemgubun)

			for i=0 to cnt
				if (Trim(comment(i)) <> "") then
					sqlstr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail "
					sqlstr = sqlstr + " set comment = '" + Trim(comment(i)) + "' "
					sqlstr = sqlstr + " where itemgubun = '" + Trim(itemgubun(i)) + "' "
					sqlstr = sqlstr + " and itemid = " + Trim(itemid(i)) + " "
					sqlstr = sqlstr + " and itemoption = '" + Trim(itemoption(i)) + "' "
					sqlstr = sqlstr + " and siteseq = " & siteseq & " "
					sqlstr = sqlstr + " and ordercode in (" + CStr(ordercodelistNoSiteInsertDate) + ") "
					rsget_Logistics.Open sqlStr, dbget_Logistics, 1
				end if
			next
		end If

		Call checkAndWriteElapsedTime("003")
		''dbget.close() : response.end

		if (CStr(siteseq) = "10") Then
			'// ���ο� �۾����� ����
			sqlStr = " exec [db_storage].[dbo].[usp_Ten_LogicsOffChulgo2SCM] " & baljuKey
			dbget.Execute sqlStr

			Call checkAndWriteElapsedTime("004")
			''dbget.close() : response.end

			sqlStr = " update "
			sqlStr = sqlStr + " 	td "
			sqlStr = sqlStr + " set "
			sqlStr = sqlStr + " 	td.realitemno = ld.fixedno "
			sqlStr = sqlStr + " 	, td.packingstate = ld.packingstate "
			sqlStr = sqlStr + " 	, td.boxsongjangno = ld.songjangno "
			sqlStr = sqlStr + " 	, td.comment = ld.comment "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_storage].[dbo].[tbl_Logistics_offline_order_detail_COPY] ld "
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail td "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		ld.sitedetailidx = td.idx "
			sqlStr = sqlStr + " 		and ld.baljuKey = " & baljuKey
			sqlstr = sqlstr + " 		and ld.ordercode in (" + CStr(ordercodelistNoSiteInsertDate) + ") "
			''response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1

			Call checkAndWriteElapsedTime("005")
			''dbget.close() : response.end

			'���� ���������� ������Ʈ
			sqlStr = " update m " + vbCrLf
			sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totforeign_sellcash,0)" + vbCrLf
			sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totforeign_suplycash,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realforeign_sellcash,0)" + vbCrLf
			sqlStr = sqlStr + " ,totalforeign_suplycash	=IsNULL(T.realforeign_suplycash,0)" + vbCrLf
			sqlStr = sqlStr + " from " + vbCrLf
			sqlStr = sqlStr + " 	( " + vbCrLf
			sqlStr = sqlStr + " 		select m.baljucode, sum(sellcash*baljuitemno) as totsell " + vbCrLf
			sqlStr = sqlStr + " 		,sum(suplycash*baljuitemno) as totsupp " + vbCrLf
			sqlStr = sqlStr + " 		,sum(buycash*baljuitemno) as totbuy " + vbCrLf
			sqlStr = sqlStr + " 		,sum(sellcash*realitemno) as realtotsell " + vbCrLf
			sqlStr = sqlStr + " 		,sum(suplycash*realitemno) as realtotsupp " + vbCrLf
			sqlStr = sqlStr + " 		,sum(buycash*realitemno) as realtotbuy " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_sellcash,0)*baljuitemno) as totforeign_sellcash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_suplycash,0)*baljuitemno) as totforeign_suplycash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_sellcash,0)*realitemno) as realforeign_sellcash " + vbCrLf
			sqlStr = sqlStr + " 		,sum(IsNull(foreign_suplycash,0)*realitemno) as realforeign_suplycash " + vbCrLf
			sqlStr = sqlStr + " 		from " + vbCrLf
			sqlStr = sqlStr + " 			[db_storage].[dbo].tbl_ordersheet_master m " + vbCrLf
			sqlStr = sqlStr + " 			join [db_storage].[dbo].tbl_ordersheet_detail d " + vbCrLf
			sqlStr = sqlStr + " 			on " + vbCrLf
			sqlStr = sqlStr + " 				m.idx = d.masteridx " + vbCrLf
			sqlStr = sqlStr + " 		where " + vbCrLf
			sqlStr = sqlStr + " 			1 = 1 " + vbCrLf
			sqlStr = sqlStr + " 			and m.baljucode in (" + CStr(ordercodelistNoSiteInsertDate) + ") " + vbCrLf
			sqlStr = sqlStr + " 			and d.deldt is null" + vbCrLf
			sqlStr = sqlStr + " 		group by " + vbCrLf
			sqlStr = sqlStr + " 			m.baljucode " + vbCrLf
			sqlStr = sqlStr + " 	) as T" + vbCrLf
			sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m " + vbCrLf
			sqlStr = sqlStr + " 	on " + vbCrLf
			sqlStr = sqlStr + " 		m.baljucode = T.baljucode " + vbCrLf
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1


			'TEN-5. �̹���ֹ� ���� üũ(���ֹ� ��� ��ǰ �˻�)
			sqlStr = " select count(d.idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail d "
			sqlStr = sqlStr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
			sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
			sqlStr = sqlStr + " and d.comment='5�ϳ����' "
			sqlStr = sqlStr + " and deldt is null "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
    			itemexists = (rsget("cnt")>0)
			rsget.Close

			sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
			sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
			sqlStr = sqlStr + " and clinkcode  is not null "
			sqlStr = sqlStr + " and clinkcode<>'' "

			Call checkAndWriteElapsedTime("006")
			''dbget.close() : response.end

			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
				itemAlreadyExists = (rsget("cnt")>0)
			rsget.Close

			if Not itemexists then
				'response.write "<script>alert('�� �ֹ��� ������ �����ϴ�.');</script>"
			elseif itemAlreadyExists then
				'response.write "<script>alert('�� �ֹ����� �̹� �ۼ��Ǿ� �ֽ��ϴ�. �ۼ��� �� �����ϴ�.');</script>"
			Else
				'�̹�� �ֹ��� �ۼ�
				'�������� �ֹ����� �ϳ��� ���´�.
				sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
				sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "

				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
					targetid = rsget("targetid")
					targetname = rsget("targetname")
					divcode = rsget("divcode")
					vatinclude = rsget("vatinclude")
					currencyUnit = rsget("currencyUnit")
					loginsite = rsget("sitename")
				rsget.Close

				'�ش� ��������ڵ�/������þ��̵� ���� �⺻������ ���Ѵ�.
				sqlStr = " select distinct m.baljuname, m.baljucode "
				sqlStr = sqlStr + " from "
				sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_shopbalju b "
				sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_master m "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		b.baljucode = m.baljucode "
				sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_ordersheet_detail d "
				sqlStr = sqlStr + " 	on "
				sqlStr = sqlStr + " 		m.idx = d.masteridx "
				sqlStr = sqlStr + " where "
				sqlStr = sqlStr + " 	b.baljunum = " & siteBaljuKey
				sqlStr = sqlStr + " 	and m.deldt is null "
				sqlStr = sqlStr + " 	and d.deldt is null "
				sqlStr = sqlStr + " 	and m.divcode in ('501','502','503') "
				sqlStr = sqlStr + " 	and m.statecd < '7' "
				sqlStr = sqlStr + " 	and d.baljuitemno <> d.realitemno "
				sqlStr = sqlStr + " 	and d.comment='5�ϳ����' "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr,dbget,1

				baljuname = ""
				baljucode = ""
				baljucodelist = ""
				if  not rsget.EOF  then
					baljuname = CStr(rsget("baljuname"))
					baljucode = CStr(rsget("baljucode"))
					baljucodelist = CStr(rsget("baljucode"))

					rsget.MoveNext
					do until rsget.eof
                        baljucodelist = baljucodelist + "," + CStr(rsget("baljucode"))
                        rsget.MoveNext
					loop
				end if
				rsget.Close

				sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0 "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr,dbget,1,3
				rsget.AddNew
				rsget("targetid") = targetid
				rsget("targetname") = targetname
				rsget("baljuid") = baljuid
				rsget("baljuname") = baljuname

				if loginsite = "WSLWEB"	 then
					rsget("currencyUnit") = currencyUnit
					rsget("foreign_statecd") = "0"
					rsget("sitename") = loginsite
				end if

				rsget("reguser") = session("ssBctId")
				rsget("regname") = session("ssBctCname")
				rsget("divcode") = divcode
				rsget("vatinclude") = vatinclude
				rsget("scheduledate") = Left(now(), 10)
				rsget("statecd") = "0"
				rsget("comment") = baljucodelist + " �̹�۰� ���ۼ�"

				rsget.update
        			iid = rsget("idx")
				rsget.close

				baljucode = "RJ" + Format00(6,Right(CStr(iid),6))

				''������ ����
				sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
				sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash," + vbCrlf
				sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
				sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
				sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash," + vbCrlf
				sqlStr = sqlStr + " sum(baljuitemno-realitemno),sum(baljuitemno-realitemno),baljudiv" + vbCrlf
				sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " where masteridx in (" + CStr(masteridxlist) + ") "
				sqlStr = sqlStr + " and baljuitemno <> realitemno "
				sqlStr = sqlStr + " and comment='5�ϳ����'"
				sqlStr = sqlStr + " and deldt is null"
				sqlStr = sqlStr + " group by itemgubun,makerid,itemid,itemoption,itemname,itemoptionname,sellcash,suplycash,buycash,foreign_sellcash, foreign_suplycash,baljudiv "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1

				''���Ӹ� ����
				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
				sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunforeign_sellcash=IsNULL(T.totsellforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,jumunforeign_suplycash=IsNULL(T.totsuppforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalforeign_sellcash=IsNULL(T.realtotsellforeign,0)" + vbCrlf
				sqlStr = sqlStr + " ,totalforeign_suplycash=IsNULL(T.realtotsuppforeign,0)" + vbCrlf
				sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
				sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
				sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
				sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
				sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
				sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_sellcash * baljuitemno) as totsellforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_suplycash * baljuitemno) as totsuppforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_sellcash * realitemno) as realtotsellforeign " + vbCrlf
				sqlStr = sqlStr + " ,sum(foreign_suplycash * realitemno) as realtotsuppforeign " + vbCrlf
				sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
				sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
				sqlStr = sqlStr + " and deldt is null" + vbCrlf
				sqlStr = sqlStr + " ) as T" + vbCrlf
				sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1


				''�귣�� ����Ʈ
				brandlist = ""
				sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
				sqlStr = sqlStr + " where masteridx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
        		do until rsget.eof
        			brandlist = brandlist + rsget("makerid") + ","
        			rsget.movenext
        		loop
				rsget.close

				if brandlist<>"" then
					brandlist = Left(brandlist,Len(brandlist)-1)
					brandlist = Left(brandlist,255)
				end if

				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
				sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
				sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
				sqlStr = sqlStr + " where idx=" + CStr(iid)
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1


				''��������ü��� ��ũ�ڵ� ����.
				sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
				sqlStr = sqlStr + " set clinkcode='" + baljucode + "'" + VbCrlf
				sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
				'response.write sqlStr & "<Br>"
				rsget.Open sqlStr, dbget, 1
			End If

			Call checkAndWriteElapsedTime("007")
			''dbget.close() : response.end

			'ī���ڽ�����
			sqlstr = " delete d "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " db_storage.dbo.tbl_cartoonbox_detail d "
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " 	join [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox c " + VbCrLf
			else
				sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox c " + VbCrLf
			end if
			sqlStr = sqlStr + " on " + VbCrLf
			sqlStr = sqlStr + " 	1 = 1 " + VbCrLf
			sqlStr = sqlStr + " 	and d.baljudate = c.baljudate " + VbCrLf
			sqlStr = sqlStr + " 	and d.shopid = c.shopid " + VbCrLf
			sqlStr = sqlStr + " 	and d.innerboxno = c.innerboxno " + VbCrLf
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			sqlstr = " insert into db_storage.dbo.tbl_cartoonbox_detail(" + VbCrLf
			sqlStr = sqlStr + " baljudate" + VbCrLf
			sqlStr = sqlStr + " ,shopid" + VbCrLf
			sqlStr = sqlStr + " ,cartoonboxno" + VbCrLf
			sqlStr = sqlStr + " ,cartoonboxweight" + VbCrLf
			sqlStr = sqlStr + " ,cartonboxsongjangdiv" + VbCrLf
			sqlStr = sqlStr + " ,cartonboxsongjangno" + VbCrLf
			sqlStr = sqlStr + " ,innerboxno" + VbCrLf
			sqlStr = sqlStr + " ,innerboxweight" + VbCrLf
			sqlStr = sqlStr + " ,innerboxidx" + VbCrLf
			sqlStr = sqlStr + " ) " + VbCrLf
			sqlStr = sqlStr + " 	select"
			sqlStr = sqlStr + " 	baljudate"
			sqlStr = sqlStr + " 	, shopid"
			sqlStr = sqlStr + " 	, cartoonboxno"
			sqlStr = sqlStr + " 	, 0"
			sqlStr = sqlStr + " 	, cartoonboxsongjangdiv"
			sqlStr = sqlStr + " 	, cartoonboxsongjangno"
			sqlStr = sqlStr + " 	, innerboxno"
			sqlStr = sqlStr + " 	, innerboxweight"
			sqlStr = sqlStr + " 	, ctIDX"
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " 	from [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			else
				sqlStr = sqlStr + " 	from [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			end if
			sqlStr = sqlStr + " where siteseq = " + CStr(siteseq) + " "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			sqlstr = " delete from "
			if application("Svr_Info") <> "Dev" then
				sqlStr = sqlStr + " [LOGISTICSDB].[db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			else
				sqlStr = sqlStr + " [db_aLogistics].[dbo].tbl_Logistics_offline_cartoonbox " + VbCrLf
			end if
			sqlStr = sqlStr + " where siteseq = " + CStr(siteseq) + " "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr


			''�ֹ����� : �����
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
			sqlStr = sqlStr + " set statecd='6'" + VbCrlf
			sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1
		End If


		'������ MASTER ���� - �Է¿Ϸ�ǥ��
		sqlStr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_master " + vbCrlf
		sqlStr = sqlStr + " set siteinsertdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where siteseq = " & siteseq & " and ordercode in (" + ordercodelistNoSiteInsertDate + ")  "
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	End If

	Call checkAndWriteElapsedTime("008")
	''dbget.close() : response.end

	'// ========================================================================
	'// �ֹ� ������ ��������
	if IsWriteChulgoSheet And ordercodelistNoBeasongDate <> ""Then

		if (CStr(siteseq) = "10") Then
			'�� �ֹ��ڵ庰 ó��(�����Ÿ ���� ��)
			tmp = split(masteridxlist,",")
			for i=0 to UBound(tmp)
				if (Trim(tmp(i)) <> "") Then
					'Ȯ�������� ���� �հ�ݾ� ���
					sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
					sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
					sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
					sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
					sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
					sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
					sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
					sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
					sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
					sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
					sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
					sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
					sqlStr = sqlStr + " where masteridx="  + CStr(Trim(tmp(i))) + vbCrlf
					sqlStr = sqlStr + " and deldt is null" + vbCrlf
					sqlStr = sqlStr + " ) as T" + vbCrlf
					sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(Trim(tmp(i)))
					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr, dbget, 1

					'�ش� ��������ڵ�/������þ��̵� ���� �⺻������ ���Ѵ�.
					sqlStr = " select distinct m.baljuname, m.baljucode "
					sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
					sqlStr = sqlStr + " where 1 = 1 "
					sqlStr = sqlStr + " and m.idx = d.masteridx "
					sqlStr = sqlStr + " and m.deldt is null "
					sqlStr = sqlStr + " and d.deldt is null "
					sqlStr = sqlStr + " and b.baljucode = m.baljucode "
					sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
					sqlStr = sqlStr + " and m.statecd < '7' "
					sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
					sqlStr = sqlStr + " and b.baljunum = " + CStr(siteBaljuKey) + " "
					sqlStr = sqlStr + " and m.idx = " + CStr(Trim(tmp(i))) + " "
					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr,dbget,1

					baljuname = ""
					baljucode = ""
					if  not rsget.EOF  then
						baljuname = CStr(rsget("baljuname"))
						baljucode = CStr(rsget("baljucode"))
					end if
					rsget.Close

					''��� ����Ÿ�� �Է�. *-1
					sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d "
					sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
					sqlStr = sqlStr + " and d.deldt is null "
					sqlStr = sqlStr + " and d.realitemno <> 0 "

					'response.write sqlStr & "<Br>"
					rsget.Open sqlStr, dbget, 1
                    	itemexists = rsget("cnt")>0
					rsget.close

					if itemexists Then
						'1.�¶��� ��� ����Ÿ
						sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0 "

						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1,3
						rsget.AddNew
						rsget("code") = ""
						rsget("socid") = baljuid
						rsget("socname") = baljuname
						rsget("chargeid") = session("ssBctId")
						rsget("divcode") = "006"
						rsget("vatcode") = "008"
						rsget("comment") = baljucode + " �ֹ� ������� �� �ڵ����ó��"
						rsget("chargename") = session("ssBctCname")
						rsget("ipchulflag") = "S"

						rsget.update
						iid = rsget("id")
						rsget.close

						newbaljucode = "SO" + Format00(6,Right(CStr(iid),6))

						sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
						sqlStr = sqlStr + " set code='" + newbaljucode + "'" + VBCrlf
						sqlStr = sqlStr + " where id=" + CStr(iid)
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'2.�¶��� ��� ������ �Է�
						sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail "
						sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,itemno, "
						sqlStr = sqlStr + " buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) "
						sqlStr = sqlStr + " select '" + newbaljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash, "
						sqlStr = sqlStr + " sum(d.realitemno*-1) as itemno, d.buycash,d.ipgoflag,d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
						sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d "
						sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
						sqlStr = sqlStr + " and deldt is null "
						sqlStr = sqlStr + " and d.realitemno<>0 "
						sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.sellcash, d.suplycash, d.buycash,d.ipgoflag, "
						sqlStr = sqlStr + " d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'3.�¶��� ��� ����Ÿ ������Ʈ
						sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
						sqlStr = sqlStr + " set executedt='" + Left(now(), 10) + "'" + VBCrlf
						sqlStr = sqlStr + " ,scheduledt='" + Left(now(), 10) + "'" + VBCrlf
						sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totsell,0)" + VBCrlf
						sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + VBCrlf
						sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totbuy,0)" + VBCrlf
						sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
						sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
						sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
						sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
						sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
						sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
						sqlStr = sqlStr + " where mastercode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " and deldt is null" + vbCrlf
						sqlStr = sqlStr + " ) as T"
						sqlStr = sqlStr + " where id=" + CStr(iid)
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr,dbget,1

						'4.���º���
						sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
						sqlStr = sqlStr + " set statecd='7'" + vbCrlf
						sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
						sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr, dbget, 1

						'' ��/��� ���ݿ�
						sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & newbaljucode & "','','',0,'',''"
						'response.write sqlStr & "<Br>"
						dbget.Execute sqlStr

						'// ������� �ݿ�
                        if (baljukey <> 74851) then
						    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & baljuid & "', '" & newbaljucode & "' "
						    'response.write sqlStr & "<Br>"
						    dbget.Execute sqlStr
                        end if
					else
						'// ������� ��ǰ���� ��� : �״�� ���Ϸ�� ��ȯ

						'4.���º���
						sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
						sqlStr = sqlStr + " set statecd='7'" + vbCrlf
						''sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
						''sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
						sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
						'response.write sqlStr & "<Br>"
						rsget.Open sqlStr, dbget, 1
					End If
				End If
			Next
		End If

		Call checkAndWriteElapsedTime("009")
		''dbget.close() : response.end

		'������ �⺻ MASTER ���� ����
		sqlStr = " update [db_aLogistics].[dbo].tbl_Logistics_offline_order_master " + vbCrlf
		sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where siteseq = " & siteseq & " and ordercode in (" + ordercodelistNoBeasongDate + ")  "
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1

		if (CStr(siteseq) = "10") Then
			'���� �⺻ MASTER ���� ����
			sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
			sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
			sqlStr = sqlStr + " where baljucode in (" + ordercodelistNoBeasongDate + ") "
			'response.write sqlStr & "<Br>"
			rsget.Open sqlStr, dbget, 1

			'// ������� �ݿ�(�����)
			sqlStr = " update d "
			sqlStr = sqlStr + " set d.shopReceive = 'N' "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_cartoonbox_detail d "
			sqlStr = sqlStr + " join ( "
			sqlStr = sqlStr + " 	select distinct b.baljuid as shopid, DATEADD(dd, DATEDIFF(dd, 0, b.baljudate), 0) as baljudate, d.packingstate as innerboxno "
			sqlStr = sqlStr + " 	from "
			sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_shopbalju b "
			sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_ordersheet_master m "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			b.baljucode = m.baljucode "
			sqlStr = sqlStr + " 		join [db_storage].[dbo].tbl_ordersheet_detail d "
			sqlStr = sqlStr + " 		on "
			sqlStr = sqlStr + " 			m.idx = d.masteridx "
			sqlStr = sqlStr + " 	where "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and b.baljunum = " & siteBaljuKey
			sqlStr = sqlStr + " ) T "
			sqlStr = sqlStr + " on "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and d.shopid = T.shopid "
			sqlStr = sqlStr + " 	and d.baljudate = T.baljudate "
			sqlStr = sqlStr + " 	and d.innerboxno = T.innerboxno "
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr

			'' ���� �������� ����
			''sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
            sqlstr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_offjupsuAll_bySiteBaljuKey] " & siteBaljuKey
			'response.write sqlStr & "<Br>"
			dbget.Execute sqlStr


			''siteBaljukey   ''���� ��� ��ǰ ���� �� ���� ���� �缳��.. 2011-06 �߰�.
			if (siteBaljukey<>0) then
				sqlstr = " exec [db_summary].[dbo].[sp_Ten_RealtimeStock_itemLimitByOffChulgo] "&siteBaljukey
				'response.write sqlStr & "<Br>"
				dbget.Execute sqlStr
			end If


			'' ���� �������� ���� 2011-08 �߰�
			'���ֹ� �������� �ʴ´�.(����������� �ǹ̰� ����.)
			'// ��ü ���꺸�ٴ� ������û�ǰ��ϸ� ������Ʈ�ϴ� ������ ���� �ʿ�
			'sqlstr = " exec [db_summary].dbo.[sp_Ten_Shop_Stock_PreOrderUpdate_ALL]"
			'dbget.Execute sqlStr
		End If
	End If

	Call checkAndWriteElapsedTime("010")
	''dbget.close() : response.end

	'// ========================================================================
	'// ������� ������ ��������
	if (isWait = "Y") then
		'������ ������ø����� IsFinished="W" �Է�
		sqlStr = " update db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
		sqlStr = sqlStr + " set IsFinished = 'W'  " + VbCrLf
		sqlStr = sqlStr + " where baljuKey = " & baljuKey & " and siteseq = " & siteseq & " " + VbCrLf
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	elseif IsWriteChulgoSheet then
		'������ ������ø����� IsFinished="Y" �Է�
		'������ ���������� beasongdate �Է�
		sqlStr = " update db_aLogistics.dbo.tbl_Logistics_offline_baljumaster " + VbCrLf
		sqlStr = sqlStr + " set IsFinished = 'Y'  " + VbCrLf
		sqlStr = sqlStr + " where baljuKey = " & baljuKey & " and siteseq = " & siteseq & " " + VbCrLf
		'response.write sqlStr & "<Br>"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	end If


	If (CStr(siteseq) <> "10") Then
		'3PL
		dim STOCK_GUID
		response.write "<script>alert('���� - ó������.');</script>"
		response.End
	End If
end if

Call checkAndWriteElapsedTime("011")
''dbget.close() : response.end

%>

<script language="javascript">
	alert('���� �Ǿ����ϴ�.');
	location.replace('baljulist_offline_new.asp');
</script>

<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
