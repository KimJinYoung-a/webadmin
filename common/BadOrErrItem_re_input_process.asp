<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%

''response.end

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, targetid, targetname, divcode, defaultmargine, itemgubunarr, itemidarr, itemoptionarr, itemnoarr
dim designer, sellcash, suplycash, buycash, mwdiv, itemname, itemoptionname
dim itemgubun, itemid, itemoption, itemno
dim pmwdiv
dim sqlStr, i

dim searchtype, actType, actyyyymmdd
dim chulgotargetid

dim datetype
dim yyyy, mm

mode = requestCheckVar(request.form("mode"),32)
targetid = requestCheckVar(request.form("brandid"),32)

itemgubunarr 	= request.form("itemgubunarr")
itemidarr 		= request.form("itemidarr")
itemoptionarr 	= request.form("itemoptionarr")
itemnoarr 		= request.form("itemnoarr")
pmwdiv    		= requestCheckVar(request.form("pmwdiv"),10)

searchtype 		= requestCheckVar(request.form("searchtype"),32)
actType    		= requestCheckVar(request.form("actType"),32)
actyyyymmdd   	= requestCheckVar(request.form("yyyymmdd"),32)
chulgotargetid  = requestCheckVar(request.form("chulgotargetid"),32)

datetype 		= requestCheckVar(request("datetype"),32)
yyyy 			= requestCheckVar(request("yyyy1"),4)
mm 				= requestCheckVar(request("mm1"),2)

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnoarr = split(itemnoarr, "|")

dim iid, ipgocode
dim yyyymmdd

if (searchtype = "bad") and (actType = "actreturn") then
	'======================================================================
	'�ҷ���ǰ+��ǰ�԰� ���
	'1.�¶��� �԰� ����Ÿ

	'��ü�� �˻� - ������� ������Ź��ǰ .
	sqlStr = " select top 1 socname_kor,maeipdiv,defaultmargine from [db_user].[dbo].tbl_user_c"
	sqlStr = sqlStr + " where userid='" + targetid + "'"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		targetname = rsget("socname_kor")
		divcode = rsget("maeipdiv")
		defaultmargine = rsget("defaultmargine")

        	if divcode="M" then
        		divcode = "001"
        	elseif divcode="W" then
        		divcode = "002"
        	end if
	end if
	rsget.close

    ''���Ա���.
    if (pmwdiv="M") then
        divcode = "001"
    elseif (pmwdiv="W") then
        divcode = "002"
    end if

	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = session("ssBctid")
	rsget("chargename") = session("ssBctCname")
	rsget("divcode") = divcode ''001-����, 002-��Ź
	rsget("vatcode") = "008"   ''�ΰ���.(�̰͸� �޴´�.)
	rsget("comment") = "�ҷ���ǰ��ǰó��"
	rsget("scheduledt") = actyyyymmdd
	rsget("executedt") = actyyyymmdd
	rsget("ipchulflag") = "I"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "ST" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr

	'''2.�¶��� �԰� ������ �Է�
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = CStr(-1 * trim(itemnoarr(i)))
			designer = targetid

			sellcash = "0"
			suplycash = "0"
			buycash = "0"
			mwdiv = "0"
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + requestCheckVar(itemid,10) + ", '" + requestCheckVar(itemoption,4) + "', " + sellcash + ", " + suplycash + ", " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + requestCheckVar(itemgubun,2) + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			dbget.Execute sqlStr
		end if
	next

    ''��ǰ���� - �¶���
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " , sellcash=[db_item].[dbo].tbl_item.sellcash"
	sqlStr = sqlStr + " , suplycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , buycash=[db_item].[dbo].tbl_item.buycash"

	'// 10 ��ü���, 90 ��ǰ => ���͸��Ա��� �̿�, 2015-04-15, skyer9
	''sqlStr = sqlStr + " , mwgubun=[db_item].[dbo].tbl_item.mwdiv"
	sqlStr = sqlStr + " , mwgubun='" + CStr(pmwdiv) + "'"

	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	dbget.Execute sqlStr

    ''�ɼǸ� - �¶���
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	dbget.Execute sqlStr

	''�������� ��ǰ��, �ɼ�
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,iitemoptionname=IsNULL(T.shopitemoptionname,'')" + vbCrlf
	sqlStr = sqlStr + " , sellcash=T.shopitemprice" + vbCrlf
	sqlStr = sqlStr + " , suplycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " , buycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf

	'// 10 ��ü���, 90 ��ǰ => ���͸��Ա��� �̿�, 2015-04-15, skyer9
	sqlStr = sqlStr + " , mwgubun = T.centermwdiv" + vbCrlf

	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_designer d on T.makerid=d.makerid and d.shopid='streetshop000'"
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
	dbget.Execute sqlStr

	sqlStr = " update D "
	sqlStr = sqlStr + " set buycash=isNULL(s.lastbuyprice,buycash) "
	sqlStr = sqlStr + " from  db_storage.dbo.tbl_acount_storage_detail D "
	sqlStr = sqlStr + " 	Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	on S.yyyymm='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " where D.mastercode='" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'======================================================================
	sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_Update] '" + CStr(ipgocode) + "', '" + CStr(searchtype) + "', '" + CStr(actyyyymmdd) + "', '" + CStr(session("ssBctId")) + "' "
	dbget.Execute sqlStr

	sqlStr = " db_summary.[dbo].[sp_Ten_RealtimeStock_IpChulUpdateByIpChulCode] '" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

elseif (searchtype = "bad") and (actType = "actloss") then

	divcode = request.form("divcode")

    '======================================================================
	'�ҷ���ǰ+�ν���� ���
	'1.�¶��� �԰� ����Ÿ
	if (chulgotargetid = "") then
		targetid = "itemloss"
	else
		targetid = chulgotargetid
	end if

	if (targetid = "itemloss") then
		targetname  = "�ս�����"
	elseif (targetid = "itemstockmodify") then
		targetname  = "�����"
	elseif (targetid = "3pl_its_loss") then
		targetname  = "�ս�����"
	elseif (targetid = "itemoutlet") then
		targetname  = "�ƿ﷿"
	else
		targetname  = targetid
	end if

	if (divcode = "") then
		divcode = "007"
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = session("ssBctid")
	rsget("chargename") = session("ssBctCname")
	rsget("divcode") = divcode ''001-����, 002-��Ź
	rsget("vatcode") = "008"   ''�ΰ���.(�̰͸� �޴´�.)
	rsget("comment") = "�ҷ���ǰ�ν�ó��"
	rsget("scheduledt") = actyyyymmdd
	rsget("executedt") = actyyyymmdd
	rsget("ipchulflag") = "E"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "SO" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'''2.�¶��� ��� ������ �Է� ��� 0
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = CStr(-1 * trim(itemnoarr(i)))
			designer = targetid

			sellcash = "0"
			suplycash = "0"
			buycash = "0"
			mwdiv = "0"
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + requestCheckVar(itemid,10) + ", '" + requestCheckVar(itemoption,4) + "', " + sellcash + ", 0, " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + requestCheckVar(itemgubun,2) + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			dbget.Execute sqlStr
		end if
	next

    ''��ǰ����
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " , sellcash=[db_item].[dbo].tbl_item.sellcash"
	sqlStr = sqlStr + " , suplycash=0"
	sqlStr = sqlStr + " , buycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , mwgubun=[db_item].[dbo].tbl_item.mwdiv"
	sqlStr = sqlStr + " , imakerid=[db_item].[dbo].tbl_item.makerid"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	dbget.Execute sqlStr

    ''�ɼǸ� - �¶���
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	dbget.Execute sqlStr

	''�������� ��ǰ��, �ɼ�
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,iitemoptionname=IsNULL(T.shopitemoptionname,'')" + vbCrlf
	sqlStr = sqlStr + " , sellcash=T.shopitemprice" + vbCrlf
	sqlStr = sqlStr + " , suplycash=0" + vbCrlf
	sqlStr = sqlStr + " , buycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " , imakerid=T.makerid"  + vbCrlf ''2014/07/29 �߰�
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_designer d on T.makerid=d.makerid and d.shopid='streetshop000'"
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
	dbget.Execute sqlStr

    '' �̺κ��� ����. lastbuyprice?? 2016/02/03  WITH (INDEX(IX_tbl_acount_storage_detail_mastercode)) �߰�
	sqlStr = " update D "
	sqlStr = sqlStr + " set buycash=isNULL(s.lastbuyprice,buycash) "
	sqlStr = sqlStr + " from  db_storage.dbo.tbl_acount_storage_detail D"	'  WITH (INDEX(IX_tbl_acount_storage_detail_mastercode))
	sqlStr = sqlStr + " 	Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	on S.yyyymm='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " where D.mastercode='" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'======================================================================
	sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_Update] '" + CStr(ipgocode) + "', '" + CStr(searchtype) + "', '" + CStr(actyyyymmdd) + "', '" + CStr(session("ssBctId")) + "' "
	dbget.Execute sqlStr

	sqlStr = " db_summary.[dbo].[sp_Ten_RealtimeStock_IpChulUpdateByIpChulCode] '" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

''����  WITH (INDEX(IX_tbl_acount_storage_detail_mastercode �߰� 2016/02/03
	sqlStr = " update S "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	etcchulgono=S.etcchulgono+D.itemno "
	sqlStr = sqlStr + " 	,totChulgono=S.totChulgono+D.itemno "
	sqlStr = sqlStr + " 	,errrealcheckno=S.errrealcheckno-D.itemno "
	sqlStr = sqlStr + " 	,totErrno=S.totErrno-D.itemno "
	sqlStr = sqlStr + " 	,totSysstock=S.totSysstock+D.itemno "
	sqlStr = sqlStr + " from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	Join db_storage.dbo.tbl_acount_storage_detail D" '' WITH (INDEX(IX_tbl_acount_storage_detail_mastercode, IX_tbl_acount_storage_detail_itemid))
	sqlStr = sqlStr + " 	on S.yyyymm>='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " 	and D.mastercode='" + CStr(ipgocode) + "' "
	''dbget.Execute sqlStr

    sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_After_accStock_Update] '" + CStr(ipgocode) + "','"+CStr(Left(actyyyymmdd, 7))+"' "
	dbget.Execute sqlStr

elseif (searchtype = "bad") and (actType = "actshopchulgo") then

	divcode = request.form("divcode")

    '======================================================================
	'�ҷ���ǰ+������� ���
	'1.�¶��� �԰� ����Ÿ
	if (chulgotargetid = "") then
		targetid = "streetshop900"
	else
		targetid = chulgotargetid
	end if

	targetname  = targetid

	if (divcode = "") then
		divcode = "007"
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = session("ssBctid")
	rsget("chargename") = session("ssBctCname")
	rsget("divcode") = divcode
	rsget("vatcode") = "008"
	rsget("comment") = "�ҷ���ǰ �������"
	rsget("scheduledt") = actyyyymmdd
	rsget("executedt") = actyyyymmdd
	rsget("ipchulflag") = "S"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "SO" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'''2.�¶��� ��� ������ �Է� ��� 0
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = CStr(-1 * trim(itemnoarr(i)))
			designer = targetid

			sellcash = "0"
			suplycash = "0"
			buycash = "0"
			mwdiv = "0"
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + requestCheckVar(itemid,10) + ", '" + requestCheckVar(itemoption,4) + "', " + sellcash + ", 0, " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + requestCheckVar(itemgubun,2) + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			dbget.Execute sqlStr
		end if
	next

    ''��ǰ����
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " , sellcash=[db_item].[dbo].tbl_item.sellcash"
	sqlStr = sqlStr + " , suplycash=0"
	sqlStr = sqlStr + " , buycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , mwgubun=[db_item].[dbo].tbl_item.mwdiv"
	sqlStr = sqlStr + " , imakerid=[db_item].[dbo].tbl_item.makerid"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	dbget.Execute sqlStr

    ''�ɼǸ� - �¶���
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	dbget.Execute sqlStr

	''�������� ��ǰ��, �ɼ�
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,iitemoptionname=IsNULL(T.shopitemoptionname,'')" + vbCrlf
	sqlStr = sqlStr + " , sellcash=T.shopitemprice" + vbCrlf
	sqlStr = sqlStr + " , suplycash=0" + vbCrlf
	sqlStr = sqlStr + " , buycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " , imakerid=T.makerid"  + vbCrlf ''2014/07/29 �߰�
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_designer d on T.makerid=d.makerid and d.shopid='streetshop000'"
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
	dbget.Execute sqlStr

	sqlStr = " update D "
	sqlStr = sqlStr + " set buycash=isNULL(s.lastbuyprice,buycash) "
	sqlStr = sqlStr + " from  db_storage.dbo.tbl_acount_storage_detail D"	'   WITH (INDEX (IX_tbl_acount_storage_detail_mastercode))
	sqlStr = sqlStr + " 	Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	on S.yyyymm='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " where D.mastercode='" + CStr(ipgocode) + "' "
	'// ����, ����, skyer9, 2016-03-16
	''dbget.Execute sqlStr

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'======================================================================
	sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_Update] '" + CStr(ipgocode) + "', '" + CStr(searchtype) + "', '" + CStr(actyyyymmdd) + "', '" + CStr(session("ssBctId")) + "' "
	dbget.Execute sqlStr

	'// ���� ���
	sqlStr = " db_summary.[dbo].[sp_Ten_RealtimeStock_IpChulUpdateByIpChulCode] '" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	'// ���� ���
	sqlStr = " db_summary.[dbo].[sp_Ten_RealtimeStock_IpChulUpdateSHOPByIpChulCode] '" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

''����
	sqlStr = " update S "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	etcchulgono=S.etcchulgono+D.itemno "
	sqlStr = sqlStr + " 	,totChulgono=S.totChulgono+D.itemno "
	sqlStr = sqlStr + " 	,errrealcheckno=S.errrealcheckno-D.itemno "
	sqlStr = sqlStr + " 	,totErrno=S.totErrno-D.itemno "
	sqlStr = sqlStr + " 	,totSysstock=S.totSysstock+D.itemno "
	sqlStr = sqlStr + " from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	Join db_storage.dbo.tbl_acount_storage_detail D " '' WITH (INDEX(IX_tbl_acount_storage_detail_mastercode, IX_tbl_acount_storage_detail_itemid))
	sqlStr = sqlStr + " 	on S.yyyymm>='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " 	and D.mastercode='" + CStr(ipgocode) + "' "
	''dbget.Execute sqlStr

	''�������
    sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_shopChulgo_After_accStock_Update] '" + CStr(ipgocode) + "','"+CStr(Left(actyyyymmdd, 7))+"' "
	dbget.Execute sqlStr

elseif (searchtype = "err") and (actType = "actloss") then

	divcode = request.form("divcode")

    '======================================================================
	'�ν� �����
	'1.�¶��� �԰� ����Ÿ
	if (chulgotargetid = "") then
		targetid = "itemloss"
	else
		targetid = chulgotargetid
	end if

	if (targetid = "itemloss") then
		targetname  = "�ս�����"
	elseif (targetid = "itemstockmodify") then
		targetname  = "�����"
	else
		targetname  = targetid
	end if

	if (divcode = "") then
		divcode = "007"
	end if

	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = session("ssBctid")
	rsget("chargename") = session("ssBctCname")
	rsget("divcode") = divcode ''001-����, 002-��Ź
	rsget("vatcode") = "008"   ''�ΰ���.(�̰͸� �޴´�.)
	rsget("comment") = "������ǰ�ν�ó��"
	rsget("scheduledt") = actyyyymmdd
	rsget("executedt") = actyyyymmdd
	rsget("ipchulflag") = "E"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "SO" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr


	'''2.�¶��� ��� ������ �Է� ��� 0
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			itemid = trim(itemidarr(i))
			itemoption = trim(itemoptionarr(i))
			itemno = CStr(-1 * trim(itemnoarr(i)))
			designer = targetid

			sellcash = "0"
			suplycash = "0"
			buycash = "0"
			mwdiv = "0"
			itemname = ""
			itemoptionname = ""

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + requestCheckVar(itemid,10) + ", '" + requestCheckVar(itemoption,4) + "', " + sellcash + ", 0, " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + requestCheckVar(itemgubun,2) + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			dbget.Execute sqlStr
		end if
	next

    ''��ǰ����
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " , sellcash=[db_item].[dbo].tbl_item.sellcash"
	sqlStr = sqlStr + " , suplycash=0"
	sqlStr = sqlStr + " , buycash=[db_item].[dbo].tbl_item.buycash"
	sqlStr = sqlStr + " , mwgubun=[db_item].[dbo].tbl_item.mwdiv"
	sqlStr = sqlStr + " , imakerid=[db_item].[dbo].tbl_item.makerid"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	dbget.Execute sqlStr

    ''�ɼǸ� - �¶���
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	dbget.Execute sqlStr

	''�������� ��ǰ��, �ɼ�
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,iitemoptionname=IsNULL(T.shopitemoptionname,'')" + vbCrlf
	sqlStr = sqlStr + " , imakerid=T.makerid"
	sqlStr = sqlStr + " , sellcash=T.shopitemprice" + vbCrlf
	sqlStr = sqlStr + " , suplycash=0" + vbCrlf
	sqlStr = sqlStr + " , buycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_designer d on T.makerid=d.makerid and d.shopid='streetshop000'"
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
	dbget.Execute sqlStr

	sqlStr = " update D "
	sqlStr = sqlStr + " set buycash=isNULL(s.lastbuyprice,buycash) "
	sqlStr = sqlStr + " from  db_storage.dbo.tbl_acount_storage_detail D"  ''WITH (INDEX (IX_tbl_acount_storage_detail_mastercode)) 2015/06/01	'  WITH (INDEX (IX_tbl_acount_storage_detail_mastercode))
	sqlStr = sqlStr + " 	Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	on S.yyyymm='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " where D.mastercode='" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	dbget.Execute sqlStr

rw now()
	'======================================================================
	sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_Update] '" + CStr(ipgocode) + "', '" + CStr(searchtype) + "', '" + CStr(actyyyymmdd) + "', '" + CStr(session("ssBctId")) + "' "
	dbget.Execute sqlStr
rw now()
	''sqlStr = " db_summary.[dbo].[sp_Ten_RealtimeStock_IpChulUpdateByIpChulCode] '" + CStr(ipgocode) + "' "
	''dbget.Execute sqlStr

	'// ����� ������Ʈ
	sqlStr = " update S " + VBCrlf
	sqlStr = sqlStr + " set " + VBCrlf
	sqlStr = sqlStr + " 	etcchulgono=S.etcchulgono+D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totChulgono=S.totChulgono+D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,errrealcheckno=S.errrealcheckno-D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totErrno=S.totErrno-D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totSysstock=S.totSysstock+D.itemno " + VBCrlf
	sqlStr = sqlStr + " from [db_summary].[dbo].[tbl_current_logisstock_summary] S " + VBCrlf
	sqlStr = sqlStr + " 	Join db_storage.dbo.tbl_acount_storage_detail D " + VBCrlf
	sqlStr = sqlStr + " 	on 1 = 1 " + VBCrlf
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun " + VBCrlf
	sqlStr = sqlStr + " 	and D.itemid=S.itemid " + VBCrlf
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption " + VBCrlf
	sqlStr = sqlStr + " 	and D.mastercode='"  + CStr(ipgocode) + "' " + VBCrlf
	dbget.Execute sqlStr

	'// �Ϻ���� ������ �Է�
	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary " + VBCrlf
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption) " + VBCrlf
	sqlStr = sqlStr + " select convert(varchar(10), M.executedt, 121), D.iitemgubun, D.itemid, D.itemoption " + VBCrlf
	sqlStr = sqlStr + " from " + VBCrlf
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_acount_storage_detail D " + VBCrlf
	sqlStr = sqlStr + " 	 join [db_storage].[dbo].tbl_acount_storage_master M " + VBCrlf
	sqlStr = sqlStr + " 	 on " + VBCrlf
	sqlStr = sqlStr + " 		M.code = D.mastercode " + VBCrlf
	sqlStr = sqlStr + " 	Left Join [db_summary].[dbo].[tbl_daily_logisstock_summary] S " + VBCrlf
	sqlStr = sqlStr + " 	on " + VBCrlf
	sqlStr = sqlStr + " 	 	1 = 1 " + VBCrlf
	sqlStr = sqlStr + " 		and D.iitemgubun=S.itemgubun " + VBCrlf
	sqlStr = sqlStr + " 		and D.itemid=S.itemid " + VBCrlf
	sqlStr = sqlStr + " 		and D.itemoption=S.itemoption " + VBCrlf
	sqlStr = sqlStr + " 		and S.yyyymmdd = convert(varchar(10), M.executedt, 121) " + VBCrlf
	sqlStr = sqlStr + " where " + VBCrlf
	sqlStr = sqlStr + " 	1 = 1 " + VBCrlf
	sqlStr = sqlStr + " 	and D.mastercode='"  + CStr(ipgocode) + "' " + VBCrlf
	sqlStr = sqlStr + " 	and S.yyyymmdd is NULL " + VBCrlf
	dbget.Execute sqlStr

	'// �Ϻ���� ������Ʈ
	sqlStr = " update S " + VBCrlf
	sqlStr = sqlStr + " set " + VBCrlf
	sqlStr = sqlStr + " 	etcchulgono=S.etcchulgono+D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totChulgono=S.totChulgono+D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,errrealcheckno=S.errrealcheckno-D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totErrno=S.totErrno-D.itemno " + VBCrlf
	sqlStr = sqlStr + " 	,totSysstock=S.totSysstock+D.itemno " + VBCrlf
	sqlStr = sqlStr + " from " + VBCrlf
	sqlStr = sqlStr + " 	db_storage.dbo.tbl_acount_storage_detail D " + VBCrlf
	sqlStr = sqlStr + " 	Join [db_summary].[dbo].[tbl_daily_logisstock_summary] S " + VBCrlf
	sqlStr = sqlStr + " 	on " + VBCrlf
	sqlStr = sqlStr + " 	 	1 = 1 " + VBCrlf
	sqlStr = sqlStr + " 		and D.iitemgubun=S.itemgubun " + VBCrlf
	sqlStr = sqlStr + " 		and D.itemid=S.itemid " + VBCrlf
	sqlStr = sqlStr + " 		and D.itemoption=S.itemoption " + VBCrlf
	sqlStr = sqlStr + " 		and D.mastercode='"  + CStr(ipgocode) + "' " + VBCrlf
	sqlStr = sqlStr + " 	 join [db_storage].[dbo].tbl_acount_storage_master M " + VBCrlf
	sqlStr = sqlStr + " 	 on " + VBCrlf
	sqlStr = sqlStr + " 		1 = 1 " + VBCrlf
	sqlStr = sqlStr + " 		and M.code = D.mastercode " + VBCrlf
	sqlStr = sqlStr + " 		and S.yyyymmdd = convert(varchar(10), M.executedt, 121) " + VBCrlf
	dbget.Execute sqlStr

rw now()
''��� �ڻ� ������Ʈ
	sqlStr = " update S "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	etcchulgono=S.etcchulgono+D.itemno "
	sqlStr = sqlStr + " 	,totChulgono=S.totChulgono+D.itemno "
	sqlStr = sqlStr + " 	,errrealcheckno=S.errrealcheckno-D.itemno "
	sqlStr = sqlStr + " 	,totErrno=S.totErrno-D.itemno "
	sqlStr = sqlStr + " 	,totSysstock=S.totSysstock+D.itemno "
	sqlStr = sqlStr + " from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary S "
	sqlStr = sqlStr + " 	Join db_storage.dbo.tbl_acount_storage_detail D  " ''WITH (INDEX(IX_tbl_acount_storage_detail_mastercode, IX_tbl_acount_storage_detail_itemid))
	sqlStr = sqlStr + " 	on S.yyyymm>='" + CStr(Left(actyyyymmdd, 7)) + "' "
	sqlStr = sqlStr + " 	and D.iitemgubun=S.itemgubun "
	sqlStr = sqlStr + " 	and D.itemid=S.itemid "
	sqlStr = sqlStr + " 	and D.itemoption=S.itemoption "
	sqlStr = sqlStr + " 	and D.mastercode='" + CStr(ipgocode) + "' "
	''dbget.Execute sqlStr

	sqlStr = " exec db_summary.dbo.[usp_Ten_BadErrProc_After_accStock_Update] '" + CStr(ipgocode) + "','"+CStr(Left(actyyyymmdd, 7))+"' "
	dbget.Execute sqlStr
rw now()
end if

%>

<script type='text/javascript'>
	alert('���� �Ǿ����ϴ�.');
	location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
