<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

''response.end

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, targetid, targetname, divcode, defaultmargine, itemgubunarr, itemidarr, itemoptionarr, itemnoarr
dim designer, sellcash, suplycash, buycash, mwdiv, itemname, itemoptionname
dim itemgubun, itemid, itemoption, itemno
dim pmwdiv
dim sqlStr, i

mode = request.form("mode")
targetid = request.form("brandid")

itemgubunarr = request.form("itemgubunarr")
itemidarr = request.form("itemidarr")
itemoptionarr = request.form("itemoptionarr")
itemnoarr = request.form("itemnoarr")
pmwdiv    = request.form("pmwdiv")

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnoarr = split(itemnoarr, "|")

dim iid, ipgocode
dim yyyymmdd

if (mode = "notused") then

elseif (mode = "lossarr") then
    '======================================================================
	'�ν� �����
	'1.�¶��� �԰� ����Ÿ
	targetid    = "itemloss"
	targetname  = "�ս�����"
	divcode     = "007"

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
	rsget("scheduledt") = Left(now(), 10)
	rsget("executedt") = Left(now(), 10)
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
			sqlStr = sqlStr + " values('" + ipgocode + "'," + itemid + ", '" + itemoption + "', " + sellcash + ", 0, " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
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
	sqlStr = sqlStr + " , suplycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " , buycash=(case when IsNULL(T.shopsuplycash,0)=0 then convert(int,T.shopitemprice*(100-d.defaultmargin)/100) else T.shopsuplycash end )" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + "     left join db_shop.dbo.tbl_shop_designer d on T.makerid=d.makerid and d.shopid='streetshop000'"
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
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
	'�ҷ���ǰ���Ӹ����� ������Ʈ -> ���ν����� �������..


	''���� �Է���
	sqlStr = "select convert(varchar(10),getdate(),21) as yyyymmdd "
	rsget.Open sqlStr,dbget,1
		yyyymmdd = rsget("yyyymmdd")
	rsget.Close

	''���Էµ� ������ �߰��� (���̺� ��ǰ�� ������ ���)
	sqlStr = " update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set errrealcheckno=errrealcheckno + IsNULL(T.itemno,0)*-1"
	sqlStr = sqlStr + " from ( "
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b, [db_summary].[dbo].tbl_erritem_daily_summary s"
	sqlStr = sqlStr + " where s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=T.itemoption"
	dbget.Execute sqlStr

	''���Էµ� ������ �߰��� (���̺� ��ǰ�� ���� ���)
	sqlStr = " insert into [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errrealcheckno,reguser)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " '" + yyyymmdd + "'"
	sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1,'" + session("ssBctId") + "'"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_erritem_daily_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " ) T"
	dbget.Execute sqlStr

	sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	''�Ϻ� ���α׿� �߰�
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary "
	sqlStr = sqlStr + " set errrealcheckno=errrealcheckno + b.itemno*-1"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "
	dbget.Execute sqlStr

	sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errrealcheckno)"
	sqlStr = sqlStr + " select "
	sqlStr = sqlStr + " '" + yyyymmdd + "'"
	sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s on s.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	sqlStr = sqlStr + " and b.itemid=s.itemid"
	sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	sqlStr = sqlStr + " where s.itemid is null and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " ) T"
	dbget.Execute sqlStr


	''���Ӹ�.
	sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	sqlStr = sqlStr + " set toterrno=errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno"
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + yyyymmdd + "'"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
	sqlStr = sqlStr + " and b.mastercode='" + CStr(ipgocode) + "' "

	dbget.Execute sqlStr


	''��������̺� ������Ʈ : �ý��� ��� ��ȭ?
	sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary "
	sqlStr = sqlStr + " set errrealcheckno=IsNULL(T.errrealcheckno,0) "
	sqlStr = sqlStr + " ,toterrno=IsNULL(T.errrealcheckno,0)+erretcno+errbaditemno "
	sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno+errcsno "
	sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+erretcno+errbaditemno "
	sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+IsNULL(T.errrealcheckno,0)+erretcno+errbaditemno "
	sqlStr = sqlStr + " ,shortageno=totipgono+totchulgono-totsellno+errcsno+IsNULL(T.errrealcheckno,0)+erretcno+requireno+ipkumdiv5+ipkumdiv4+ipkumdiv2+offconfirmno+offjupno+errbaditemno "
	sqlStr = sqlStr + " ,lastupdate=getdate()"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + "     select s.itemgubun, s.itemid, s.itemoption, sum(s.errrealcheckno) as errrealcheckno"
	sqlStr = sqlStr + "     from [db_summary].[dbo].tbl_erritem_daily_summary s,"
	sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_detail b "
	sqlStr = sqlStr + "     where s.itemgubun=b.iitemgubun"
	sqlStr = sqlStr + "     and s.itemid=b.itemid"
	sqlStr = sqlStr + "     and s.itemoption=b.itemoption"
	sqlStr = sqlStr + "     and b.mastercode='" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + "     group by s.itemgubun, s.itemid, s.itemoption"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"
	dbget.Execute sqlStr

	''����� ������Ʈ
	sqlStr = "exec db_summary.dbo.sp_Ten_recentIpChul_Update '" & ipgocode & "','','',0,'',''"
	dbget.Execute sqlStr
end if

%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->