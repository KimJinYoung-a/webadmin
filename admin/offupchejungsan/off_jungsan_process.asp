<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim mode, gubuncd
dim yyyy, mm, yyyymm, masteridx, makerid
dim taxlinkidx, neotaxno, billsiteCode, eseroEvalSeq
dim ipchulArr

dim ispreFixTaxDateForce, preFixedTaxDate
dim isrefPay, refPayreqIdx, payRealDate, PayreqIdxValid
dim iCheExists,ipfileNo, jgubun, itemvatYN, commission
dim jacctcd

mode    = request("mode")
gubuncd = request("gubuncd")
yyyy    = request("yyyy")
mm      = request("mm")
masteridx  = request("masteridx")
makerid    = request("makerid")
taxlinkidx = requestCheckVar(request("taxlinkidx"),10)
neotaxno   = requestCheckVar(request("neotaxno"),32)
billsiteCode = requestCheckVar(request("billsiteCode"),10)
eseroEvalSeq = requestCheckVar(Trim(replace(request("eseroEvalSeq"),"-","")),24)
ipchulArr = request("ipchulArr")
jacctcd = Trim(requestCheckVar(request("jacctcd"),10))

ispreFixTaxDateForce = request("ispreFixTaxDateForce")
preFixedTaxDate      = request("preFixedTaxDate")
isrefPay             = request("isrefPay")
refPayreqIdx         = request("refPayreqIdx")

jgubun      = request("jgubun") ''2014/02
itemvatYN   = request("itemvatYN") ''2014/02
commission  = request("commission") ''2014/02

if (ipchulArr<>"") then
    ipchulArr = "'"&replace(replace(trim(ipchulArr)," ",""),",","','")&"'"

end if
yyyymm = yyyy + "-" + mm

dim startYYYYMMDD, nextYYYYMMDD
startYYYYMMDD = CStr(dateserial(yyyy,mm,1))
nextYYYYMMDD = CStr(dateserial(yyyy,mm+1,1))



'response.write mode
'response.write "<br>"
'response.write gubuncd
'response.write "<br>"

dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim sqlStr,i
dim chargediv, differencekey, taxtype, titleStr

differencekey = request("differencekey")
taxtype = request("taxtype")


dim taxregdate, ipkumdate, comment
dim groupid, availneoport
taxregdate = request("taxregdate")
ipkumdate = request("ipkumdate")
comment   = request("comment")

groupid = request("groupid")
availneoport = request("availneoport")

if (availneoport="on") then
    availneoport="1"
else
    availneoport="0"
end if

dim shopid
dim itemgubun, itemid, itemoption
dim itemname, itemoptionname
dim sellprice, suplyprice, itemno


shopid          = request("shopid")
itemgubun       = request("itemgubun")
itemid          = request("itemid")
itemoption      = request("itemoption")
itemname        = html2db(request("itemname"))
itemoptionname  = html2db(request("itemoptionname"))
sellprice       = request("sellprice")
suplyprice      = request("suplyprice")
itemno          = request("itemno")

dim detailidx, idxarr
detailidx       = request("detailidx")
idxarr          = request("idxarr")

dim IsDataExists
dim AssignedCount, AssignedRow

function MakeDefaultJungsanMaster(yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid)
    dim sqlStr
    dim startYYYYMMDD, nextYYYYMMDD, PreStartYYYYMMDD

    startYYYYMMDD = CStr(dateserial(yyyy,mm,1))
    nextYYYYMMDD = CStr(dateserial(yyyy,mm+1,1))
    PreStartYYYYMMDD = LEFT(DATEADD("m",-2,startYYYYMMDD),10)


    if (gubuncd="B011") or (gubuncd="B012") then
    ''��Ź(��ü/����)�Ǹ��� ���
        '' ���� �ۼ��� ���� 2�� ''streetshop095 �ʰ� �ø���CASE ����
        ''IF (makerid<>"") then
        ''    startYYYYMMDD = LEFT(DATEADD("m",-2,startYYYYMMDD),10)
        ''END IF

        sqlStr = " DECLARE @TMPTABLE TABLE ( "&VbCrlf
	    sqlStr = sqlStr + " SHOPID	varchar(32) "&VbCrlf
        sqlStr = sqlStr + " ,	MAKERID	varchar(32) "&VbCrlf
        sqlStr = sqlStr + " ) "&VbCrlf

        sqlStr = sqlStr + " insert into @TMPTABLE "&VbCrlf
        sqlStr = sqlStr + "     	select distinct m.shopid, d.makerid"
        sqlStr = sqlStr + "     	from [db_shop].[dbo].tbl_shopjumun_master m,"
        sqlStr = sqlStr + "     	[db_shop].[dbo].tbl_shopjumun_detail d"
        sqlStr = sqlStr + " 			left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " 			on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " 			and d.itemid=s.shopitemid"
		sqlStr = sqlStr + " 			and d.itemoption=s.itemoption"
        sqlStr = sqlStr + "     	where m.orderno=d.orderno"
        sqlStr = sqlStr + "     	and d.jcomm_cd='"&gubuncd&"'"
        IF (makerid<>"") THEN
            sqlStr = sqlStr + "     	and ((m.shopregdate>='" + startYYYYMMDD + "' and m.shopregdate<'" + nextYYYYMMDD + "')"
            sqlStr = sqlStr + "     	    or (m.shopid='streetshop095' and m.shopregdate>='" + PreStartYYYYMMDD + "' and m.shopregdate<'" + nextYYYYMMDD + "'))"
        ELSE
            sqlStr = sqlStr + "     	and m.shopregdate>='" + startYYYYMMDD + "'"
            sqlStr = sqlStr + "     	and m.shopregdate<'" + nextYYYYMMDD + "'"
        END IF

        sqlStr = sqlStr + "     	and m.cancelyn='N'"
        sqlStr = sqlStr + "     	and d.cancelyn='N'"
        if makerid<>"" then
            sqlStr = sqlStr + " and d.makerid='" + makerid + "'"
        end if
        if taxtype="01" then
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')='N'" &VbCrlf
        end if

        ''sqlStr = sqlStr + " 		and IsNULL(s.centermwdiv,'W')='W'" &VbCrlf
        sqlStr = sqlStr + " ;" &VbCrlf
        sqlStr = sqlStr + " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
        sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

        sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', T.makerid,"
        sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
        sqlStr = sqlStr + " from @TMPTABLE T "

        '' �̹��� ���곻���� ���� BrandID
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
        sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
        if makerid<>"" then
            sqlStr = sqlStr + " and j.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and T.makerid=j.makerid"

        '' Groupid ����
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + "     on T.makerid=p.id"


        sqlStr = sqlStr + "     where j.makerid is null"

        ''''sqlStr = sqlStr + "     and s.autojungsan ='Y'"
'response.write sqlStr
'response.end

        rsget.Open sqlStr,dbget,1

    elseif gubuncd="B031" then
    ''��� ���� - ��Ź ��ǰ�� ���ؼ���. shopitem's centermwdiv(���Ϳ��� �������� �޴��� ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
        sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

        sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', s.makerid,"
        sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer s"

        '' �̹��� ���곻���� ���� BrandID
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
        sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
        if makerid<>"" then
            sqlStr = sqlStr + " and j.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and s.makerid=j.makerid"

        '' Groupid ����
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + "     on s.makerid=p.id"

        '' �̹��� ������� �ִ� BrandID
        sqlStr = sqlStr + "     left join ("
        sqlStr = sqlStr + " 		select distinct m.socid, d.imakerid as makerid from "
        sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 			left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " 			on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + " 			and d.itemid=s.shopitemid"
		sqlStr = sqlStr + " 			and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 		where m.executedt>='" + startYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.ipchulflag='S'"
		sqlStr = sqlStr + " 		and m.code=d.mastercode"
		sqlStr = sqlStr + " 		and m.deldt is NULL"
		sqlStr = sqlStr + " 		and d.deldt is NULL"
		sqlStr = sqlStr + " 		and d.itemno<>0"
		if makerid<>"" then
            sqlStr = sqlStr + " and d.imakerid='" + makerid + "'"
        end if
		if taxtype="01" then
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')='N'"
        end if

        '' ���Ա����� ���� �ȵ��ִ°�� �������� �ν�
		''sqlStr = sqlStr + " 		and IsNULL(s.centermwdiv,'M')='W'"
		sqlStr = sqlStr + " 		and IsNULL(d.mwgubun,'M')='C'"  ''������� //2014/01/27
        sqlStr = sqlStr + "     ) T  "
        sqlStr = sqlStr + "     on s.shopid=T.socid"
        sqlStr = sqlStr + "     and s.makerid=T.makerid"

        sqlStr = sqlStr + "     where s.chargediv in (" + chargediv + ")"
        sqlStr = sqlStr + "     and j.makerid is null"
        sqlStr = sqlStr + "     and T.makerid is not null"

        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B021" then
    ''���� ���� (���� �������� ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
        sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

        sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', s.makerid,"
        sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer s"

        '' �̹��� ���곻���� ���� BrandID
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
        sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
        if makerid<>"" then
            sqlStr = sqlStr + " and j.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and s.makerid=j.makerid"

        '' Groupid ����
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + "     on s.makerid=p.id"

        '' �̹��� ���Գ�����  �ִ� BrandID
        sqlStr = sqlStr + "     left join ("

        sqlStr = sqlStr + " 		select distinct d.imakerid as makerid from "
        sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_master m,"
		sqlStr = sqlStr + " 		[db_storage].[dbo].tbl_acount_storage_detail d"
		sqlStr = sqlStr + " 			left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " 			on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + " 			and d.itemid=s.shopitemid"
		sqlStr = sqlStr + " 			and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 		where m.executedt>='" + startYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.ipchulflag='I'"
		sqlStr = sqlStr + " 		and m.divcode='801'"
		sqlStr = sqlStr + " 		and m.code=d.mastercode"
		if makerid<>"" then
            sqlStr = sqlStr + "     and d.imakerid='" + makerid + "'"
        end if
		sqlStr = sqlStr + " 		and m.deldt is NULL"
		sqlStr = sqlStr + " 		and d.deldt is NULL"
		sqlStr = sqlStr + " 		and d.itemno<>0"

		if taxtype="01" then
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')='N'"
        end if

        '' ���Ա����� ���� �ȵ��ִ°�� �������� �ν� **
		sqlStr = sqlStr + " 		and IsNULL(s.centermwdiv,'M')='M'"
        sqlStr = sqlStr + "     ) T  "
        sqlStr = sqlStr + "     on s.makerid=T.makerid"

        sqlStr = sqlStr + "     where s.chargediv in (" + chargediv + ")"
        sqlStr = sqlStr + "     and j.makerid is null"
        sqlStr = sqlStr + "     and T.makerid is not null"

        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B022" then
    ''���� ���� (���� ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
        sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

        sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', s.makerid,"
        sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer s"

        '' �̹��� ���곻���� ���� BrandID
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
        sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
        if makerid<>"" then
            sqlStr = sqlStr + " and j.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and  s.makerid=j.makerid"

        '' Groupid ����
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + "     on s.makerid=p.id"

        '' �̹��� ���� ���Գ�����  �ִ� BrandID
        sqlStr = sqlStr + "     left join ("

        sqlStr = sqlStr + " 		select distinct m.shopid, d.designerid as makerid from "
        sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_ipchul_master m,"
		sqlStr = sqlStr + " 		[db_shop].[dbo].tbl_shop_ipchul_detail d"
		sqlStr = sqlStr + " 			left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " 			on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " 			and d.shopitemid=s.shopitemid"
		sqlStr = sqlStr + " 			and d.itemoption=s.itemoption"

		sqlStr = sqlStr + " 		where m.execdt>='" + startYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.execdt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " 		and m.idx=d.masteridx"
		if makerid<>"" then
            sqlStr = sqlStr + "     and d.designerid='" + makerid + "'"
        end if
		sqlStr = sqlStr + " 		and m.statecd>=7"
		sqlStr = sqlStr + " 		and m.deleteyn='N'"
		sqlStr = sqlStr + " 		and d.deleteyn='N'"
		sqlStr = sqlStr + " 		and d.itemno<>0"
		sqlStr = sqlStr + " 		and m.comm_cd='B022'"                           '''2012/06/01�߰�

		if taxtype="01" then
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')='N'"
        end if

        '' ���Ա����� ���� �ȵ��ִ°�� �������� �ν� ** : ����.. �߸������Ǿ��ִ°�찡 ����.
		'''sqlStr = sqlStr + " 		and IsNULL(s.centermwdiv,'M')='M'"
        sqlStr = sqlStr + "     ) T  "
        sqlStr = sqlStr + "     on s.makerid=T.makerid"

        sqlStr = sqlStr + "     where s.chargediv ='" + chargediv + "'"
        sqlStr = sqlStr + "     and j.makerid is null"
        sqlStr = sqlStr + "     and T.makerid is not null"

        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B077" then
    ''��ü���
        gubuncd = "B012"  ''' ��ü��Ź �ʿ� ����.

        sqlStr = " DECLARE @TMPTABLE TABLE ( "&VbCrlf
	    sqlStr = sqlStr + " SHOPID	varchar(32) "&VbCrlf
        sqlStr = sqlStr + " ,	MAKERID	varchar(32) "&VbCrlf
        sqlStr = sqlStr + " ) "&VbCrlf

        sqlStr = sqlStr + " insert into @TMPTABLE "&VbCrlf
        sqlStr = sqlStr + "     	select distinct m.shopid, d.makerid"
        sqlStr = sqlStr + "     	from db_shop.dbo.tbl_shopbeasong_order_master m,"
        sqlStr = sqlStr + "     	[db_shop].[dbo].tbl_shopbeasong_order_detail d"
        sqlStr = sqlStr + " 			left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + " 			on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + " 			and d.itemid=s.shopitemid"
		sqlStr = sqlStr + " 			and d.itemoption=s.itemoption"
        sqlStr = sqlStr + "     	where m.masteridx=d.masteridx"
        sqlStr = sqlStr + "     	and m.ipkumdiv>3"
        sqlStr = sqlStr + "     	and d.beasongdate>='" + startYYYYMMDD + "'"
        sqlStr = sqlStr + "     	and d.beasongdate<'" + nextYYYYMMDD + "'"
        sqlStr = sqlStr + "     	and m.cancelyn='N'"
        sqlStr = sqlStr + "     	and d.cancelyn='N'"
        sqlStr = sqlStr + "     	and d.isupchebeasong='Y'"

        if makerid<>"" then
            sqlStr = sqlStr + " and d.makerid='" + makerid + "'"
        end if
        if taxtype="01" then
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + "     	and IsNULL(s.vatinclude,'Y')='N'" &VbCrlf
        end if


        ''sqlStr = sqlStr + " 		and IsNULL(s.centermwdiv,'W')='W'" &VbCrlf
        sqlStr = sqlStr + " ;" &VbCrlf
        sqlStr = sqlStr + " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
        sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

        sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', s.makerid,"
        sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer s"

        '' �̹��� ���곻���� ���� BrandID
        sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
        sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
        if makerid<>"" then
            sqlStr = sqlStr + " and j.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and s.makerid=j.makerid"

        '' Groupid ����
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + "     on s.makerid=p.id"

        '' �̹��� �Ǹų����� �ִ� BrandID
        sqlStr = sqlStr + "     left join @TMPTABLE T  "
        sqlStr = sqlStr + "         on s.shopid=T.shopid"
        sqlStr = sqlStr + "         and s.makerid=T.makerid"

        sqlStr = sqlStr + "     where s.chargediv='" + chargediv + "'"
        sqlStr = sqlStr + "     and j.makerid is null"
        sqlStr = sqlStr + "     and T.makerid is not null"

'response.write sqlStr
'response.end

        rsget.Open sqlStr,dbget,1
    end if
end function

function MakeDefaultJungsanDetail(yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid,ipchulArr)
    dim sqlStr
    dim startYYYYMMDD, nextYYYYMMDD, PreStartYYYYMMDD

    startYYYYMMDD = CStr(dateserial(yyyy,mm,1))
    nextYYYYMMDD = CStr(dateserial(yyyy,mm+1,1))
    PreStartYYYYMMDD = LEFT(DATEADD("m",-2,startYYYYMMDD),10)

    if (gubuncd="B011") or (gubuncd="B012") or (gubuncd="B013") then
    ''��Ź(����,��ü)
        ''���԰�ó��.
        '' ���� �ۼ��� ���� 2�� ''streetshop095 �ʰ� �ø���CASE ����
        ''IF (makerid<>"") then
        ''    startYYYYMMDD = LEFT(DATEADD("m",-2,startYYYYMMDD),10)
        ''END IF

        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail"
        sqlStr = sqlStr + " (masteridx,shopid,gubuncd,orderno,"
        sqlStr = sqlStr + " itemgubun,itemid,itemoption,itemname,itemoptionname,"
        sqlStr = sqlStr + " sellprice,realsellprice,suplyprice,itemno,makerid,linkidx)"
        sqlStr = sqlStr + " select  J.idx as masteridx, m.shopid, '" + gubuncd + "', m.orderno, "
        sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname,"
        sqlStr = sqlStr + " d.sellprice,d.realsellprice,"
        ''���԰�(���갡)
        sqlStr = sqlStr + " ( case "
        if gubuncd="B011" then
        ''������ ��� ���Ի�ǰ ó�� (���԰� 0)
            sqlStr = sqlStr + "     when IsNULL(s.centermwdiv,'M')<>'W' then 0"
        end if

        ''���� ����...(2011-03-01 : �ֹ��� ������ ���԰��� ����) => ���� ��� �̰����� ����.
        ''IF (makerid<>"") then
            sqlStr = sqlStr + "  	when d.suplyprice<>0 then d.suplyprice"
        ''END IF

        ''���ǸŰ�(���� ��� ����) - ������.(2011-03-01)
        '''sqlStr = sqlStr + "		when (n.autojungsandiv='R') then convert(int,d.realsellprice - d.realsellprice*n.defaultmargin/100)"
		''������ ���԰��� ������� : �⺻��������, d.discountprice : ���ν� �����.

		''���� ���� �Ǹ��̰�, �������� �ǸŵȰ��.
		sqlStr = sqlStr + "  	when (s.shopsuplycash=0) and (s.discountsellprice<>0) and (d.discountprice<>0) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"

		''�׿� �����Ǹ��� ��� �Һ� ��� %
		sqlStr = sqlStr + "		when (s.shopsuplycash=0) then convert(int,d.sellprice - d.sellprice*n.defaultmargin/100)"

		''(s.shopsuplycash<>0)
		sqlStr = sqlStr + "		when (d.sellprice<>s.shopitemprice) and (d.realsellprice<>s.discountsellprice) then convert(int,d.realsellprice - d.realsellprice*n.defaultmargin/100)"
		''sqlStr = sqlStr + "		when (s.discountsellprice<>0) and (d.discountprice<>s.discountsellprice) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"

		'''sqlStr = sqlStr + "  	when (s.shopsuplycash=0) and (s.discountprice<>0) and (s.discountprice<>d.discountprice) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"
		''������ ���԰��� �������.
		sqlStr = sqlStr + "    else s.shopsuplycash "
		sqlStr = sqlStr + "    end ) as suplyprice, "

        sqlStr = sqlStr + " d.itemno,d.makerid,d.idx"

        sqlStr = sqlStr + " from "  ''[db_shop].[dbo].tbl_shop_designer n,
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m"
        sqlStr = sqlStr + "     Join [db_shop].[dbo].tbl_shopjumun_detail d"
        sqlStr = sqlStr + "     on m.orderno=d.orderno"

        sqlStr = sqlStr + "     Join [db_jungsan].[dbo].tbl_off_jungsan_master j WITH (INDEX(IX_tbl_off_jungsan_master_makerid,IX_tbl_off_jungsan_master_yyyymm))"
        sqlStr = sqlStr + "     on J.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + "     and d.makerid=J.makerid"
        sqlStr = sqlStr + "     and J.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + "     and J.taxtype='" + taxtype + "'"
        sqlStr = sqlStr + "     and J.finishflag='0'"  ''�������λ��¸�
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"
		sqlStr = sqlStr + "     left join db_summary.dbo.tbl_monthly_shop_designer n"
		sqlStr = sqlStr + "     on n.yyyymm='" + yyyymm + "'"
		sqlStr = sqlStr + "     and m.shopid=n.shopid"
		sqlStr = sqlStr + "     and d.makerid=n.makerid"

        '' �ߺ� ���� ���ϰ�.
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.linkidx "
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m WITH (INDEX(IX_tbl_off_jungsan_master_makerid,IX_tbl_off_jungsan_master_yyyymm)),"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d WITH (INDEX(IX_tbl_off_jungsan_detail_masteridx))"
        sqlStr = sqlStr + "     where m.idx=d.masteridx "

        if makerid<>"" then
            sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
        else
            sqlStr = sqlStr + "     and m.yyyymm='" + yyyymm + "'"
        end if
        sqlStr = sqlStr + "     and d.gubuncd in ('B011','B012','B013')"
        sqlStr = sqlStr + "     ) JD"
        sqlStr = sqlStr + "     on JD.linkidx=d.idx"

        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and d.jcomm_cd='"&gubuncd&"'"

        IF (makerid<>"") THEN
            sqlStr = sqlStr + "     	and ((m.shopregdate>='" + startYYYYMMDD + "' and m.shopregdate<'" + nextYYYYMMDD + "')"
            sqlStr = sqlStr + "     	    or (m.shopid='streetshop095' and m.shopregdate>='" + PreStartYYYYMMDD + "' and m.shopregdate<'" + nextYYYYMMDD + "'))"
        ELSE
            sqlStr = sqlStr + " and m.shopregdate>='" + startYYYYMMDD + "'"
            sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
        END IF

        sqlStr = sqlStr + " and m.cancelyn='N'"
        if makerid<>"" then
            sqlStr = sqlStr + " and d.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + " and d.cancelyn='N'"

        if taxtype="01" then
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')='N'"
        end if
        sqlStr = sqlStr + " and JD.linkidx is null"  '' �Է� �ȵ� ����.

        'response.write sqlStr
        'dbget.close()	:	response.End
        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B031" then
    ''��� ���� - ��Ź ��ǰ�� ���ؼ���. shopitem's centermwdiv(���Ϳ��� �������� �޴��� ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail"
        sqlStr = sqlStr + " (masteridx,shopid,gubuncd,orderno,"
        sqlStr = sqlStr + " itemgubun,itemid,itemoption,itemname,itemoptionname,"
        sqlStr = sqlStr + " sellprice,realsellprice,suplyprice,itemno,makerid,linkidx)"

        sqlStr = sqlStr + " select  J.idx as masteridx, m.socid, '" + gubuncd + "', m.code, "
        sqlStr = sqlStr + " d.iitemgubun, d.itemid, d.itemoption, d.iitemname, d.iitemoptionname,"
        sqlStr = sqlStr + " d.sellcash,d.sellcash, d.buycash,"
        sqlStr = sqlStr + " d.itemno*-1,d.imakerid,d.id"

        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer n,"
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master j,"
        sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
        sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"

        ' �ߺ� ���� ���ϰ�.
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.linkidx "
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        sqlStr = sqlStr + "     where m.idx=d.masteridx "
        sqlStr = sqlStr + "     and m.yyyymm='" + yyyymm + "'"
        if makerid<>"" then
            sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and d.gubuncd in ('B031','B021')"
        sqlStr = sqlStr + "     ) JD"
        sqlStr = sqlStr + "     on JD.linkidx=d.id"

        if (ipchulArr<>"") then     ''2012/11/08 ����
            sqlStr = sqlStr + " where m.executedt>=dateadd(m,-24,'" + startYYYYMMDD + "')"  '' -3
        else
            sqlStr = sqlStr + " where m.executedt>='" + startYYYYMMDD + "'"
		    sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
	    end if
		sqlStr = sqlStr + " and m.ipchulflag='S'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		if makerid<>"" then
            sqlStr = sqlStr + "     and d.imakerid='" + makerid + "'"
        end if
        if (ipchulArr<>"") then
            sqlStr = sqlStr + " and m.code in ("&ipchulArr&")"
        end if
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"

        if taxtype="01" then
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')='N'"
        end if

        if (makerid="ithinkso") then
            ''ithinkso �ΰ�� ����. 2013/12/02 �߰�
            sqlStr = sqlStr + " and d.mwgubun='C'"
        else
            ''sqlStr = sqlStr + " and IsNULL(s.centermwdiv,'M')='W'"
            sqlStr = sqlStr + " and IsNULL(d.mwgubun,'M')='C'"  ''������� //2014/01/27
        end if

        sqlStr = sqlStr + " and m.socid=n.shopid"
        if (ipchulArr<>"") then                                                     '''2013/02/01 �߰�
            sqlStr = sqlStr + " and d.mwgubun='C'"
        else
            sqlStr = sqlStr + " and n.chargediv in (" + chargediv + ")"
        end if
        sqlStr = sqlStr + " and n.makerid=d.imakerid"
        sqlStr = sqlStr + " and J.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + " and d.imakerid=J.makerid"
        sqlStr = sqlStr + " and J.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + " and J.taxtype='" + taxtype + "'"
        sqlStr = sqlStr + " and J.finishflag='0'"  ''�������λ��¸�
        sqlStr = sqlStr + " and JD.linkidx is null"  '' �Է� �ȵ� ����.
''rw sqlStr
        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B021" then
    ''���� ���� (���� �������� ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail"
        sqlStr = sqlStr + " (masteridx,shopid,gubuncd,orderno,"
        sqlStr = sqlStr + " itemgubun,itemid,itemoption,itemname,itemoptionname,"
        sqlStr = sqlStr + " sellprice,realsellprice,suplyprice,itemno,makerid,linkidx)"

        '' distinct..
        sqlStr = sqlStr + " select distinct J.idx as masteridx, '', '" + gubuncd + "', m.code, "
        sqlStr = sqlStr + " d.iitemgubun, d.itemid, d.itemoption, d.iitemname, d.iitemoptionname,"
        sqlStr = sqlStr + " d.sellcash,d.sellcash, d.buycash,"
        sqlStr = sqlStr + " d.itemno,d.imakerid,d.id"

        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer n,"
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master j,"
        sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master m,"
        sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_detail d"
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.iitemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"

        ' �ߺ� ���� ���ϰ�.
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.linkidx "
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        sqlStr = sqlStr + "     where m.idx=d.masteridx "
        sqlStr = sqlStr + "     and m.yyyymm='" + yyyymm + "'"
        if makerid<>"" then
            sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and d.gubuncd in ('B031','B021')"
        sqlStr = sqlStr + "     ) JD"
        sqlStr = sqlStr + "     on JD.linkidx=d.id"


        sqlStr = sqlStr + " where m.executedt>='" + startYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.executedt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.ipchulflag='I'"
		sqlStr = sqlStr + " and m.divcode='801'"
		sqlStr = sqlStr + " and m.code=d.mastercode"
		if makerid<>"" then
            sqlStr = sqlStr + "     and d.imakerid='" + makerid + "'"
        end if

        if (ipchulArr<>"") then
            sqlStr = sqlStr + " and m.code in ("&ipchulArr&")"
        end if
		sqlStr = sqlStr + " and m.deldt is NULL"
		sqlStr = sqlStr + " and d.deldt is NULL"
		sqlStr = sqlStr + " and d.itemno<>0"

        if taxtype="01" then
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')='N'"
        end if
        sqlStr = sqlStr + " and IsNULL(s.centermwdiv,'M')='M'"

        '''sqlStr = sqlStr + " and m.socid=n.shopid"
        sqlStr = sqlStr + " and n.chargediv in (" + chargediv + ")"
        sqlStr = sqlStr + " and n.makerid=d.imakerid"
        sqlStr = sqlStr + " and J.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + " and d.imakerid=J.makerid"
        sqlStr = sqlStr + " and J.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + " and J.taxtype='" + taxtype + "'"
        sqlStr = sqlStr + " and J.finishflag='0'"  ''�������λ��¸�
        sqlStr = sqlStr + " and JD.linkidx is null"  '' �Է� �ȵ� ����.

'''TimeOUT ������. : @TMP���̺� Ȱ��.
''response.write sqlStr
''response.end

        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B022" then
    ''���� ���� (��ü ����)
        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail"
        sqlStr = sqlStr + " (masteridx,shopid,gubuncd,orderno,"
        sqlStr = sqlStr + " itemgubun,itemid,itemoption,itemname,itemoptionname,"
        sqlStr = sqlStr + " sellprice,realsellprice,suplyprice,itemno,makerid,linkidx)"

        '' distinct..
        sqlStr = sqlStr + " select distinct J.idx as masteridx, m.shopid, '" + gubuncd + "', m.idx, "
        sqlStr = sqlStr + " d.itemgubun, d.shopitemid, d.itemoption, d.itemname, d.itemoptionname,"
        sqlStr = sqlStr + " d.sellcash,d.sellcash, d.suplycash,"
        sqlStr = sqlStr + " d.itemno,d.designerid,d.idx"

        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer n,"
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master j,"
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_master m,"
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shop_ipchul_detail d"
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.shopitemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"

        ' �ߺ� ���� ���ϰ�.
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.linkidx "
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        sqlStr = sqlStr + "     where m.idx=d.masteridx "
        sqlStr = sqlStr + "     and m.yyyymm='" + yyyymm + "'"
        if makerid<>"" then
            sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and d.gubuncd='" + gubuncd + "'"
        sqlStr = sqlStr + "     ) JD"
        sqlStr = sqlStr + "     on JD.linkidx=d.idx"


        sqlStr = sqlStr + " where m.execdt>='" + startYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.execdt<'" + nextYYYYMMDD + "'"
		sqlStr = sqlStr + " and m.idx=d.masteridx"
		sqlStr = sqlStr + " and m.statecd>=7"
		if makerid<>"" then
            sqlStr = sqlStr + "     and d.designerid='" + makerid + "'"
        end if
        if (ipchulArr<>"") then
            sqlStr = sqlStr + " and convert(varchar(10),m.idx) in ("&ipchulArr&")"
        end if
		sqlStr = sqlStr + " and m.deleteyn='N'"
		sqlStr = sqlStr + " and d.deleteyn='N'"
		sqlStr = sqlStr + " and d.itemno<>0"
        sqlStr = sqlStr + " and m.comm_cd='B022'"                           '''2012/06/01�߰�
        if taxtype="01" then
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')='N'"
        end if
        ''sqlStr = sqlStr + " and IsNULL(s.centermwdiv,'M')='M'"

        sqlStr = sqlStr + " and m.shopid=n.shopid"
       ''''sqlStr = sqlStr + " and n.chargediv ='" + chargediv + "'"
        sqlStr = sqlStr + " and n.makerid=d.designerid"
        sqlStr = sqlStr + " and J.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + " and d.designerid=J.makerid"
        sqlStr = sqlStr + " and J.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + " and J.taxtype='" + taxtype + "'"
        sqlStr = sqlStr + " and J.finishflag='0'"  ''�������λ��¸�
        sqlStr = sqlStr + " and JD.linkidx is null"  '' �Է� �ȵ� ����.
''rw  sqlStr
        rsget.Open sqlStr,dbget,1
    elseif gubuncd="B077" then
        rw "������"
        gubuncd = "B012"    '' ��ü��Ź���� ����

        sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail"
        sqlStr = sqlStr + " (masteridx,shopid,gubuncd,orderno,"
        sqlStr = sqlStr + " itemgubun,itemid,itemoption,itemname,itemoptionname,"
        sqlStr = sqlStr + " sellprice,realsellprice,suplyprice,itemno,makerid,linkidx)"

        sqlStr = sqlStr + " select  J.idx as masteridx, m.shopid, '" + gubuncd + "', m.orderno, "
        sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname,"
        sqlStr = sqlStr + " d.sellprice,d.realsellprice,"
        ''���԰�(���갡)
        sqlStr = sqlStr + " ( case "

        ''���� ����...(2011-03-01 : �ֹ��� ������ ���԰��� ����) => ���� ��� �̰����� ����.
        IF (makerid<>"") then
            sqlStr = sqlStr + "  	when d.suplyprice<>0 then d.suplyprice"
        END IF

        ''���ǸŰ�(���� ��� ����) - ������.(2011-03-01)
        '''sqlStr = sqlStr + "		when (n.autojungsandiv='R') then convert(int,d.realsellprice - d.realsellprice*n.defaultmargin/100)"
		''������ ���԰��� ������� : �⺻��������, d.discountprice : ���ν� �����.

		''���� ���� �Ǹ��̰�, �������� �ǸŵȰ��.
		sqlStr = sqlStr + "  	when (s.shopsuplycash=0) and (s.discountsellprice<>0) and (d.discountprice<>0) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"

		''�׿� �����Ǹ��� ��� �Һ� ��� %
		sqlStr = sqlStr + "		when (s.shopsuplycash=0) then convert(int,d.sellprice - d.sellprice*n.defaultmargin/100)"

		''(s.shopsuplycash<>0)
		sqlStr = sqlStr + "		when (d.sellprice<>s.shopitemprice) and (d.realsellprice<>s.discountsellprice) then convert(int,d.realsellprice - d.realsellprice*n.defaultmargin/100)"
		''sqlStr = sqlStr + "		when (s.discountsellprice<>0) and (d.discountprice<>s.discountsellprice) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"

		'''sqlStr = sqlStr + "  	when (s.shopsuplycash=0) and (s.discountprice<>0) and (s.discountprice<>d.discountprice) then convert(int,d.discountprice - d.discountprice*n.defaultmargin/100)"
		''������ ���԰��� �������.
		sqlStr = sqlStr + "    else s.shopsuplycash "
		sqlStr = sqlStr + "    end ) as suplyprice, "

        sqlStr = sqlStr + " d.itemno,d.makerid,d.idx"

        sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer n,"
        sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_off_jungsan_master j,"
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_master m,"
        sqlStr = sqlStr + " [db_shop].[dbo].tbl_shopjumun_detail d"
        sqlStr = sqlStr + "     left join [db_shop].[dbo].tbl_shop_item s"
		sqlStr = sqlStr + "     on d.itemgubun=s.itemgubun"
		sqlStr = sqlStr + "     and d.itemid=s.shopitemid"
		sqlStr = sqlStr + "     and d.itemoption=s.itemoption"

        '' �ߺ� ���� ���ϰ�.
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "     select d.linkidx "
        sqlStr = sqlStr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
        sqlStr = sqlStr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d "
        sqlStr = sqlStr + "     where m.idx=d.masteridx "
        sqlStr = sqlStr + "     and m.yyyymm='" + yyyymm + "'"
        if makerid<>"" then
            sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + "     and d.gubuncd in ('B011','B012')"
        sqlStr = sqlStr + "     ) JD"
        sqlStr = sqlStr + "     on JD.linkidx=d.idx"

        sqlStr = sqlStr + " where m.idx=d.masteridx"
        sqlStr = sqlStr + " and d.idx in ("                                             ''---��ǰ�ֹ��� ó��Ȯ��
        sqlStr = sqlStr + "     select bd.orgDetailidx"
        sqlStr = sqlStr + "     from db_shop.dbo.tbl_shopbeasong_order_master bm"
        sqlStr = sqlStr + "     	Join db_shop.dbo.tbl_shopbeasong_order_detail bd"
        sqlStr = sqlStr + "     	on bm.orderno=bd.orderno"
        sqlStr = sqlStr + "     where bm.cancelyn='N'"
        sqlStr = sqlStr + "     and  bd.cancelyn='N'"
        sqlStr = sqlStr + "     and bm.ipkumdiv>3"
        sqlStr = sqlStr + "     and bd.beasongdate>'" + startYYYYMMDD + "'"
        sqlStr = sqlStr + "     and bd.beasongdate<'" + nextYYYYMMDD + "'"
        sqlStr = sqlStr + "     and bd.currstate=7"
        sqlStr = sqlStr + "     and bd.isupchebeasong='Y'"
        sqlStr = sqlStr + " )"
        sqlStr = sqlStr + " and m.shopregdate>='" + CStr(DateADD("m",-3,startYYYYMMDD)) + "'"                 '''-N ��
        sqlStr = sqlStr + " and m.shopregdate<'" + nextYYYYMMDD + "'"
        sqlStr = sqlStr + " and m.cancelyn='N'"
        if makerid<>"" then
            sqlStr = sqlStr + " and d.makerid='" + makerid + "'"
        end if
        sqlStr = sqlStr + " and d.makerid=n.makerid"
        sqlStr = sqlStr + " and d.cancelyn='N'"

        if taxtype="01" then
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')<>'N'"
        else
            sqlStr = sqlStr + " and IsNULL(s.vatinclude,'Y')='N'"
        end if

        sqlStr = sqlStr + " and m.shopid=n.shopid"
        sqlStr = sqlStr + " and n.makerid in ("
        sqlStr = sqlStr + "     select distinct bd.makerid"
        sqlStr = sqlStr + "     from db_shop.dbo.tbl_shopbeasong_order_master bm"
        sqlStr = sqlStr + "     	Join db_shop.dbo.tbl_shopbeasong_order_detail bd"
        sqlStr = sqlStr + "     	on bm.orderno=bd.orderno"
        sqlStr = sqlStr + "     where bm.cancelyn='N'"
        sqlStr = sqlStr + "     and  bd.cancelyn='N'"
        sqlStr = sqlStr + "     and bm.ipkumdiv>3"
        sqlStr = sqlStr + "     and bd.beasongdate>'" + startYYYYMMDD + "'"
        sqlStr = sqlStr + "     and bd.beasongdate<'" + nextYYYYMMDD + "'"
        sqlStr = sqlStr + "     and bd.currstate=7"
        sqlStr = sqlStr + "     and bd.isupchebeasong='Y'"
        sqlStr = sqlStr + " )"
        sqlStr = sqlStr + " and n.makerid=d.makerid"
        sqlStr = sqlStr + " and J.yyyymm='" + yyyymm + "'"
        sqlStr = sqlStr + " and d.makerid=J.makerid"
        sqlStr = sqlStr + " and J.differencekey=" + CStr(differencekey)
        sqlStr = sqlStr + " and J.taxtype='" + taxtype + "'"
        sqlStr = sqlStr + " and J.finishflag='0'"  ''�������λ��¸�
        sqlStr = sqlStr + " and JD.linkidx is null"  '' �Է� �ȵ� ����.


        '''dbget.close()	:	response.End
        rsget.Open sqlStr,dbget,1
    end if
end function



function SummaryDefaultJungsanMaster(yyyymm, gubuncd)
    dim sqlStr
    dim startYYYYMMDD, nextYYYYMMDD

    startYYYYMMDD = CStr(dateserial(yyyy,mm,1))
    nextYYYYMMDD = CStr(dateserial(yyyy,mm+1,1))


    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlstr = sqlstr + " set tot_itemno=T.tot_itemno"
	sqlstr = sqlstr + " ,tot_orgsellprice=T.tot_orgsellprice"
	sqlstr = sqlstr + " ,tot_realsellprice=T.tot_realsellprice"
	sqlstr = sqlstr + " ,tot_jungsanprice=T.tot_jungsanprice"

	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, "
	sqlstr = sqlstr + "     sum(itemno) as tot_itemno,"
	sqlstr = sqlstr + "     sum(sellprice*itemno) as tot_orgsellprice, "
	sqlstr = sqlstr + "     sum(realsellprice*itemno) as tot_realsellprice, "
	sqlstr = sqlstr + "     sum(suplyprice*itemno) as tot_jungsanprice"
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1



    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"

	if (gubuncd="B011") then
	    sqlstr = sqlstr + " set TW_price=T.tot_jungsanprice"
	elseif (gubuncd="B012") or (gubuncd="B013") then
	    sqlstr = sqlstr + " set UW_price=T.tot_jungsanprice"
	elseif (gubuncd="B031") then
	    sqlstr = sqlstr + " set CM_price=T.tot_jungsanprice"
	elseif (gubuncd="B021") then
	    sqlstr = sqlstr + " set OM_price=T.tot_jungsanprice"
	elseif (gubuncd="B022") then
	    sqlstr = sqlstr + " set SM_price=T.tot_jungsanprice"
	else
	    sqlstr = sqlstr + " set ET_price=T.tot_jungsanprice"
	end if

	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, "
	sqlstr = sqlstr + "     sum(itemno) as tot_itemno,"
	sqlstr = sqlstr + "     sum(sellprice*itemno) as tot_orgsellprice, "
	sqlstr = sqlstr + "     sum(realsellprice*itemno) as tot_realsellprice, "
	sqlstr = sqlstr + "     sum(suplyprice*itemno) as tot_jungsanprice"
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	if (gubuncd="B012") or (gubuncd="B013") then
	    sqlstr = sqlstr + "     and d.gubuncd in ('B012','B013')"
	else
	    sqlstr = sqlstr + "     and d.gubuncd='" + gubuncd + "'"
    end if
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1

end function

function SummaryDefaultJungsanMasterByBrand(yyyymm, differencekey, taxtype, makerid)
    dim sqlStr
    dim startYYYYMMDD, nextYYYYMMDD

    startYYYYMMDD = CStr(dateserial(yyyy,mm,1))
    nextYYYYMMDD = CStr(dateserial(yyyy,mm+1,1))


    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlstr = sqlstr + " set tot_itemno=T.tot_itemno"
	sqlstr = sqlstr + " ,tot_orgsellprice=T.tot_orgsellprice"
	sqlstr = sqlstr + " ,tot_realsellprice=T.tot_realsellprice"
	sqlstr = sqlstr + " ,tot_jungsanprice=T.tot_jungsanprice"

	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, "
	sqlstr = sqlstr + "     sum(itemno) as tot_itemno,"
	sqlstr = sqlstr + "     sum(sellprice*itemno) as tot_orgsellprice, "
	sqlstr = sqlstr + "     sum(realsellprice*itemno) as tot_realsellprice, "
	sqlstr = sqlstr + "     sum(suplyprice*itemno) as tot_jungsanprice"
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
    sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
    sqlStr = sqlStr + "     and m.differencekey=" + CStr(differencekey) + ""
	sqlStr = sqlStr + "     and m.taxtype='" + CStr(taxtype) + "'"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.makerid='" + makerid + "'"
	sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.differencekey=" + CStr(differencekey) + ""
	sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.taxtype='" + CStr(taxtype) + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1



    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
    sqlstr = sqlstr + " set TW_price=IsNULL(T.tot_jungsanprice_TW,0)"
    sqlstr = sqlstr + " , UW_price=IsNULL(T.tot_jungsanprice_UW,0)"
    sqlstr = sqlstr + " , CM_price=IsNULL(T.tot_jungsanprice_CM,0)"
    sqlstr = sqlstr + " , OM_price=IsNULL(T.tot_jungsanprice_OM,0)"
    sqlstr = sqlstr + " , SM_price=IsNULL(T.tot_jungsanprice_SM,0)"
    sqlstr = sqlstr + " , ET_price=IsNULL(T.tot_jungsanprice_ET,0)"

	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, sum(case when d.gubuncd='B011' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_TW , "
	sqlstr = sqlstr + "     sum(case when d.gubuncd in ('B012','B013') then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_UW , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B031' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_CM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B021' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_OM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B022' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_SM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B999' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_ET  "
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.yyyymm='" + yyyymm + "'"
	sqlStr = sqlStr + "     and m.makerid='" + makerid + "'"
	sqlStr = sqlStr + "     and m.differencekey=" + CStr(differencekey) + ""
	sqlStr = sqlStr + "     and m.taxtype='" + taxtype + "'"
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.yyyymm='" + yyyymm + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.makerid='" + makerid + "'"
	sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.differencekey=" + CStr(differencekey) + ""
	sqlStr = sqlStr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.taxtype='" + taxtype + "'"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1

end function


function SummaryDefaultJungsanMasterByIdx(idx)
    dim sqlStr
    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
    sqlstr = sqlstr + " set tot_itemno=0"
	sqlstr = sqlstr + " ,tot_orgsellprice=0"
	sqlstr = sqlstr + " ,tot_realsellprice=0"
	sqlstr = sqlstr + " ,tot_jungsanprice=0"
	sqlstr = sqlstr + " ,totalcommission=0"
	sqlstr = sqlstr + " where idx=" + CStr(idx) + ""

	rsget.Open sqlStr,dbget,1


    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlstr = sqlstr + " set tot_itemno=T.tot_itemno"
	sqlstr = sqlstr + " ,tot_orgsellprice=T.tot_orgsellprice"
	sqlstr = sqlstr + " ,tot_realsellprice=T.tot_realsellprice"
	sqlstr = sqlstr + " ,tot_jungsanprice=T.tot_jungsanprice"
    sqlstr = sqlstr + " ,totalcommission=T.tot_commission"
	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, "
	sqlstr = sqlstr + "     sum(itemno) as tot_itemno,"
	sqlstr = sqlstr + "     sum(sellprice*itemno) as tot_orgsellprice, "
	sqlstr = sqlstr + "     sum(realsellprice*itemno) as tot_realsellprice, "
	sqlstr = sqlstr + "     sum(suplyprice*itemno) as tot_jungsanprice,"
	sqlstr = sqlstr + "     sum(commission*itemno) as tot_commission"
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.idx=" + CStr(idx) + ""
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1



    sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
    sqlstr = sqlstr + " set TW_price=IsNULL(T.tot_jungsanprice_TW,0)"
    sqlstr = sqlstr + " , UW_price=IsNULL(T.tot_jungsanprice_UW,0)"
    sqlstr = sqlstr + " , CM_price=IsNULL(T.tot_jungsanprice_CM,0)"
    sqlstr = sqlstr + " , OM_price=IsNULL(T.tot_jungsanprice_OM,0)"
    sqlstr = sqlstr + " , SM_price=IsNULL(T.tot_jungsanprice_SM,0)"
    sqlstr = sqlstr + " , ET_price=IsNULL(T.tot_jungsanprice_ET,0)"

	sqlstr = sqlstr + " from ("
	sqlstr = sqlstr + "     select m.idx, sum(case when d.gubuncd='B011' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_TW , "
	sqlstr = sqlstr + "     sum(case when d.gubuncd in ('B012','B013') then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_UW , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B031' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_CM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B021' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_OM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B022' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_SM , "
    sqlstr = sqlstr + "     sum(case when d.gubuncd='B999' then (suplyprice*itemno) else 0 end ) as tot_jungsanprice_ET  "
	sqlstr = sqlstr + "     from [db_jungsan].[dbo].tbl_off_jungsan_master m,"
	sqlstr = sqlstr + "         [db_jungsan].[dbo].tbl_off_jungsan_detail d"
	sqlstr = sqlstr + "     where m.idx=" + CStr(idx) + ""
	sqlstr = sqlstr + "     and m.idx=d.masteridx"
	sqlstr = sqlstr + "     and m.finishflag='0'"
	sqlstr = sqlstr + "     group by m.idx"
	sqlstr = sqlstr + " ) as T"
	sqlstr = sqlstr + " where [db_jungsan].[dbo].tbl_off_jungsan_master.idx=T.idx"
	sqlstr = sqlstr + " and [db_jungsan].[dbo].tbl_off_jungsan_master.finishflag='0'"

	rsget.Open sqlStr,dbget,1


end function

function AddBatchLog(dupleValid,jGubun,yyyymm,jstep,jsteplog)
    dim sqlStr,AssignedRow
    AddBatchLog = FALSE
    IF (dupleValid) then
        sqlStr = "IF Exists(select * from db_jungsan.dbo.tbl_jungsan_batchLog where jGubun='"&jGubun&"' and yyyymm='"&yyyymm&"' and jstep="&jstep&")"&VbCRLF
        sqlStr = sqlStr & " BEGIN"
        sqlStr = sqlStr & " update db_jungsan.dbo.tbl_jungsan_batchLog"
        sqlStr = sqlStr & " set actionCnt=actionCnt+1"
        sqlStr = sqlStr & " ,jsteplog='"&jsteplog&"'"
        sqlStr = sqlStr & " ,lastupdt=getdate()"
        sqlStr = sqlStr & " where jGubun='"&jGubun&"' and yyyymm='"&yyyymm&"' and jstep="&jstep&""
        sqlStr = sqlStr & " END"
        sqlStr = sqlStr & " ELSE"
        sqlStr = sqlStr & " BEGIN"
        sqlStr = sqlStr & " insert into db_jungsan.dbo.tbl_jungsan_batchLog"
        sqlStr = sqlStr & " (jGubun,yyyymm,jstep,jsteplog)"
        sqlStr = sqlStr & " values('"&jGubun&"'"&VbCRLF
        sqlStr = sqlStr & " ,'"&yyyymm&"'"&VbCRLF
        sqlStr = sqlStr & " ,"&jstep&""&VbCRLF
        sqlStr = sqlStr & " ,'"&jsteplog&"'"&VbCRLF
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " END"

        dbget.Execute sqlStr,AssignedRow
    ELSE
        sqlStr = "IF Not Exists(select * from db_jungsan.dbo.tbl_jungsan_batchLog where jGubun='"&jGubun&"' and yyyymm='"&yyyymm&"' and jstep="&jstep&")"&VbCRLF
        sqlStr = sqlStr & " BEGIN"
        sqlStr = sqlStr & " insert into db_jungsan.dbo.tbl_jungsan_batchLog"
        sqlStr = sqlStr & " (jGubun,yyyymm,jstep,jsteplog)"
        sqlStr = sqlStr & " values('"&jGubun&"'"&VbCRLF
        sqlStr = sqlStr & " ,'"&yyyymm&"'"&VbCRLF
        sqlStr = sqlStr & " ,"&jstep&""&VbCRLF
        sqlStr = sqlStr & " ,'"&jsteplog&"'"&VbCRLF
        sqlStr = sqlStr & " )"
        sqlStr = sqlStr & " END"

        dbget.Execute sqlStr,AssignedRow
    END IF

    AddBatchLog = (AssignedRow>0)
end function

if (mode="batchprocess") then
    ''���꼱���۾�
    if (gubuncd="0001") then
        ''���� shop_designer �ۼ�

        sqlStr = "exec db_summary.[dbo].[sp_Ten_monthly_ShopDesigner_Make] '"&YYYYMM&"',''"
        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,1,"���� �귣�� ���걸�� �ۼ� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('���� �귣�� ���걸�� �ۼ� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"

    elseif (gubuncd="0002") then

        sqlStr = "update db_shop.dbo.tbl_shopjumun_detail"
        sqlStr = sqlStr + " set suplyprice=0"
        sqlStr = sqlStr + " ,shopbuyprice=0"
        sqlStr = sqlStr + " where itemgubun='90'"
        sqlStr = sqlStr + " and itemid in (32681,34978,35215)"
        sqlStr = sqlStr + " and suplyprice<>0"
        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,2,"��ǰ�� ����/��� 0ó�� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('��ǰ�� ����/��� 0ó�� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"

        sqlStr = " update [db_shop].[dbo].tbl_shop_item"
    	sqlStr = sqlStr + " set centermwdiv=i.mwdiv"
    	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
    	sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_item.itemgubun='10'"
    	sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.shopitemid=i.itemid"
    	sqlStr = sqlStr + " and i.mwdiv<>'U'"
    	sqlStr = sqlStr + " and (([db_shop].[dbo].tbl_shop_item.centermwdiv is null) or ( i.mwdiv<>[db_shop].[dbo].tbl_shop_item.centermwdiv ))"

        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,3,"CNETER ���Ա��м��� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('CNETER ���Ա��м��� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"

        sqlStr = " update [db_shop].[dbo].tbl_shop_item"
        sqlStr = sqlStr + " set vatinclude=[db_item].[dbo].tbl_item.vatinclude"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
        sqlStr = sqlStr + " where [db_shop].[dbo].tbl_shop_item.itemgubun='10'"
        sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.shopitemid=[db_item].[dbo].tbl_item.itemid"
        sqlStr = sqlStr + " and [db_shop].[dbo].tbl_shop_item.vatinclude<>[db_item].[dbo].tbl_item.vatinclude"

        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,4,"���� ���� ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('���� ���� ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"


        sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
        sqlStr = sqlStr + " set mwgubun='C'"
        sqlStr = sqlStr + " where id in ("
        sqlStr = sqlStr + "     select d.id"
        sqlStr = sqlStr + "     from "
        sqlStr = sqlStr + "      db_summary.dbo.tbl_monthly_shop_designer sd,"      ''[db_shop].[dbo].tbl_shop_designer
        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_master m,"
        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_detail d"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s"
        sqlStr = sqlStr + " 		on d.iitemgubun=s.itemgubun and d.itemid=s.shopitemid and d.itemoption=s.itemoption"
        sqlStr = sqlStr + " where m.code=d.mastercode"
        sqlStr = sqlStr + " and m.executedt>='" & YYYY & "-" & MM & "-01'"
        sqlStr = sqlStr + " and m.executedt<'" & dateserial(YYYY,MM+1,1) & "'"
        sqlStr = sqlStr + " and m.ipchulflag='S'"                                   ''������
        sqlStr = sqlStr + " and m.deldt is null"
        sqlStr = sqlStr + " and d.deldt is null"
        sqlStr = sqlStr + " and s.centermwdiv='W'"                                  ''���� ��Ź�԰�
        sqlStr = sqlStr + " and m.socid=sd.shopid and d.imakerid=sd.makerid"
        sqlStr = sqlStr + " and sd.yyyymm=convert(varchar(7),m.executedt,21)"
        sqlStr = sqlStr + " and sd.comm_cd in ('B031')"                             ''�������.
        sqlStr = sqlStr + " and isNULL(d.mwgubun,'')=''"
        ''sqlStr = sqlStr + " and ( d.mwgubun<>'C' or d.mwgubun is null )"
        sqlStr = sqlStr + " "
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,5,"��� FLAG(C) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('��� FLAG(C) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"


        sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
        sqlStr = sqlStr + " set mwgubun='W'"
        sqlStr = sqlStr + " where id in ("
        sqlStr = sqlStr + "     select d.id"
        sqlStr = sqlStr + "     from "
        sqlStr = sqlStr + "      db_summary.dbo.tbl_monthly_shop_designer sd," ''' [db_shop].[dbo].tbl_shop_designer
        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_master m,"
        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_detail d"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s"
        sqlStr = sqlStr + " 		on d.iitemgubun=s.itemgubun and d.itemid=s.shopitemid and d.itemoption=s.itemoption"
        sqlStr = sqlStr + " where m.code=d.mastercode"
        sqlStr = sqlStr + " and m.executedt>='" & YYYY & "-" & MM & "-01'"
        sqlStr = sqlStr + " and m.executedt<'" & dateserial(YYYY,MM+1,1) & "'"
        sqlStr = sqlStr + " and m.ipchulflag='S'"                                   ''������
        sqlStr = sqlStr + " and m.deldt is null"
        sqlStr = sqlStr + " and d.deldt is null"
        ''sqlStr = sqlStr + " and s.centermwdiv='W'"                                  ''
        sqlStr = sqlStr + " and m.socid=sd.shopid and d.imakerid=sd.makerid"
        sqlStr = sqlStr + " and sd.comm_cd in ('B013')"                             ''�����Ź.
        sqlStr = sqlStr + " and isNULL(d.mwgubun,'')=''"
        ''sqlStr = sqlStr + " and ( d.mwgubun<>'C' or d.mwgubun is null )"
        sqlStr = sqlStr + " "
        sqlStr = sqlStr + " )"

        dbget.Execute sqlStr, AssignedCount

        call AddBatchLog(true,"OF",yyyymm ,6,"�����Ź FLAG(W) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
        response.write "<script>alert('��� FLAG(C) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"

''        ''2013/01/01 �߰� 10, �����/ 90 ���Ի�ǰ
''        sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail"
''        sqlStr = sqlStr + " set mwgubun='F'"
''        sqlStr = sqlStr + " where id in ("
''        sqlStr = sqlStr + "     select d.id"
''        sqlStr = sqlStr + "     from "
''        sqlStr = sqlStr + "      db_summary.dbo.tbl_monthly_shop_designer sd," ''' [db_shop].[dbo].tbl_shop_designer
''        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_master m,"
''        sqlStr = sqlStr + "     [db_storage].[dbo].tbl_acount_storage_detail d"
''        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_item s"
''        sqlStr = sqlStr + " 		on d.iitemgubun=s.itemgubun and d.itemid=s.shopitemid and d.itemoption=s.itemoption"
''        sqlStr = sqlStr + " where m.code=d.mastercode"
''        sqlStr = sqlStr + " and m.executedt>='" & YYYY & "-" & MM & "-01'"
''        sqlStr = sqlStr + " and m.executedt<'" & dateserial(YYYY,MM+1,1) & "'"
''        sqlStr = sqlStr + " and m.ipchulflag='S'"                                   ''������
''        sqlStr = sqlStr + " and m.deldt is null"
''        sqlStr = sqlStr + " and d.deldt is null"
''        sqlStr = sqlStr + " and s.centermwdiv='M'"                                  ''
''        sqlStr = sqlStr + " and m.socid=sd.shopid and d.imakerid=sd.makerid"
''        sqlStr = sqlStr + " and sd.comm_cd in ('B013')"                             ''�����Ź.
''        sqlStr = sqlStr + " and isNULL(d.mwgubun,'')=''"
''        ''sqlStr = sqlStr + " and ( d.mwgubun<>'C' or d.mwgubun is null )"
''        sqlStr = sqlStr + " "
''        sqlStr = sqlStr + " )"
''
''        dbget.Execute sqlStr, AssignedCount
''
''        call AddBatchLog(true,"OF",yyyymm ,7,"����(����)��� FLAG(F) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.")
''        response.write "<script>alert('��� FLAG(C) ���� " & AssignedCount & " �� �ݿ��Ǿ����ϴ�.');</script>"
    ''��Ź�Ǹ�
    elseif (gubuncd="B011") then
        chargediv = "2"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""


        ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, gubuncd

    ''��ü��Ź �Ǹ�
    elseif gubuncd="B012" then
        chargediv = "6"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""


        ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr


        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr


        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, gubuncd

    ''������
    elseif gubuncd="B031" then
        chargediv = "'4','5'"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""


        ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, gubuncd

    ''��������
    elseif gubuncd="B021" then
        chargediv = "'4','5'"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""


        ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, gubuncd
    ''�������
    elseif gubuncd="B022" then
        chargediv = "8"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, ""


        ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, "", ipchulArr

        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, gubuncd

        SummaryDefaultJungsanMaster yyyymm, "B021"
    ''��ü���
    elseif gubuncd="B077" then
        chargediv = "6"     '' chargeDiv : ���� ��� .. gubuncd ���·�
        differencekey = "0"
        taxtype = "01"
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"

        ''����.
        MakeDefaultJungsanMaster yyyymm, "B077", chargediv, differencekey, taxtype, titleStr, ""

        ''�鼼
        taxtype = "02"
        MakeDefaultJungsanMaster yyyymm, "B077", chargediv, differencekey, taxtype, titleStr, ""

         ''�󼼳����Է�
        taxtype = "01"
        MakeDefaultJungsanDetail yyyymm, "B077", chargediv, differencekey, taxtype, "", ipchulArr

        taxtype = "02"
        MakeDefaultJungsanDetail yyyymm, "B077", chargediv, differencekey, taxtype, "", ipchulArr

        ''���Ӹ�
        SummaryDefaultJungsanMaster yyyymm, "B012"
    else
        response.write "<script>alert('Not Valid gubun key');</script>"
        dbget.close()	:	response.End

    end if
elseif (mode="brandbatchprocess") then
    if ((jgubun="") or (makerid="") or (yyyy="") or (mm="") or (differencekey="") or (itemvatYN="")) then
        response.write "<script>alert('Not Valid Params key ');</script>"
        dbget.close()	:	response.End
    end if

    ''2014
    sqlStr = " exec db_jungsan.dbo.sp_Ten_jungsanMakeByBrandOFF '"&jgubun&"','"&makerid&"','"&yyyy+"-"+mm&"','"&itemvatYN&"','"&differencekey&"'"
    dbget.Execute sqlStr
    response.write "<script>alert('OK');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

    if differencekey<>"0" then
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� (" + differencekey + "��) ������ ����"
    else
        titleStr = Left(yyyymm,4) + "�� " + Right(yyyymm,2) + "�� ������ ����"
    end if

''���� : master�� ������ ������ �ִ´�
    sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_master "
    sqlStr = sqlStr + " (yyyymm, differencekey,taxtype,makerid,title,finishflag,groupid) "

    sqlStr = sqlStr + " select distinct '" + yyyymm + "', " + differencekey + ", '" + taxtype + "', s.makerid,"
    sqlStr = sqlStr + " '" + titleStr + "', '0', p.groupid"
    sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_designer s"

    '' �̹��� ���곻���� ���� BrandID
    sqlStr = sqlStr + "     left join [db_jungsan].[dbo].tbl_off_jungsan_master j"
    sqlStr = sqlStr + "     on j.yyyymm='" + yyyymm + "'"
    sqlStr = sqlStr + "     and j.differencekey=" + CStr(differencekey)
    sqlStr = sqlStr + "     and j.taxtype='" + taxtype + "'"
    sqlStr = sqlStr + "     and j.makerid='" + makerid + "'"

    sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner p"
    sqlStr = sqlStr + "     on s.makerid=p.id"
    sqlStr = sqlStr + "     and s.makerid='" + makerid + "'"

    sqlStr = sqlStr + " where s.makerid='" + makerid + "'"
    sqlStr = sqlStr + " and j.makerid is null"
'rw sqlStr
'response.end

    rsget.Open sqlStr,dbget,1

''���� master
    '' ���� gubuncd="B011"
'    gubuncd="B011"
'    chargediv="2"
'    MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid

    '' ���� gubuncd="B012"
'    gubuncd="B012"
'    chargediv="6"
'    MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid


    ''������ gubuncd="B031"
'    gubuncd="B031"
'    chargediv="'4','5'"
'    MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid

    ''�������� gubuncd="B021"
'    gubuncd="B021"
'    chargediv="'4','5'"
'    MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid

    ''������� gubuncd="B022"
'    gubuncd="B022"
'    chargediv="8"
'    MakeDefaultJungsanMaster yyyymm, gubuncd, chargediv, differencekey, taxtype, titleStr, makerid

''�󼼳���
    '' ���� gubuncd="B011"
    gubuncd="B011"
    chargediv="2"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr

    '' ���� gubuncd="B012"
    gubuncd="B012"
    chargediv="6"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr

    '' �����Ź gubuncd="B012"
    gubuncd="B013"
    chargediv="6"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr


    ''������ gubuncd="B031"
    gubuncd="B031"
    chargediv="'4','5'"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr

    ''�������� gubuncd="B021"
    gubuncd="B021"
    chargediv="'4','5'"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr

    ''������� gubuncd="B022"
    gubuncd="B022"
    chargediv="8"
    MakeDefaultJungsanDetail yyyymm, gubuncd, chargediv, differencekey, taxtype, makerid, ipchulArr


    SummaryDefaultJungsanMasterByBrand yyyymm, differencekey, taxtype, makerid

    response.write "<script>alert('OK');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="delmaster") then
    if (masteridx="") then
        response.write "<script>alert('Not Valid idx key');</script>"
        dbget.close()	:	response.End
    end if

    ''�������� ������ ���� ����
    sqlstr = " delete from [db_jungsan].[dbo].tbl_off_jungsan_detail" + VbCrlf
    sqlstr = sqlstr + " where detailidx in (" + VbCrlf
    sqlstr = sqlstr + "     select d.detailidx from" + VbCrlf
    sqlstr = sqlstr + "     [db_jungsan].[dbo].tbl_off_jungsan_master m," + VbCrlf
    sqlstr = sqlstr + "     [db_jungsan].[dbo].tbl_off_jungsan_detail d" + VbCrlf
    sqlstr = sqlstr + "     where m.finishflag='0'" + VbCrlf
    sqlstr = sqlstr + "     and m.idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + "     and m.idx=d.masteridx" + VbCrlf
    sqlstr = sqlstr + " )" + VbCrlf

    rsget.Open sqlStr,dbget,1

    ''�����File�� ������ ������ ���� �Ұ� //2012/12/12
    iCheExists = FALSE
    sqlstr = " select ipfileNo from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
    sqlstr = sqlstr + "     where targetGbn='OF'"
    sqlstr = sqlstr + "     and targetIdx=" + CStr(masteridx) + VbCrlf
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ipfileNo = rsget("ipfileNo")
        iCheExists = true
    end if
    rsget.close

    if (iCheExists) then
        response.write "<script>alert('��ü ���� ������ �����մϴ�.(���Ϲ�ȣ:"&ipfileNo&") ������ ������ �� �����ϴ�..');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if

    sqlstr = " delete from [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="batchnextstep") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='1'" + VbCrlf
    sqlstr = sqlstr + " where yyyymm='" + CStr(yyyymm) + "'" + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"  + VbCrlf
    sqlstr = sqlstr + " and tot_jungsanprice<>0"  + VbCrlf
    ''sqlstr = sqlstr + " and makerid not in ('zumreed','tovoroni','arevo10','emotional','twobe100','universal')"
    dbget.Execute sqlstr

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('/admin/offupchejungsan/off_jungsanlist.asp?menupos=926');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="step1to0") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='0'" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('1','2')"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="step0to1") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='1'" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="step1to3") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='3'" + VbCrlf
    sqlstr = sqlstr + " , taxregdate='" + taxregdate + "'" + VbCrlf
    ''sqlstr = sqlstr + " , taxinputdate=getdate()" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('1','2')"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="step1to3noTax") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='3'" + VbCrlf
    sqlstr = sqlstr + " , taxregdate=NULL" + VbCrlf
    sqlstr = sqlstr + " , taxinputdate=NULL" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('1','2')"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="step3to7") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='7'" + VbCrlf
    sqlstr = sqlstr + " , ipkumdate='" + ipkumdate + "'" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag='3'"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="step3to0") then
    ''check
    dim logExists
    sqlstr = " select count(idx) as cnt from [db_jungsan].[dbo].tbl_tax_history_master" + VbCrlf
    sqlstr = sqlstr + " where jungsanid=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and jungsangubun='OF'" + VbCrlf
    sqlstr = sqlstr + " and resultmsg='OK'" + VbCrlf
    sqlstr = sqlstr + " and deleteyn='N'"

    rsget.Open sqlStr,dbget,1
        logExists = (rsget("cnt")>0)
    rsget.Close


    if (logExists) then
        response.write "<script>alert('���� ��꼭 ���� ������ �����մϴ�. ������ ���� �����մϴ�.');</script>"
        response.write "<script>location.replace('" + refer + "');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if

    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='0'" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag='3'"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="step7to0") then
    '' �����ڸ� ����
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set finishflag='0'" + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag='7'"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
elseif (mode="deltaxinfo") then
    '' �����ڸ� ����
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set taxlinkidx=NULL"
    sqlstr = sqlstr + " ,neotaxno=NULL"
    sqlstr = sqlstr + " ,eseroevalseq=NULL"
    sqlstr = sqlstr + " ,taxregdate=NULL"
    sqlstr = sqlstr + " ,taxinputdate=NULL"
    sqlstr = sqlstr + " ,billsitecode=NULL"
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
    sqlstr = sqlstr + " and finishflag in ('0','1','2')"  + VbCrlf

    rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif (mode="masteretcedit") then
    sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " set comment='" + Html2Db(comment) + "'" + VbCrlf
    sqlstr = sqlstr + " ,taxtype='" + taxtype + "'" + VbCrlf
    if ipkumdate<>"" then
        sqlstr = sqlstr + " ,ipkumdate='" + ipkumdate + "'" + VbCrlf
    end if

    if taxregdate<>"" then
        sqlStr = sqlStr + " ,finishflag=(CASE WHEN finishflag='1' THEN '3' ELSE finishflag END)"
        sqlstr = sqlstr + " ,taxregdate='" + taxregdate + "'" + VbCrlf

        IF (neotaxno<>"") or (taxlinkidx="") then   '''�� �̻�..
    	    sqlStr = sqlStr + " ,neotaxno='"&neotaxno&"'"+ VbCrlf
    	    sqlStr = sqlStr + " ,billsiteCode='"&billsiteCode&"'"+ VbCrlf
        end if
        sqlStr = sqlStr + " ,eseroEvalSeq='"&eseroEvalSeq&"'"+ VbCrlf
    end if

    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf

    dbget.Execute sqlStr,AssignedRow

    ''��꼭 ���� �� ������ ������ �߰� ����.
    if (ispreFixTaxDateForce<>"") and (IsDate(preFixedTaxDate)) then
        sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
        sqlstr = sqlstr + " set preFixedTaxDate='"&preFixedTaxDate&"'"
        sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
        dbget.Execute sqlStr

    end if

    if (isrefPay<>"")  then
        PayreqIdxValid = FALSE

        sqlstr = " select D.etaxKey,p.payRealDate "
        sqlstr = sqlstr + " from db_partner.dbo.tbl_eappPayrequest p"
	    sqlstr = sqlstr + "     Join db_partner.dbo.tbl_eappPayDoc D"
	    sqlstr = sqlstr + "     on p.payrequestIdx=D.payrequestIdx"
        sqlstr = sqlstr + " where P.payrequestIdx="&refPayreqIdx
        sqlstr = sqlstr + " and D.etaxKey='"&eseroEvalSeq&"'"

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            payRealDate = rsget("payRealDate")
            PayreqIdxValid = TRUE
        end if
        rsget.Close

        IF (PayreqIdxValid) then
            sqlstr = " update [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
            sqlstr = sqlstr + " set refPayreqIdx="&refPayreqIdx&""
            IF isNULL(payRealDate) then
                sqlstr = sqlstr + " ,ipkumdate=NULL"
            ELSE
                sqlstr = sqlstr + " ,ipkumdate='"&payRealDate&"'"
                sqlStr = sqlStr + " ,finishflag=(CASE WHEN finishflag in ('0','1','3') THEN '7' ELSE finishflag END)"
            ENd IF
            sqlstr = sqlstr + " where idx=" + CStr(masteridx) + "" + VbCrlf
            dbget.Execute sqlStr
        ENd IF
    end if


    if (taxlinkidx="") then
        if (refPayreqIdx<>"")  then ''(isrefPay<>"") and and (PayreqIdxValid)
            ''���� ����..
        else
    	    sqlStr = " exec db_partner.[dbo].[sp_Ten_Esero_Tax_MatchOne] '"&eseroEvalSeq&"',2,"&masteridx&""
    	    dbget.Execute sqlStr,AssignedRow
    	    ''if (AssignedRow<1) then AssignedRow=0
    	    ''response.write "<script>alert('Tax ���� : "&AssignedRow&" ��');</script>"
    	end if
	end if

    response.write "<script>alert('OK...');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End

elseif mode="editAvailNeo" then
	sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlStr = sqlStr + " set availneo="&availneoport
	sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf

	rsget.Open sqlStr,dbget,1

    response.write "<script>alert('OK');</script>"
    response.write "<script>location.replace('" + refer + "');</script>"
    dbget.close()	:	response.End

elseif (mode="editGroupid") then
    ''���� ���� �������� üũ.
    sqlStr = "select count(idx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		IsDataExists = (rsget("cnt")>0)
	end if
	rsget.close

	if Not (IsDataExists) then
		response.write "<script>alert('���� ���� ���°� �ƴմϴ�.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
	sqlStr = sqlStr + " set groupid='" + groupid + "'"
	sqlStr = sqlStr + " where idx=" + CStr(masteridx) + VbCrlf

	rsget.Open sqlStr,dbget,1

    response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif (mode="editJAcctCd") then
    if (jacctcd="") then
        sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
    	sqlStr = sqlStr + " set jacctcd=NULL"
    	sqlStr = sqlStr + " where idx=" + CStr(masteridx) + VbCrlf
    
    	dbget.Execute sqlStr        
    else
        sqlStr = "update [db_jungsan].[dbo].tbl_off_jungsan_master"
    	sqlStr = sqlStr + " set jacctcd='" + jacctcd + "'"
    	sqlStr = sqlStr + " where idx=" + CStr(masteridx) + VbCrlf
    
    	dbget.Execute sqlStr
    end if

    response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif (mode="addetcdetail") then

    ''���� ���� �������� üũ.
    sqlStr = "select jgubun,itemvatyn,count(idx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"
    sqlstr = sqlstr + " group by jgubun,itemvatyn"

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		IsDataExists = (rsget("cnt")>0)
		jgubun      =rsget("jgubun")
		itemvatyn   =rsget("itemvatyn")

	end if
	rsget.close

	if Not (IsDataExists) then
		response.write "<script>alert('���� ���� ���°� �ƴմϴ�.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	sqlStr = " insert into [db_jungsan].[dbo].tbl_off_jungsan_detail" + VbCrlf
	sqlStr = sqlStr + "  (masteridx, shopid, gubuncd, orderno, itemgubun, itemid, itemoption, itemname"
	sqlStr = sqlStr + " ,itemoptionname,sellprice,realsellprice,suplyprice,itemno,makerid"
	sqlStr = sqlStr + " ,linkidx,commission,vatyn)"

	sqlStr = sqlStr + "  values("
	sqlStr = sqlStr + "  " + CStr(masteridx)
	sqlStr = sqlStr + "  ,'" + shopid + "'"
	sqlStr = sqlStr + "  ,'" + gubuncd + "'"
	sqlStr = sqlStr + "  ,'0'"
	sqlStr = sqlStr + "  ,'" + itemgubun + "'"
	sqlStr = sqlStr + "  ," + CStr(itemid) + ""
	sqlStr = sqlStr + "  ,'" + itemoption + "'"
	sqlStr = sqlStr + "  ,'" + itemname + "'"
	sqlStr = sqlStr + "  ,'" + itemoptionname + "'"
	sqlStr = sqlStr + "  ," + sellprice + ""
	sqlStr = sqlStr + "  ," + sellprice + ""
	sqlStr = sqlStr + "  ," + suplyprice + ""
	sqlStr = sqlStr + "  ," + itemno + ""
	sqlStr = sqlStr + "  ,'" + makerid + "'"
	sqlStr = sqlStr + "  ,0"
	sqlStr = sqlStr + "  ,'" + commission + "'"
	sqlStr = sqlStr + "  ,'" + itemvatyn + "'"
	sqlStr = sqlStr + "  )"

	rsget.Open sqlStr,dbget,1

	SummaryDefaultJungsanMasterByIdx masteridx

	response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif (mode="deldetail") then
    ''���� ���� �������� üũ.
    sqlStr = "select count(idx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		IsDataExists = (rsget("cnt")>0)
	end if
	rsget.close

	if Not (IsDataExists) then
		response.write "<script>alert('���� ���� ���°� �ƴմϴ�.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

    sqlStr = " delete from [db_jungsan].[dbo].tbl_off_jungsan_detail" + VbCrlf
    sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)
    rsget.Open sqlStr,dbget,1

    SummaryDefaultJungsanMasterByIdx masteridx

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
elseif mode="modiedtailarr" then
    ''���� ���� �������� üũ.
    sqlStr = "select jgubun,count(idx) as cnt from [db_jungsan].[dbo].tbl_off_jungsan_master " + VbCrlf
    sqlstr = sqlstr + " where idx=" + CStr(masteridx) + VbCrlf
    sqlstr = sqlstr + " and finishflag='0'"
    sqlstr = sqlstr + " group by jgubun"

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
	    jgubun       = rsget("jgubun")
		IsDataExists = (rsget("cnt")>0)
	end if
	rsget.close

	if Not (IsDataExists) then
		response.write "<script>alert('���� ���� ���°� �ƴմϴ�.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if


	idxarr = split(idxarr,"|")
	suplyprice = split(suplyprice,"|")
	itemno = split(itemno,"|")

	for i=LBound(idxarr) to UBound(idxarr)
		if Trim(idxarr(i))<>"" then
			sqlStr = " update [db_jungsan].[dbo].tbl_off_jungsan_detail" + VbCrlf
			sqlStr = sqlStr + "  set suplyprice=" + replace(CStr(suplyprice(i)),",","")  + VbCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(itemno(i))  + VbCrlf
			if (jgubun="CC") then
			    ''sqlStr = sqlStr + " ,commission=(CASE WHEN gubuncd='B999' THEN 0 ELSE realsellprice-"&replace(CStr(suplyprice(i)),",","")&" END )" + VbCrlf
			    sqlStr = sqlStr + " ,commission=realsellprice-"&replace(CStr(suplyprice(i)),",","")&"" + VbCrlf ''2014/04/07 ����
			else
			    sqlStr = sqlStr + " ,commission=0"
			end if
			sqlStr = sqlStr + "  where detailidx=" + CStr(idxarr(i))  + VbCrlf

            dbget.Execute sqlStr
			''rsget.Open sqlStr,dbget,1
		end if
	next

    SummaryDefaultJungsanMasterByIdx masteridx

	response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	response.write "<script>location.replace('" + refer + "');</script>"
    response.write "<script>window.close();</script>"
	dbget.close()	:	response.End

else
    response.write "<script>alert('Not Valid mode key');</script>"
    dbget.close()	:	response.End
end if

response.write "OK"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->