<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
' �����ϴµ�
response.end

dim mode
dim baljunum, baljuid, baljudate, itemgubun, itemid, itemoption, comment
dim i,cnt,sqlStr, errstring
dim masteridxlist, baljuname, baljucodelist, divcode, vatinclude, targetid, targetname, baljucode, brandlist, obaljucode

dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = request("mode")
baljunum = request("baljunum")
baljuid = request("baljuid")
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")
comment = request("comment")

dim itemexists, iid, newbaljucode, itemAlreadyExists, tmp

if mode="chulgoproc" then

        '�߸��� �Է�(�ڽ���ȣ�� 0 �̸鼭 �����ȣ�� �ִ°�� or realitemno �� �����鼭, �ڽ���ȣ�� ���°��)üũ
        sqlStr = " select d.itemname,d.itemoptionname "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "
        sqlStr = sqlStr + " and m.deldt is null "
        sqlStr = sqlStr + " and d.deldt is null "
        sqlStr = sqlStr + " and b.baljucode = m.baljucode "
        sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + " and (((isnull(d.packingstate,0) = 0) and (isnull(d.boxsongjangno,'0') <> '0')) or ((d.realitemno > 0) and (isnull(d.boxsongjangno,'0') = '0'))) "

        if (baljuid <> "") then
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
        end if

        if (baljunum <> "") then
                sqlStr = sqlStr + " and b.baljunum = '" + CStr(baljunum) + "' "
        end if

        if (baljudate <> "") then
                sqlStr = sqlStr + " and b.baljudate >= '" + CStr(baljudate) + "' "
                sqlStr = sqlStr + " and b.baljudate < '" + CStr(Left(dateadd("d",1,baljudate),10)) + "' "
        end if

        rsget.Open sqlStr, dbget, 1
        if  not rsget.EOF  then
                do until rsget.eof
                        if (trim(errstring) = "") then
                                errstring = rsget("itemname") + "(" + rsget("itemoptionname") + ")"
                        else
                                errstring = errstring + ", " + rsget("itemname") + "(" + rsget("itemoptionname") + ")"
                        end if

                        rsget.MoveNext
                loop
        else
                errstring = ""
        end if
        rsget.close

        if (errstring <> "") then
                response.write "<script>alert('�߸��� �Է��� �ֽ��ϴ�. �����ȣ �Ǵ� ��ǰ�ڵ带 ������ �ٽ� �Է��ϼ���.\n\n" + errstring + "');</script>"
                response.write "<script>history.back();</script>"
                dbget.close()	:	response.End
        end if


        '�ش� �����ڵ�/���־��̵� ���� masteridx �� ���Ѵ�.
        sqlStr = " select distinct d.masteridx "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "
        sqlStr = sqlStr + " and m.deldt is null "
        sqlStr = sqlStr + " and d.deldt is null "
        sqlStr = sqlStr + " and b.baljucode = m.baljucode "
        sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + " and m.statecd <> '7' "
        sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
        sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
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
        rsget.close

        if (masteridxlist = "") then
                response.write "<script>alert('�ش� ���ֹ�ȣ/�� �� ���� ���ó���� �Ҽ� �����ϴ�.\n�̹� ���Ϸ�� �ֹ����� �����ɼ� �����ϴ�.');</script>"
                response.write "<script>history.back();</script>"
                dbget.close()	:	response.End
        end if


		'�⺻ master ����
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where idx in (" + masteridxlist + ") "
		rsget.Open sqlStr, dbget, 1
		'response.write sqlStr


        '�̹���ֹ���ǰ�� ���� �ڸ�Ʈ �Է�
        if (Trim(comment) <> "") then
        	itemgubun = split(itemgubun,"|")
        	itemid = split(itemid,"|")
        	itemoption = split(itemoption,"|")
        	comment = split(comment,"|")
        	cnt = ubound(itemgubun)

        	for i=0 to cnt
    	        if (Trim(comment(i)) <> "") then
    	                '1. �̹�� ��ǰ�� ���� �ڸ�Ʈ �Է�
    	                sqlstr = " update [db_storage].[dbo].tbl_ordersheet_detail "
    	                sqlstr = sqlstr + " set comment = '" + Trim(comment(i)) + "' "
    	                sqlstr = sqlstr + " where itemgubun = '" + Trim(itemgubun(i)) + "' "
    	                sqlstr = sqlstr + " and itemid = " + Trim(itemid(i)) + " "
    	                sqlstr = sqlstr + " and itemoption = '" + Trim(itemoption(i)) + "' "
    	                sqlstr = sqlstr + " and masteridx in (" + CStr(masteridxlist) + ") "
    	                sqlstr = sqlstr + " and baljuitemno <> realitemno "
    	                sqlstr = sqlstr + " and deldt is null "
    	                rsget.Open sqlStr, dbget, 1
                    end if
            next


        	'2. �̹���ֹ� ���� üũ(���ֹ� ��� ��ǰ �˻�)
        	sqlStr = " select count(d.idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail d "
        	sqlStr = sqlStr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
        	sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
        	sqlStr = sqlStr + " and d.comment='5�ϳ����' "
        	sqlStr = sqlStr + " and deldt is null "
                'response.write sqlStr
        	rsget.Open sqlStr, dbget, 1
        	itemexists = (rsget("cnt")>0)
        	rsget.Close

        	sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
        	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
        	sqlStr = sqlStr + " and clinkcode  is not null "
        	sqlStr = sqlStr + " and clinkcode<>'' "
        	rsget.Open sqlStr, dbget, 1
        	itemAlreadyExists = (rsget("cnt")>0)
        	rsget.Close

        	if Not itemexists then
        		'response.write "<script>alert('�� �ֹ��� ������ �����ϴ�.');</script>"
        	elseif itemAlreadyExists then
        		'response.write "<script>alert('�� �ֹ����� �̹� �ۼ��Ǿ� �ֽ��ϴ�. �ۼ��� �� �����ϴ�.');</script>"
        	else
            	'�̹�� �ֹ��� �ۼ�
            	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
            	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
            	rsget.Open sqlStr, dbget, 1
        		targetid = rsget("targetid")
        		targetname = rsget("targetname")
        		divcode = rsget("divcode")
        		vatinclude = rsget("vatinclude")
            	rsget.Close

                '�ش� �����ڵ�/���־��̵� ���� �⺻������ ���Ѵ�.
                sqlStr = " select distinct m.baljuname, m.baljucode "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
                sqlStr = sqlStr + " and m.statecd <> '7' "
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
                sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
            	sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
            	sqlStr = sqlStr + " and d.comment='5�ϳ����' "
            	sqlStr = sqlStr + " and d.deldt is null "
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
                rsget.close



            	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0 "
            	rsget.Open sqlStr,dbget,1,3
            	rsget.AddNew
            	rsget("targetid") = targetid
            	rsget("targetname") = targetname
            	rsget("baljuid") = baljuid
            	rsget("baljuname") = baljuname
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
            	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
            	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
            	sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
            	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
            	sqlStr = sqlStr + " sum(baljuitemno-realitemno),sum(baljuitemno-realitemno),baljudiv" + vbCrlf
            	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
            	sqlStr = sqlStr + " where masteridx in (" + CStr(masteridxlist) + ") "
            	sqlStr = sqlStr + " and baljuitemno <> realitemno "
            	sqlStr = sqlStr + " and comment='5�ϳ����'"
            	sqlStr = sqlStr + " and deldt is null"
            	sqlStr = sqlStr + " group by itemgubun,makerid,itemid,itemoption,itemname,itemoptionname,sellcash,suplycash,buycash,baljudiv "
            	rsget.Open sqlStr, dbget, 1


            	''���Ӹ� ����
            	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
            	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
            	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
            	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
            	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
            	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
            	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
            	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
            	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
            	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
            	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
            	sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
            	sqlStr = sqlStr + " and deldt is null" + vbCrlf
            	sqlStr = sqlStr + " ) as T" + vbCrlf
            	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)
            	rsget.Open sqlStr, dbget, 1


            	''�귣�� ����Ʈ
            	brandlist = ""
            	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
            	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
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
            	'sqlStr = sqlStr + " , obaljucode='" + obaljucode + "'" + VbCrlf
            	sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
            	sqlStr = sqlStr + " where idx=" + CStr(iid)
            	rsget.Open sqlStr, dbget, 1


            	''�����ּ��� ��ũ�ڵ� ����.
            	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
            	sqlStr = sqlStr + " set clinkcode='" + baljucode + "'" + VbCrlf
            	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
            	rsget.Open sqlStr, dbget, 1

               	'response.write "<script>alert('�� �ֹ����� �ۼ��Ǿ� �ֽ��ϴ�.');</script>"
        	end if

        end if



        '�� �ֹ��ڵ庰 ó��(�����Ÿ ���� ��)
        tmp = split(masteridxlist,",")
        for i=0 to UBound(tmp)
                if (Trim(tmp(i)) <> "") then

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
                	rsget.Open sqlStr, dbget, 1


                    '�ش� �����ڵ�/���־��̵� ���� �⺻������ ���Ѵ�.
                    sqlStr = " select distinct m.baljuname, m.baljucode "
                    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                    sqlStr = sqlStr + " where 1 = 1 "
                    sqlStr = sqlStr + " and m.idx = d.masteridx "
                    sqlStr = sqlStr + " and m.deldt is null "
                    sqlStr = sqlStr + " and d.deldt is null "
                    sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                    sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
                    sqlStr = sqlStr + " and m.statecd <> '7' "
                    sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
                    sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
                    sqlStr = sqlStr + " and m.idx = " + CStr(Trim(tmp(i))) + " "
                    rsget.Open sqlStr,dbget,1

                    baljuname = ""
                    baljucode = ""
                    if  not rsget.EOF  then
                            baljuname = CStr(rsget("baljuname"))
                            baljucode = CStr(rsget("baljucode"))
                    end if
                    rsget.close


                	''��� ����Ÿ�� �Է�. *-1
                	sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d "
                	sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
                	sqlStr = sqlStr + " and d.deldt is null "
                	sqlStr = sqlStr + " and d.realitemno <> 0 "
                	rsget.Open sqlStr, dbget, 1
                        itemexists = rsget("cnt")>0
                	rsget.close

                	if itemexists then
                		'1.�¶��� ��� ����Ÿ
                		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0 "
                		rsget.Open sqlStr,dbget,1,3
                		rsget.AddNew
                		rsget("code") = ""
                		rsget("socid") = baljuid
                		rsget("socname") = baljuname
                		rsget("chargeid") = session("ssBctId")
                		rsget("divcode") = "006"
                		rsget("vatcode") = "008"
                		rsget("comment") = baljucode + " �ֹ� �ڵ����ó��"
                		rsget("chargename") = session("ssBctCname")
                		rsget("ipchulflag") = "S"

                		rsget.update
                		iid = rsget("id")
                		rsget.close

                		newbaljucode = "SO" + Format00(6,Right(CStr(iid),6))

                		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
                		sqlStr = sqlStr + " set code='" + newbaljucode + "'" + VBCrlf
                		sqlStr = sqlStr + " where id=" + CStr(iid)
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
                		rsget.Open sqlStr,dbget,1


                		'4.���º���
                		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
                		sqlStr = sqlStr + " set statecd='7'" + vbCrlf
                		sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
                		sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
                		sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
                		rsget.Open sqlStr, dbget, 1

                        '' ��/��� ���ݿ�
                        sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & newbaljucode & "','','',0,'',''"
	                    dbget.Execute sqlStr


                        '5.���� ���� �����Ǹ� �缳��("�ֹ���-Ȯ����" ��ŭ ���ش�.) -> ���� ��ǰ�غ� ��ȯ�� ��.
                        'sqlstr = " update [db_item].[dbo].tbl_item "
                        'sqlstr = sqlstr + " set limitsold=limitsold - T.itemno "
                        'sqlstr = sqlstr + " from ( "
                        'sqlstr = sqlstr + " select sum(d.baljuitemno - d.realitemno) as itemno, d.itemid "
                        'sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
                        'sqlstr = sqlstr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
                        'sqlstr = sqlstr + " and d.itemid=i.itemid "
                        'sqlstr = sqlstr + " and d.deldt is null "
                        'sqlstr = sqlstr + " and d.baljuitemno <> d.realitemno "
                        'sqlstr = sqlstr + " and d.itemgubun = '10' "
                        'sqlstr = sqlstr + " and i.limityn='Y' "
                        'sqlstr = sqlstr + " group by d.itemid "
                        'sqlstr = sqlstr + " ) as T "
                        'sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid "
                        'rsget.Open sqlStr, dbget, 1

                    end if
                end if
        next

        '' ���� �������� ����
        sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
        dbget.Execute sqlStr

end if

%>
<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('baljulistoffline.asp');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
