<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �ֹ�ó��
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
function FnChangForeignMallUpbeaDLV(iItemID,iItemOption,imakerid,byref t_upchebeasong,byref buf_mwdiv,byref buf_iitembuycash,byref buf_orgprice,byref buf_orgsuplycash)
    Dim sqlStr
    Dim offMarginExists : offMarginExists=false
    Dim comm_cd, defaultCenterMwDiv,defaultmargin, shopitemprice, shopsuplycash, shopbuyprice

    sqlStr = "select top 1 sd.shopid,sd.comm_cd,sd.defaultmargin,sd.defaultsuplymargin" &VbCRLF
    sqlStr = sqlStr & " ,(select isNULL(defaultCenterMwDiv,'') from db_shop.dbo.tbl_shop_designer where shopid='streetshop000' and makerid='yougreat') as defaultCenterMwDiv"&VbCRLF
    sqlStr = sqlStr & " ,si.shopitemprice, isNULL(si.shopsuplycash,0) as shopsuplycash,isNULL(si.shopbuyprice,0) as shopbuyprice"&VbCRLF
    sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_designer sd WITH(NOLOCK) "&VbCRLF
    sqlStr = sqlStr & "     left join db_shop.dbo.tbl_shop_item si WITH(NOLOCK) "&VbCRLF
	sqlStr = sqlStr & "     on si.itemgubun='10'"&VbCRLF
	sqlStr = sqlStr & "     and si.shopitemid="&iItemID&VbCRLF
	sqlStr = sqlStr & "     and si.itemoption='"&iItemOption&"'"&VbCRLF
    sqlStr = sqlStr & " where sd.shopid in ('streetshop000','streetshop700')"&VbCRLF
    sqlStr = sqlStr & " and sd.makerid='"&imakerid&"'"&VbCRLF
    sqlStr = sqlStr & " order by sd.comm_cd desc, sd.defaultmargin desc, sd.shopid"&VbCRLF

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        offMarginExists = true
        comm_cd = rsget("comm_cd")
        defaultCenterMwDiv = rsget("defaultCenterMwDiv")
        defaultmargin = rsget("defaultmargin")
        shopitemprice = rsget("shopitemprice")
        shopsuplycash = rsget("shopsuplycash")
        shopbuyprice = rsget("shopbuyprice")

        if (comm_cd="B031") then ''�������� ���.
            t_upchebeasong = "N"

            if (defaultCenterMwDiv="M") then
                buf_mwdiv = "M"
            end if

            if isNULL(shopitemprice) then ''���� ��� ��ǰ�� �ƴѰ��..
                offMarginExists = false      '' �߰� ����
            else
                if (buf_orgprice>shopitemprice) then buf_orgprice=shopitemprice

                if (shopsuplycash=0) then
                    buf_orgsuplycash = CLNG(buf_orgprice*(100-defaultmargin)/100)
                    buf_iitembuycash = buf_orgsuplycash
                else
                    buf_orgsuplycash = shopsuplycash
                end if
            end if
        else
            offMarginExists = false
        end if


    end if
    rsget.Close

    ''������ ������ ��Ź(W), ��ü���, �¶��� �⺻ ����.
    if (NOT offMarginExists) then
        t_upchebeasong = "N"
        buf_mwdiv ="W"
    end if
end function

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

dim cksel, dummyseqarr, countryCode, companygubun, splitedSeq,ixsiteOrderSerial, i, j, sqlStr, buf
dim isexist, tenOrderSerial, OutMallOrderSerial, orderItemName, orderItemOptionName, mode, reguserid, isexist2
Dim ErrMsg, iid, OrderSerial, outMallorderSeq, orderItemID, orderItemOption, AssignedRow
dim buf_sellcash, buf_sellvat, buf_mileage, buf_totcost, buf_totvat, buf_sellcount, buf_itemdiv, buf_iitemname, buf_CpnNotAppliedSellcash, buf_totCpnNotAppliedcost, buf_CurSellcash
dim buf_iitemoptionname, buf_iitembuycash, buf_iitembuyvat, buf_onlyitembuycash, buf_onlyoptaddbuyprice, buf_onlyoptaddprice
dim buf_iitemmakerid, buf_iitemvatinclude, buf_deliverytype , buf_mwdiv, buf_sailsellcash, buf_sailbuycash
dim buf_sailyn, buf_orgprice, buf_orgsuplycash, mayOrderDate, t_upchebeasong
	mode = requestCheckVar(html2db(request("mode")),32)
	cksel = request("cksel")

rw mode
rw cksel

Dim sumItemOrderCount, sumRealsellprice, avgRealsellprice, orgdetailkey, orgdetailkeyMin, orgdetailkeyNotMin, requireDetailAdd
Dim orgdetailkeyGRoup, matchItemid, matchItemoption
Dim orgdetailkeylength, lp
Dim sitegbn

dim isbatchMode : isbatchMode = (request("xtype")="batch")
dim oseq : oseq = requestCheckvar(request("oseq"),10)
dim preactedorder : preactedorder = false
Dim kakaogiftCount

if (isbatchMode) then
	sqlStr = "select 1 from db_temp.[dbo].[tbl_xSite_TMPOrder_BatchAct] where OutMallOrderSeq="&oseq
	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if NOT rsget.Eof then
		preactedorder = true
	end if
	rsget.Close

	if (preactedorder) then
		response.write "<script>parent.addResultLog("&oseq&",'ERR1');fnNextOrderInputProc();</script>"
		dbget.Close() : response.end
	end if

	sqlStr = "insert into db_temp.[dbo].[tbl_xSite_TMPOrder_BatchAct](outmallorderseq,actuser)"
	sqlStr = sqlStr & " values("&oseq&",'"&session("ssBctID")&"')"
	dbget.Execute(sqlStr)

end if

if (mode = "add") then
	''response.write "TEST��"
	''response.end

	'1|2|3 >> 0,1,2,3
	dummyseqarr = cksel
	dummyseqarr = Replace(dummyseqarr, ", ", ",")
	dummyseqarr = Replace(dummyseqarr, ",", "','")
	dummyseqarr = "'"&dummyseqarr&"'"

	'shopify �ϴ� �ֹ� ����
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder as T WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and sellsite = 'shopify' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If rsget("cnt") > 0 Then
		response.write "<script>alert('���������� shopify�� �ֹ��Է� ���ҽ��ϴ�.')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

	''�ֹ����۹��� üũ
	''2016-04-15 ������ �߰�
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder as T WITH(NOLOCK) "
	sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on T.orderitemid = i.itemid "
	sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and isnull(T.requireDetail, '') = '' "
	sqlStr = sqlStr & " and T.sellsite in ('nvstorefarm', 'nvstoremoonbangu', 'nvstoregift', 'Mylittlewhoopee') "
	sqlStr = sqlStr & " and i.itemdiv = '06' "
	sqlStr = sqlStr & " and T.orderserial is null "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If rsget("cnt") > 0 Then
		response.write "<script>alert('�������߿� �ֹ����۹��� �������� �ֽ��ϴ�\n�ֹ����۹����� �����ϼ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

	''2017-06-01 ������ �߰�
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder as T WITH(NOLOCK) "
	sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on T.orderitemid = i.itemid "
	sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
'	sqlStr = sqlStr & " and isnull(T.requireDetail11stYN, '') = ''  "
	sqlStr = sqlStr & " and LEN(isNULL(T.requiredetail, '')) = 0  "
	sqlStr = sqlStr & " and T.sellsite = '11st1010' "
	sqlStr = sqlStr & " and i.itemdiv = '06' "
	sqlStr = sqlStr & " and T.orderserial is null "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    If rsget("cnt") > 0 Then
		response.write "<script>alert('�������߿� �ֹ����۹����� �ֽ��ϴ�\n�ֹ����۹����� �����ϼ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

	''zipcode üũ
	''2013-11-27 ������ �߰�
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE Len(isNULL(replace(receiveZipCode,'-',''),''))<5 "		'2015-10-14 17:17 ������ <6 ���� <5�� ����
	sqlStr = sqlStr & " and OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and sellsite <> 'cnglob10x10' "
	sqlStr = sqlStr & " and sellsite <> 'cnhigo' "
	sqlStr = sqlStr & " and sellsite <> '11stmy' "
	sqlStr = sqlStr & " and sellsite <> 'shopify' "
	sqlStr = sqlStr & " and sellsite <> 'cnugoshop' "
	sqlStr = sqlStr & " and sellsite <> 'etsy' "
	sqlStr = sqlStr & " and sellsite <> 'zilingo' "
	sqlStr = sqlStr & " and sellsite <> 'nvstorefarmclass' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		response.write "<script>alert('�������߿� �����ȣ�� �� �� �Ȱ��� �ֽ��ϴ�\n�����ȣ ������ư�� Ŭ���ϼż� �����ϼ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close


	' ''2021-12-17 ������ �߰�
	' If session("ssBctID") <> "kjy8517" Then
	' 	sqlStr = ""
	' 	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	' 	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	' 	sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
	' 	sqlStr = sqlStr & " and sellsite = 'gseshop' "
	' 	rsget.CursorLocation = adUseClient
	' 	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	' 	If rsget("cnt") > 0 Then
	' 		response.write "<script>alert('�ӽ÷� �ֹ���� �����ϴ�.by kjy8517')</script>"
	' 		dbget.close()	:	response.End
	' 	End If
	' 	rsget.Close
	' End If

    ''�ǸŰ� üũ ������
    '2014/03/10
    sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE sellprice<realsellprice"
	sqlStr = sqlStr & " and OutMallOrderSerial in (" & dummyseqarr & ") "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		response.write "<script>alert('�ǸŰ����� ���ǸŰ��� �� ū ������ �ֽ��ϴ�. ������ ���� ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

    ''�̹� ���۵� �ֹ������� check
    sqlStr = " select top 1 T.OutMallOrderSerial, m.orderserial from " + VbCrlf
    sqlStr = sqlStr + " db_temp.dbo.tbl_xSite_TMPOrder T WITH(NOLOCK) " + VbCrlf
    sqlStr = sqlStr + " 	Join db_order.dbo.tbl_order_master m WITH(NOLOCK) " + VbCrlf
    sqlStr = sqlStr + " 	on T.OutMallOrderSerial=m.authcode" + VbCrlf
    sqlStr = sqlStr + " 	and m.sitename=T.sellSite" + VbCrlf
    sqlStr = sqlStr + " where T.OutMallOrderSerial in (" & dummyseqarr & ") " + VbCrlf

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    isexist = (not rsget.EOF)
    if (isexist = true) then
	    OutMallOrderSerial 		= rsget("OutMallOrderSerial")
	    tenOrderSerial          = rsget("orderserial")
	end if
    rsget.Close

    ''���ǸŰ� üũ ������
    '2019/02/07
    sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE realsellprice < 1"
	sqlStr = sqlStr & " and OutMallOrderSerial in (" & dummyseqarr & ") "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		response.write "<script>alert('���ǸŰ��� 0���� ������ �ֽ��ϴ�.')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

	'2015-12-23 ������ matchitemoption�� FF�� ���� üũ
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE  OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and left(matchitemoption,2) ='FF' "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		response.write "<script>alert('�ɼ��� FF�� �����ϴ� ���� �ֽ��ϴ�. Ȯ�� �� �Է��ϼ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

	'2016-03-09 ������ 3���� �̻� ���� �����ʹ� �ֹ��Է� �� �ǰ� üũ
	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_XSite_TMporder WITH(NOLOCK) "
	sqlStr = sqlStr & " WHERE  OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and datediff(m, regdate, getdate()) > 3 "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		response.write "<script>alert('3�����̻� ���� �ֹ����� �Է� �Ͻ� �� �����ϴ�.')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	End If
	rsget.Close

    ''��ۿ�û���� üũ(�ֹ����۹��� ���ԵǴ� ��찡 ����)..2021-11-16 ������ �߰�
	sqlStr = ""
	sqlStr = sqlStr & " select "
	sqlStr = sqlStr & " OutMallOrderSerial, count(*) "
	sqlStr = sqlStr & " from ( "
	sqlStr = sqlStr & " 	select OutMallOrderSerial, deliverymemo, count(*) as cnt "
	sqlStr = sqlStr & " 	from db_temp.dbo.tbl_xSite_TMPOrder with (nolock) "
	sqlStr = sqlStr & " 	where OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " 	group by OutMallOrderSerial, deliverymemo "
	sqlStr = sqlStr & " ) as t "
	sqlStr = sqlStr & " group by OutMallOrderSerial "
	sqlStr = sqlStr & " having count(*)> 1 "
	sqlStr = sqlStr & " order by 1 "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    isexist2 = (not rsget.EOF)
    if (isexist2 = true) then
		response.write "<script>alert('��ۿ�û������ �ٸ��� �Էµ��ֽ��ϴ�. �������ּ���')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if
    rsget.Close

    if (isexist) then
        '' CS�� ����
        sqlStr = "update T"
        sqlStr = sqlStr & " set matchstate='C'"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " where T.OutMallOrderSerial='"&OutMallOrderSerial&"'"
        sqlStr = sqlStr & " and T.matchstate='I'"
        sqlStr = sqlStr & " and IsNULL(T.orderCSGbn,0) in (3,8)"
        sqlStr = sqlStr & " and T.sellsite in ('lotteimall','lotteCom')"
        dbget.Execute sqlStr

        sqlStr = "update T"
        sqlStr = sqlStr & " set orderCSGbn=8"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m"
        sqlStr = sqlStr & " 	on T.outmallorderserial=m.authcode"
        sqlStr = sqlStr & " 	and T.sellsite=m.sitename"
        sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
        sqlStr = sqlStr & " 	on m.orderserial=D.orderserial"
        sqlStr = sqlStr & " 	and D.itemid=T.matchitemid"
        sqlStr = sqlStr & " 	and D.itemoption=T.matchItemOption"
        sqlStr = sqlStr & " where T.OutMallOrderSerial='"&OutMallOrderSerial&"'"
        sqlStr = sqlStr & " and T.matchstate='I'"
        sqlStr = sqlStr & " and T.orderserial is NULL"
        sqlStr = sqlStr & " and IsNULL(T.orderCSGbn,0)=0"
        dbget.Execute sqlStr

		response.write "<script>alert('ERROR : �̹� ���۵� �ֹ���ȣ:" & CStr(OutMallOrderSerial) & " Ten�ֹ���ȣ:" & CStr(tenOrderSerial) & "')</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

	'�� �ֹ��� ���� ��Ī�� ��� �̷�������� Ȯ��
	sqlStr = " SELECT TOP 1 " + VbCrlf
	sqlStr = sqlStr + " 	T.OutMallOrderSerial " + VbCrlf
	sqlStr = sqlStr + " 	, T.orderItemName " + VbCrlf
	sqlStr = sqlStr + " 	, IsNULL(T.orderItemOptionName,'') as orderItemOptionName ,t.countryCode" + VbCrlf
	sqlStr = sqlStr + " from db_temp.dbo.tbl_xSite_TMPOrder T WITH(NOLOCK) " + VbCrlf
	sqlStr = sqlStr + " left join db_item.dbo.tbl_item i WITH(NOLOCK) " + VbCrlf
	sqlStr = sqlStr + " 	on T.matchItemID=i.itemid " + VbCrlf
	sqlStr = sqlStr + " left join db_item.dbo.tbl_item_option o WITH(NOLOCK) " + VbCrlf
	sqlStr = sqlStr + " 	on T.matchItemID=o.itemid " + VbCrlf
	sqlStr = sqlStr + " 	and T.matchItemOption=o.itemoption " + VbCrlf
	sqlStr = sqlStr + " where " + VbCrlf
	sqlStr = sqlStr + " 	1 = 1 " + VbCrlf
	''sqlStr = sqlStr + " 	and IsNull(i.itemid, '') = '' " + VbCrlf  ''2013/01/09 ���� �ɼǰ˻�.
	sqlStr = sqlStr + " 	and ( (IsNull(i.itemid, '') = '')"
	sqlStr = sqlStr + " 			OR "
	sqlStr = sqlStr + " 			((T.matchItemOption<>'0000') and (isNULL(o.optionname,'')=''))"
	sqlStr = sqlStr + " 	)"
	sqlStr = sqlStr + " 	and T.OutMallOrderSerial in (" & dummyseqarr & ") " + VbCrlf

	rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    	isexist = (not rsget.EOF)
    	if (isexist = true) then
	    	OutMallOrderSerial 		= rsget("OutMallOrderSerial")
	    	countryCode 		= rsget("countryCode")
	    	orderItemName			= rsget("orderItemName")
	    	orderItemOptionName		= rsget("orderItemOptionName")
	    	IF IsNULL(orderItemOptionName) then orderItemOptionName=""
	    end if
    rsget.close

	if (isexist) then
		if ucase(countryCode) <> "KR" then
			response.write "<script>alert('ERROR : ��ǰ�� ���ε��� �ʾҽ��ϴ�. �ֹ���ȣ:" & CStr(OutMallOrderSerial) & " ����ǰ��:" & CStr(orderItemName) & " �ɼ�:" & CStr(orderItemOptionName) & "')</script>"
			dbget.close()	:	response.End
		end if
	end if

    ''�ɼ� üũ1
    sqlStr = " select SUM(CASE WHEN T.matchitemoption='0000' and i.optionCNT>0 THEN 1 ELSE 0 END) ckCNT"
	sqlStr = sqlStr + " from db_temp.[dbo].tbl_xSite_TMPOrder T WITH(NOLOCK) "
	sqlStr = sqlStr + " left join db_item.dbo.tbl_item i WITH(NOLOCK) "
	sqlStr = sqlStr + " on matchitemid=i.itemid"
    sqlStr = sqlStr + " where T.OutMallOrderSerial in (" & dummyseqarr & ") " + VbCrlf

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        isexist = rsget("ckCNT")>0
    end if
    rsget.close

    if (isexist) then
        response.write "<script>alert('ERROR : ��ǰ�ɼ��ڵ� Ȯ�ο�� - ������ ���� ���.')</script>"
    	dbget.close()	:	response.End
    end if

	'�� �ֹ��� ���� ������ ��� �̷�������� Ȯ��
	sqlStr = " select " + VbCrlf
	sqlStr = sqlStr + " 	T.OutMallOrderSerial " + VbCrlf
	sqlStr = sqlStr + " from " + VbCrlf
	sqlStr = sqlStr + " 	( " + VbCrlf
	sqlStr = sqlStr + " 		select " + VbCrlf
	sqlStr = sqlStr + " 			OutMallOrderSerial " + VbCrlf
	sqlStr = sqlStr + " 			, count(OutMallOrderSerial) as cnt " + VbCrlf
	sqlStr = sqlStr + " 			, sum(case when OutMallOrderSerial in (" & dummyseqarr & ") then 1 else 0 end) as chk " + VbCrlf
	sqlStr = sqlStr + " 		from " + VbCrlf
	sqlStr = sqlStr + " 		db_temp.dbo.tbl_xSite_TMPOrder WITH(NOLOCK) " + VbCrlf
	sqlStr = sqlStr + " 		where " + VbCrlf
	sqlStr = sqlStr + " 			1 = 1 " + VbCrlf
	sqlStr = sqlStr + " 			and OutMallOrderSerial in ( " + VbCrlf
	sqlStr = sqlStr + " 				select " + VbCrlf
	sqlStr = sqlStr + " 					OutMallOrderSerial " + VbCrlf
	sqlStr = sqlStr + " 				from " + VbCrlf
	sqlStr = sqlStr + " 				db_temp.dbo.tbl_xSite_TMPOrder " + VbCrlf
	sqlStr = sqlStr + " 				where " + VbCrlf
	sqlStr = sqlStr + " 					1 = 1 " + VbCrlf
	sqlStr = sqlStr + " 					and OutMallOrderSerial in (" & dummyseqarr & ") " + VbCrlf
	sqlStr = sqlStr + " 			) " + VbCrlf
	sqlStr = sqlStr + " 		group by " + VbCrlf
	sqlStr = sqlStr + " 			OutMallOrderSerial " + VbCrlf
	sqlStr = sqlStr + " 	) T " + VbCrlf
	sqlStr = sqlStr + " where T.cnt <> T.chk " + VbCrlf

'	rsget.Open sqlStr,dbget,1
'    	isexist = (not rsget.EOF)
'
'    	if (isexist = true) then
'	    	OutMallOrderSerial 		= rsget("OutMallOrderSerial")
'	    end if
'    rsget.close

'	if (isexist) then
'		response.write "<script>alert('ERROR : �ϳ��� �ֹ��ǿ� ���� ��� ��ǰ�� ���õǾ�� �մϴ�. �ֹ���ȣ:" & CStr(OutMallOrderSerial) & "')</script>"
'		response.write "<script>history.back();</script>"
'		dbget.close()	:	response.End
'	end if

''rw dummyseqarr


splitedSeq = split(cksel,",")
Dim otmpOrder
dim IsForeignDLV '' �ؿܹ�ۿ��� 2016/05/24 �߰�

For j=LBound(splitedSeq) to UBound(splitedSeq)
    ixsiteOrderSerial = Trim(splitedSeq(j))

    if (ixsiteOrderSerial<>"") then

        set otmpOrder = new CxSiteTempOrder
        otmpOrder.FPageSize = 200
        otmpOrder.FCurrPage = 1
        otmpOrder.FRectOutMallOrderSerial = ixsiteOrderSerial
        otmpOrder.FRectMatchState ="I" '''
        'otmpOrder.getOnlineTmpOrderList(false)		'2017-01-11 10:06 ������ �ּ�ó��
        otmpOrder.getOnlineTmpOrderRealInputList()

        rw otmpOrder.FItemList(0).FOutMallOrderSerial

        countryCode = otmpOrder.FItemList(0).fcountryCode
		if countryCode="" then ucase(countryCode)="KR"
        IsForeignDLV = (ucase(countryCode)<>"KR")

        ErrMsg = "[001]"

        dbget.beginTrans
        	'�ֹ��Է�(������)
        	sqlStr = "select * from [db_order].[dbo].tbl_order_master where 1=0"
        	rsget.Open sqlStr,dbget,1,3
        	rsget.AddNew
        	rsget("orderserial") = Left(Left(otmpOrder.FItemList(0).FSellSite,2)&otmpOrder.FItemList(0).FOutMallOrderSerial ,11)

			rsget("reqemail") = otmpOrder.FItemList(0).forderemail
			rsget("jumundiv") = "5"
        	rsget("userid") = ""
        	rsget("ipkumdiv") = "1"
        	rsget("accountname") = ""
        	rsget("accountdiv") = "50"
        	rsget("authcode") = ixsiteOrderSerial
        	rsget("sitename") = otmpOrder.FItemList(0).FSellSite
        	rsget("DlvcountryCode") = countryCode
	''2017-01-06 ������..2017-01-09�� ���Ŀ� �ؿܶ�� Fbeadaldiv�� 80 �ƴϸ� otmpOrder.FItemList(0).Fbeadaldiv ���� �����ؾ���..
	''����� [db_temp].[dbo].[sp_TEN_xSiteTmpOrderList]�� ���ν������� �����ϴ��� �ƴϸ� ���� asp���Ͽ��� �����ϴ���
        	rsget("beadaldiv") = otmpOrder.FItemList(0).Fbeadaldiv
        	rsget.update
        	iid = rsget("idx")
        	rsget.close

        	orderserial = Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),3,256)
        	orderserial = orderserial & Format00(5,Right(CStr(iid),5))

            if Err then
                dbget.RollBackTrans
                response.write ErrMsg & Err.Description
                response.end
            else
                ErrMsg = "[002]"
            end if

        	sqlStr = "update M" & vbCrlf
            sqlStr = sqlStr + " set orderserial='" + CStr(orderserial) + "'," & vbCrlf
            sqlStr = sqlStr + " accountname='" + html2db(otmpOrder.FItemList(0).FOrderName) + "'," & vbCrlf
            sqlStr = sqlStr + " totalsum=0," & vbCrlf
            sqlStr = sqlStr + " ipkumdiv='4'," & vbCrlf
'			sqlStr = sqlStr + " ipkumdate=getdate()," & vbCrlf		'�Ա����� tmporder�� paydate�� ����..2021-12-07 ������
			sqlStr = sqlStr + " ipkumdate='" & dateconvert(otmpOrder.FItemList(0).FPaydate) & "'," & vbCrlf
            sqlStr = sqlStr + " regdate=getdate()," & vbCrlf
            ''sqlStr = sqlStr + " beadaldiv='1'," & vbCrlf
            sqlStr = sqlStr + " buyname='" + html2db(otmpOrder.FItemList(0).FOrderName) + "'," & vbCrlf
            sqlStr = sqlStr + " buyphone='" + replace(otmpOrder.FItemList(0).FOrderTelNo,"'","") + "'," & vbCrlf
            sqlStr = sqlStr + " buyhp='" + replace(otmpOrder.FItemList(0).FOrderHpNo,"'","") + "'," & vbCrlf
            sqlStr = sqlStr + " buyemail=''," & vbCrlf
            sqlStr = sqlStr + " reqname='" + html2db(otmpOrder.FItemList(0).FReceiveName) + "'," & vbCrlf

            if ucase(countryCode)="KR" then
                sqlStr = sqlStr + " reqzipcode='" + Trim(otmpOrder.FItemList(0).FReceiveZipCode) + "'," & vbCrlf
            else
            	sqlStr = sqlStr + " reqzipcode='00000'," & vbCrlf
        	end if

            sqlStr = sqlStr + " reqaddress='" + TRIM(html2db(otmpOrder.FItemList(0).FReceiveAddr2)) + "'," & vbCrlf
            sqlStr = sqlStr + " reqphone='" + replace(otmpOrder.FItemList(0).FReceiveTelNo,"'","") + "'," & vbCrlf
            sqlStr = sqlStr + " reqhp='" + replace(otmpOrder.FItemList(0).FReceiveHpNo,"'","") + "'," & vbCrlf
'			sqlStr = sqlStr + " comment='" + replace(TRIM(html2db(otmpOrder.FItemList(0).Fdeliverymemo)), "'", "") + "'," & vbCrlf
			sqlStr = sqlStr + " comment='" & TRIM(html2db(otmpOrder.FItemList(0).Fdeliverymemo)) & "'," & vbCrlf		'2021-12-20 ������ replace����
            sqlStr = sqlStr + " discountrate=1," & vbCrlf
            sqlStr = sqlStr + " subtotalprice=0," & vbCrlf
            sqlStr = sqlStr + " reqzipaddr='" + html2db(otmpOrder.FItemList(0).FReceiveAddr1) + "'" & vbCrlf
            sqlStr = sqlStr + " From [db_order].[dbo].tbl_order_master M" & vbCrlf
            sqlStr = sqlStr + " where idx=" + CStr(iid)

            sqlStr = Replace(sqlStr, CHr(230), "")
        	dbget.Execute sqlStr

            if Err then
                dbget.RollBackTrans
                response.write ErrMsg & Err.Description
                response.end
            else
                ErrMsg = "[003]"
            end if

			buf_totcost = 0
			buf_totvat = 0
			buf_totCpnNotAppliedcost =0
			buf_iitemmakerid = ""

'			IF (otmpOrder.FItemList(0).FSellSite="wemakeprice") then
'				''����ũ �����̽� ��ۺ� 0 =������ �Է�.
'				sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
'				sqlStr = sqlStr + "itemoption, makerid, itemno, itemcost, buycash, itemvat, mileage, itemname, itemoptionname, reducedPrice)" & vbCrlf
'				sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
'				sqlStr = sqlStr + " '" & orderserial & "'," & vbCrlf
'				sqlStr = sqlStr + " 0," & vbCrlf
'				sqlStr = sqlStr + " '0501'," & vbCrlf
'				sqlStr = sqlStr + " ''," & vbCrlf
'				sqlStr = sqlStr + " 1," & vbCrlf
'				sqlStr = sqlStr + "	0," & vbCrlf
'				sqlStr = sqlStr + "	0," & vbCrlf
'				sqlStr = sqlStr + "	0,"
'				sqlStr = sqlStr + "	0,"
'				sqlStr = sqlStr + "	'','',"
'				sqlStr = sqlStr + "	0" & vbCrlf
'				sqlStr = sqlStr + " )"
'				dbget.Execute sqlStr
'			end IF

			if (otmpOrder.FItemList(0).FSellSite="dnshop") then
				'2015-03-12������ �ϴ� if�� ���� �ּ� �� 5���� �̻� ���������� ����
				''5���� �̻� ������.
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

'				if (otmpOrder.FItemList(0).ForderDlvPay=-1) then
'					''������ȣ4 ������
'					otmpOrder.FItemList(0).ForderDlvPay=0
'				else
'					''5���� �̻� ������.
'					otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
'				end if
			elseIF (otmpOrder.FItemList(0).FSellSite="interpark") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cjmall") then     ''skip
			elseIF (otmpOrder.FItemList(0).FSellSite="lotteCom") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

			elseIF (otmpOrder.FItemList(0).FSellSite="lotteimall") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

			elseIF (otmpOrder.FItemList(0).FSellSite="bandinlunis") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

			elseIF (otmpOrder.FItemList(0).FSellSite="its29cm") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

			elseIF (otmpOrder.FItemList(0).FSellSite="ithinksoshop") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)

            elseIF (otmpOrder.FItemList(0).FSellSite="ssg") then    ''2018/02/28 �߰�.
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (otmpOrder.FItemList(0).FSellSite="11stITS") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cookatmall") then
			elseIF (otmpOrder.FItemList(0).FSellSite="11st1010") then
			elseIF (otmpOrder.FItemList(0).FSellSite="ezwel") then
			elseIF (otmpOrder.FItemList(0).FSellSite="boribori1010") then
			elseIF (otmpOrder.FItemList(0).FSellSite="auction1010") then
			elseIF (otmpOrder.FItemList(0).FSellSite="gmarket1010") then
			elseIF (otmpOrder.FItemList(0).FSellSite="nvstorefarm") then
			elseIF (otmpOrder.FItemList(0).FSellSite="nvstoremoonbangu") then
			elseIF (otmpOrder.FItemList(0).FSellSite="Mylittlewhoopee") then
			elseIF (otmpOrder.FItemList(0).FSellSite="nvstoregift") then
			elseIF (otmpOrder.FItemList(0).FSellSite="nvstorefarmclass") then
			elseIF (otmpOrder.FItemList(0).FSellSite="WMP") then
			elseIF (otmpOrder.FItemList(0).FSellSite="wmpfashion") then
			elseIF (otmpOrder.FItemList(0).FSellSite="GS25") then
			elseIF (otmpOrder.FItemList(0).FSellSite="thinkaboutyou") then
			elseIF (otmpOrder.FItemList(0).FSellSite="aboutpet") then
			elseIF (otmpOrder.FItemList(0).FSellSite="momQ") then
			elseIF (otmpOrder.FItemList(0).FSellSite="giftting") then
			elseIF (otmpOrder.FItemList(0).FSellSite="kakaogift") then
			elseIF (otmpOrder.FItemList(0).FSellSite="itskakaotalkstore") then
			elseIF (otmpOrder.FItemList(0).FSellSite="itskakao") then
				otmpOrder.FItemList(0).ForderDlvPay = 0
			elseIF (otmpOrder.FItemList(0).FSellSite="coupang") then
			elseIF (otmpOrder.FItemList(0).FSellSite="hmall1010") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="lfmall") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="kakaostore") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="wconcept1010") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="benepia1010") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="withnature1010") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="goodshop1010") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="lotteon") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="shintvshopping") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="goodwearmall10") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="skstoa") then
			elseIF (LCASE(otmpOrder.FItemList(0).FSellSite)="wetoo1300k") then
			elseIF (otmpOrder.FItemList(0).FSellSite="yes24") then
			elseIF (otmpOrder.FItemList(0).FSellSite="alphamall") then
			elseIF (otmpOrder.FItemList(0).FSellSite="ohou1010") then
			elseIF (otmpOrder.FItemList(0).FSellSite="casamia_good_com") then
			elseIF (otmpOrder.FItemList(0).FSellSite="wadsmartstore") then
			elseIF (otmpOrder.FItemList(0).FSellSite="privia") then
			elseIF (otmpOrder.FItemList(0).FSellSite="NJOYNY") or (otmpOrder.FItemList(0).FSellSite="itsNJOYNY") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cn10x10") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cnglob10x10") then
			elseIF (otmpOrder.FItemList(0).FSellSite="11stmy") then
			elseIF (otmpOrder.FItemList(0).FSellSite="shopify") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cnhigo") then
			elseIF (otmpOrder.FItemList(0).FSellSite="etsy'") then
			elseIF (otmpOrder.FItemList(0).FSellSite="zilingo") then
			elseIF (otmpOrder.FItemList(0).FSellSite="cnugoshop") then
			elseIF (otmpOrder.FItemList(0).FSellSite="wemakeprice") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (otmpOrder.FItemList(0).FSellSite="cjmallITS") or (otmpOrder.FItemList(0).FSellSite="itsCjmall") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (otmpOrder.FItemList(0).FSellSite="fashionplus") or (otmpOrder.FItemList(0).FSellSite="itsFashionplus") then
				'//otmpOrder.FItemList(0).ForderDlvPay ������ ��۷� �ֽ�
				'if (otmpOrder.FItemList(0).ForderDlvPay=0) then
					''3���� �̻� ������.
					'otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
				'end if

			elseIF (otmpOrder.FItemList(0).FSellSite="byulshopITS") or (otmpOrder.FItemList(0).FSellSite="itsByulshop") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (otmpOrder.FItemList(0).FSellSite="gabangpop") or (otmpOrder.FItemList(0).FSellSite="itsGabangpop") then
			elseIF (otmpOrder.FItemList(0).FSellSite="musinsaITS") or (otmpOrder.FItemList(0).FSellSite="itsMusinsa") then
			elseIF (otmpOrder.FItemList(0).FSellSite="mintstore") or (otmpOrder.FItemList(0).FSellSite="itsMintstore") then
			elseIF (otmpOrder.FItemList(0).FSellSite="gseshop") then ''2013/08/05
			elseIF (otmpOrder.FItemList(0).FSellSite="itsbenepia") then
			elseif (otmpOrder.FItemList(0).FSellSite="itsKaKaoMakers") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseif (otmpOrder.FItemList(0).FSellSite="itsWadiz") then
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			elseIF (otmpOrder.FItemList(0).FSellSite="gsshop") then
			'''elseIF (otmpOrder.FItemList(0).FSellSite="hottracks") then     ����Ŀ������� ������ �ּ�����
			else    '' �׿ܸ� 3�����̻� ����
				otmpOrder.FItemList(0).ForderDlvPay=otmpOrder.getDlvPayBySubPrice(otmpOrder.FItemList(0).FSellSite)
			end if

			'// ��ۺ� �Է��� ��ǰ�Է��� ��� ���� �ڿ� �Ѵ�.
			For i=0 to otmpOrder.FResultCount-1

				if (FALSE) and (otmpOrder.FItemList(i).FmatchItemID=0) then
					'// ��ۺ�
				else
					sqlStr= "select top 1 convert(int, i.sellcash) as sellcash, " & vbCrlf
					sqlStr = sqlStr + " i.mileage, i.itemdiv , convert(int,i.buycash) as buycash ," & vbCrlf
					sqlStr = sqlStr + " convert(int, i.orgprice) as orgprice, convert(int,i.orgsuplycash) as orgsuplycash ," & vbCrlf
					sqlStr = sqlStr + " i.itemname, i.makerid, i.vatinclude, i.deliverytype, i.sailyn, i.mwdiv,"
					sqlStr = sqlStr + " IsNull(v.optionname,'') as codeview, IsNull(v.optaddbuyprice,0) as optaddbuyprice" & vbCrlf
					sqlStr = sqlStr + " , IsNull(v.optaddprice,0) as optaddprice" & vbCrlf
					sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i WITH(NOLOCK) " & vbCrlf
					sqlStr = sqlStr + "     left join [db_item].[dbo].tbl_item_option v WITH(NOLOCK) "
					sqlStr = sqlStr + "     on (i.itemid=v.itemid) and (v.itemoption='" + CStr(otmpOrder.FItemList(i).FmatchItemOption) + "')" & vbCrlf
					sqlStr = sqlStr + " where i.itemid = " + CStr(otmpOrder.FItemList(i).FmatchItemID) + ""

					rsget.CursorLocation = adUseClient
    				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					if  not rsget.EOF  then
						buf_deliverytype = rsget("deliverytype")
						if (buf_deliverytype="2") or (buf_deliverytype="5")  or (buf_deliverytype="9") or (buf_deliverytype="7") then
							t_upchebeasong="Y"
						else
							t_upchebeasong="N"
						end if

                        buf_CurSellcash = rsget("sellcash")
						if ucase(countryCode)="KR" then
							buf_sellcash   = buf_CurSellcash
						else
							buf_sellcash = otmpOrder.FItemList(0).FSellPrice
						end if


						buf_sellvat         = CLng(buf_sellcash*11/10)-CLng(buf_sellcash)  ''�ǹ̾���.
						buf_mileage         = rsget("mileage")
						buf_itemdiv         = rsget("itemdiv")
						buf_iitemname       = replace(rsget("itemname"),"'","")
						buf_iitemoptionname = replace(rsget("codeview"),"'","")
						buf_onlyitembuycash = rsget("buycash")
						buf_onlyoptaddbuyprice = rsget("optaddbuyprice")
						buf_onlyoptaddprice     = rsget("optaddprice")
						buf_iitembuycash    = buf_onlyitembuycash + buf_onlyoptaddbuyprice
						buf_iitemmakerid    = rsget("makerid")
						buf_iitemvatinclude = rsget("vatinclude")
						buf_mwdiv           = rsget("mwdiv")

						buf_sailyn          = rsget("sailyn")
						buf_orgprice        = rsget("orgprice")
						buf_orgsuplycash    = rsget("orgsuplycash")

						buf_CurSellcash = buf_CurSellcash + buf_onlyoptaddprice

					end if
					rsget.close

                    if (IsForeignDLV) and (t_upchebeasong="Y") then ''�ؿܹ���� ������ ��ü���ó�� 2016/06/07
                        Call FnChangForeignMallUpbeaDLV(otmpOrder.FItemList(i).FmatchItemID,otmpOrder.FItemList(i).FmatchItemOption,buf_iitemmakerid,t_upchebeasong,buf_mwdiv,buf_iitembuycash,buf_orgprice,buf_orgsuplycash)

					end if

					''�����Ǹ��ΰ�� üũ 20100201 �߰� / ���� ���̺� ���� �ִ°��..
					''================================================================
'���� buycash ��������
'1. otmpOrder.FItemList(0).FSellSite=�� ���޸� �߰�
'2. select top 5 itemid,saleprice,salesupplycash from db_event.dbo.tbl_saleitem�� �ɸ��� �ʾҴٸ� buf_sailyn <> 'Y'������ �ɷ��� Q�� ����
'   history���̺� �ں��� 2�ֳ��� regdate�� Top1������ ������..�����������̴� desc�ؾ߰���

					'if (CLng(buf_sellcash) > CLng(otmpOrder.FItemList(i).FSellPrice))  then  ''�����Ǹ��� ��� FRealSellPrice=>
					If CLng(buf_sellcash) + CLng(buf_onlyoptaddprice) > CLng(otmpOrder.FItemList(i).FSellPrice) Then
'rw "maysale:"&buf_sellcash&":"&otmpOrder.FItemList(i).FSellPrice
						if ((otmpOrder.FItemList(0).FSellSite="dnshop") or (otmpOrder.FItemList(0).FSellSite="interpark") or (otmpOrder.FItemList(0).FSellSite="cjmall") or (otmpOrder.FItemList(0).FSellSite="lotteCom") or (otmpOrder.FItemList(0).FSellSite="lotteimall")  or (otmpOrder.FItemList(0).FSellSite="gmarket1010") or (otmpOrder.FItemList(0).FSellSite="auction1010") or (otmpOrder.FItemList(0).FSellSite="boribori1010") or (otmpOrder.FItemList(0).FSellSite="gseshop") or (otmpOrder.FItemList(0).FSellSite="11st1010") or (otmpOrder.FItemList(0).FSellSite="nvstorefarm") or (otmpOrder.FItemList(0).FSellSite="nvstoremoonbangu") or (otmpOrder.FItemList(0).FSellSite="Mylittlewhoopee") or (otmpOrder.FItemList(0).FSellSite="nvstoregift") or (otmpOrder.FItemList(0).FSellSite="ssg") or (otmpOrder.FItemList(0).FSellSite="halfclub") or (otmpOrder.FItemList(0).FSellSite="hmall1010") or (otmpOrder.FItemList(0).FSellSite="coupang") or (otmpOrder.FItemList(0).FSellSite="WMP") or (otmpOrder.FItemList(0).FSellSite="wmpfashion") or (otmpOrder.FItemList(0).FSellSite="lotteon")) then
							mayOrderDate = otmpOrder.FItemList(i).FSelldate
							if isNULL(mayOrderDate) then mayOrderDate=LEFT(CStr(NOW()),10)
'rw "mayOrderDate:"&mayOrderDate
							if IsDate(mayOrderDate) then
								sqlStr= " select top 5 itemid,saleprice,salesupplycash from db_event.dbo.tbl_saleitem WITH(NOLOCK) "  & vbCrlf
								sqlStr = sqlStr + " where itemid=" & CStr(otmpOrder.FItemList(i).FmatchItemID)  & vbCrlf
								sqlStr = sqlStr + " and convert(varchar(10),opendate,21)<='"&mayOrderDate&"'"  & vbCrlf
								sqlStr = sqlStr + " and convert(varchar(10),IsNULL(closedate,'2099-12-31'),21)>='"&mayOrderDate&"'"  & vbCrlf
								sqlStr = sqlStr + " order by saleitem_idx desc"  & vbCrlf
'rw sqlStr
								rsget.CursorLocation = adUseClient
        						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
								if  not rsget.EOF  then
								    do until rsget.Eof
'rw CLng(otmpOrder.FItemList(i).FSellPrice) + CLng(buf_onlyoptaddprice)
    									if CLng(rsget("saleprice"))=CLng(otmpOrder.FItemList(i).FSellPrice) - CLng(buf_onlyoptaddprice) then   ''' -�μ��� 2017/03/24
    										buf_onlyitembuycash = rsget("salesupplycash")
    										buf_iitembuycash    = buf_onlyitembuycash + buf_onlyoptaddbuyprice
    										buf_sailyn   = "Y"

    										exit Do
    									end if
									    rsget.moveNext
									loop
								end if
								rsget.close
							end if
						end if

''						''�߱�����Ʈ ���� ���� //2013/07/29
''						if (otmpOrder.FItemList(0).FSellSite="cn10x10") then
''						    mayOrderDate = otmpOrder.FItemList(i).FSelldate
''						    if IsDate(mayOrderDate) then
''						        sqlStr= " select top 1 itemid,discountPrice as saleprice,I.discountBuyMoney as salesupplycash  "
''                                sqlStr = sqlStr + " 	from db_item.dbo.tbl_kaffa_Discount_List L"
''                                sqlStr = sqlStr + " 	Join db_item.dbo.tbl_kaffa_Discount_Item I"
''                                sqlStr = sqlStr + " 	on L.discountKey=I.discountKey"
''                                sqlStr = sqlStr + " where I.itemid=" & CStr(otmpOrder.FItemList(i).FmatchItemID)  & vbCrlf
''                                sqlStr = sqlStr + " and L.openDate<='"&mayOrderDate&"'"
''                                sqlStr = sqlStr + " and L.openDate is Not NULL"
''                                sqlStr = sqlStr + " and convert(varchar(10),IsNULL(L.expiredDate,'2099-12-31'),21)>='"&mayOrderDate&"' "
''                                sqlStr = sqlStr + " and convert(varchar(10),IsNULL(I.expiredDate,'2099-12-31'),21)>='"&mayOrderDate&"' "
''                                sqlStr = sqlStr + " order by L.discountKey desc"
''
''                                rsget.Open sqlStr, dbget, 1
''								if  not rsget.EOF  then
''									if CLng(rsget("saleprice"))=CLng(otmpOrder.FItemList(i).FRealSellPrice) then ''�����ǸŰ��� ������.
''										buf_onlyitembuycash = rsget("salesupplycash")
''										buf_iitembuycash    = buf_onlyitembuycash + buf_onlyoptaddbuyprice
''										buf_sailyn   = "Y"
''									end if
''								end if
''								rsget.close
''						    end if
''						end if

					elseif (CLng(buf_sellcash) < CLng(otmpOrder.FItemList(i).FSellPrice)) then ''��ΰ� �ȸ���� and (buf_sailyn="Y")
						if (CLng(otmpOrder.FItemList(i).FSellPrice)>=buf_orgprice+buf_onlyoptaddprice) then  ''�Һ��ڰ��� ���ų� ũ��
'rw 	"buf_orgprice+buf_onlyoptaddprice"&buf_orgprice+buf_onlyoptaddprice
'rw 	"buf_orgsuplycash+buf_onlyoptaddbuyprice"&buf_orgsuplycash+buf_onlyoptaddbuyprice
							buf_iitembuycash = buf_orgsuplycash + buf_onlyoptaddbuyprice
							buf_sailyn   = "N"
						end if
					end if
					''================================================================
					if Err then
						dbget.RollBackTrans
						response.write ErrMsg & Err.Description
						response.end
					else
						ErrMsg = "[003.1]"
					end if

                    ''buf_CpnNotAppliedSellcash �߰�
                    ''�ʱ�ȭ �߰� 2014/10/02 (0���ΰ�� �������尨)
                    buf_sellcash = 0
                    buf_CpnNotAppliedSellcash = 0
                    buf_CpnNotAppliedSellcash = 0

                    if otmpOrder.FItemList(i).FRealSellPrice<>0 then
						buf_sellcash = otmpOrder.FItemList(i).FRealSellPrice
						buf_CpnNotAppliedSellcash = buf_sellcash
					end if

                    if otmpOrder.FItemList(i).FSellPrice<>0 then
                        buf_CpnNotAppliedSellcash = otmpOrder.FItemList(i).FSellPrice
                    end if

					buf_totcost = buf_totcost + CLng(buf_sellcash) * CLng(otmpOrder.FItemList(i).FItemOrderCount)
					buf_totvat  = buf_totvat + CLng(buf_sellvat) * CLng(otmpOrder.FItemList(i).FItemOrderCount)
                    buf_totCpnNotAppliedcost = buf_totCpnNotAppliedcost + CLng(buf_CpnNotAppliedSellcash) * CLng(otmpOrder.FItemList(i).FItemOrderCount)

					'##########################################################################
					'2021-03-24 ������ �߰�
					'Ư�������� ��ǰ�� ������ ���԰� ������Ʈ �ϱ�
					Dim mustBuyPrice
					mustBuyPrice = ""

					sqlStr = ""
					sqlStr = sqlStr & " SELECT TOP 1 isnull(mustBuyPrice, 0) as mustBuyPrice "
					sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
					sqlStr = sqlStr & " WHERE GETDATE() BETWEEN startDate and endDate "
					sqlStr = sqlStr & " and mustPrice = '"& CStr(buf_CpnNotAppliedSellcash - buf_onlyoptaddprice) &"' " 'Ư���� �ǸŰ� - �ɼ��߰��ݾ��̶� ����
					sqlStr = sqlStr & " and mallgubun = '"& otmpOrder.FItemList(0).FSellSite &"' "
					sqlStr = sqlStr & " and itemid = '"& CStr(otmpOrder.FItemList(i).FmatchItemID) &"' "
					sqlStr = sqlStr & " and isnull(mustBuyPrice, 0) > 0 "
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					If not rsget.EOF Then
						mustBuyPrice = rsget("mustBuyPrice")
					Else
						mustBuyPrice = ""
					End If
					rsget.Close

					If mustBuyPrice <> "" Then
						buf_iitembuycash = mustBuyPrice + buf_onlyoptaddbuyprice
					End If
					'##########################################################################

					sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
					sqlStr = sqlStr + "itemoption, itemno, itemcost, itemvat, mileage, reducedPrice, " & vbCrlf
					sqlStr = sqlStr + "orgitemcost,itemcostcouponnotApplied,bonuscouponidx,buycashcouponNotApplied, " & vbCrlf
					sqlStr = sqlStr + "itemname,itemoptionname,makerid,buycash," & vbCrlf
					sqlStr = sqlStr + "vatinclude,isupchebeasong,issailitem,oitemdiv,omwdiv,odlvType,requiredetail)" & vbCrlf
					sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
					sqlStr = sqlStr + " '" + orderserial + "'," & vbCrlf
					sqlStr = sqlStr + " " + CStr(otmpOrder.FItemList(i).FmatchItemID) + "," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(otmpOrder.FItemList(i).FmatchItemOption) + "'," & vbCrlf
					sqlStr = sqlStr + " " + CStr(otmpOrder.FItemList(i).FItemOrderCount) + "," & vbCrlf
					sqlStr = sqlStr + " " + CStr(buf_CpnNotAppliedSellcash) + "," & vbCrlf  '' buf_sellcash => buf_CpnNotAppliedSellcash ����
					sqlStr = sqlStr + " " + CStr(buf_sellvat) + "," & vbCrlf
					sqlStr = sqlStr + " 0," & vbCrlf
					sqlStr = sqlStr + " " + CStr(buf_sellcash) + "," & vbCrlf               '' reducedPrice
					sqlStr = sqlStr + " " + CStr(buf_orgprice+buf_onlyoptaddprice) + "," & vbCrlf   ''buf_onlyoptaddprice �߰� 2015/05/18

					if (IsForeignDLV) then ''�б�ó�� 2016/05/25
					    sqlStr = sqlStr + " " + CStr(buf_CurSellcash) + "," & vbCrlf            ''���� �ǸŰ�
					else
					    sqlStr = sqlStr + " " + CStr(buf_CpnNotAppliedSellcash) + "," & vbCrlf  ''����
				    end if

					if (buf_CpnNotAppliedSellcash>buf_sellcash) then ''���ʽ���������
					    sqlStr = sqlStr + "-1,"
					else
					    sqlStr = sqlStr + "NULL,"
				    end if
					sqlStr = sqlStr + " " + CStr(buf_iitembuycash) + "," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(buf_iitemname) + "'," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(buf_iitemoptionname) + "'," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(buf_iitemmakerid) + "'," & vbCrlf
					sqlStr = sqlStr + " " + CStr(buf_iitembuycash) + "," & vbCrlf
					sqlStr = sqlStr + " '" + CStr(buf_iitemvatinclude) + "'," & vbCrlf
					sqlStr = sqlStr + " '" + t_upchebeasong + "'," & vbCrlf
					sqlStr = sqlStr + " '" + buf_sailyn + "'," & vbCrlf
					sqlStr = sqlStr + " '" + buf_itemdiv + "'," & vbCrlf
					sqlStr = sqlStr + " '" + buf_mwdiv + "'," & vbCrlf
					sqlStr = sqlStr + " '" + buf_deliverytype + "'," & vbCrlf
					sqlStr = sqlStr + " '" + replace(CStr(otmpOrder.FItemList(i).FrequireDetail),"'","''") + "'" & vbCrlf
					sqlStr = sqlStr + " )"
					dbget.Execute sqlStr
				end if
			next
			'// ========================================================================
			'' ���޸� �ֹ� ��ۺ� �Է�
			''
			''  - �ǸŰ�
			''    - ���޸��� ��ۺ� ���Ѵ�.(otmpOrder.FItemList(0).ForderDlvPay)
			''
			''  - �Ѱ��� �ǸŰ� ����Ѵ�.
			''  - ������ ��ǰ(odlvtype = ��ü������:2, ����:7, ���ٹ�����:4)
			''  - ��ü���ǹ��:9
			''
			''  - �ٹ���� �ְų� ��ü�ֹ� ��ü�� ������(���ǹ�� ��ǰ�� ���� ���)�̸� �귣�� ���� ��ۺ� �Է��Ѵ�.
			''    - �ǸŰ�=���԰�
			''
			''  - ��ü��� �귣���̸鼭 ���ǹ�ۻ�ǰ�� �ִ� ���
			''    - ��ۺ� �Է��� ������ ��� : ù��° �귣�忡 ��ۺ� �ǸŰ� �Է�(�� �̿� �ǸŰ� 0��)
			''
			''    - ������ ��ǰ �ִ� ��� : ���԰� 0��
			''
			''    - ������ ��ǰ ���� ���
			''      - �귣�� �����Ѿװ� ���ǹ�ۺ� ���Ѵ�. ������ ������ �����ϸ� ���԰� 0��. ��Ÿ ���ǹ�ۺ� ���԰�.
			'// ========================================================================

			dim IsTenbeaItemExist			: IsTenbeaItemExist = False
			dim IsBeasongSellPriceInserted	: IsBeasongSellpriceInserted = False
			dim IsAllFreeUpcheBeasong		: IsAllFreeUpcheBeasong = True
			dim UpcheBeasongOptionIdx		: UpcheBeasongOptionIdx = 0
			dim IsUpchePaticleBeasongExists : IsUpchePaticleBeasongExists = False

			dim arrtotitemcost, arrmakerid, arrfreebeasongcount, arrupchebeasongcondcount, arrdefaultFreeBeasongLimit, arrdefaultDeliverPay, arrdefaultDeliveryType
            dim arritemcostcouponnotApplied, t_recordCNT ''2016/05/25
			'response.write orderserial

			sqlStr = "select totitemcost,totitemcostcouponnotApplied,makerid,freebeasongcount,upchebeasongcondcount,defaultFreeBeasongLimit,defaultDeliverPay,defaultDeliveryType "
			sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + " select "
			sqlStr = sqlStr + " 	sum(d.itemcost*d.itemno) as totitemcost "
			sqlStr = sqlStr + " 	, sum(d.itemcostcouponnotApplied*d.itemno) as totitemcostcouponnotApplied " ''�����ǸŰ�(�ؿܸ�)..
			sqlStr = sqlStr + " 	, (case when d.isupchebeasong = 'Y' then d.makerid else '' end) as makerid "
			sqlStr = sqlStr + " 	, sum(case when IsNull(d.odlvtype, 0) in (2) then 1 else 0 end) as freebeasongcount "  ''7 ����, 4�ٹ���, 2��ü���� '' , 4, 7 ���� ''2016/05/25
			sqlStr = sqlStr + " 	, sum(case when IsNull(d.odlvtype, 0) in (9) then 1 else 0 end) as upchebeasongcondcount "
			sqlStr = sqlStr + " 	, max(c.defaultFreeBeasongLimit) as defaultFreeBeasongLimit, max(c.defaultDeliverPay) as defaultDeliverPay, max(c.defaultDeliveryType) as defaultDeliveryType "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	[db_order].[dbo].tbl_order_detail d WITH(NOLOCK) "
			sqlStr = sqlStr + " 	left join [db_user].[dbo].tbl_user_c c WITH(NOLOCK) "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		d.makerid = c.userid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and orderserial = '" + CStr(orderserial) + "' "
			sqlStr = sqlStr + " 	and itemid <> 0 "
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	(case when isupchebeasong = 'Y' then makerid else '' end) "
			sqlStr = sqlStr + " ) T"
			sqlStr = sqlStr + " order by (case when makerid<>'' then 0 else 1 end),(CASE WHEN freebeasongcount=0 then 0 ELSE 1 END), (CASE WHEN upchebeasongcondcount>0 then 0 ELSE 1 END), (CASE WHEN defaultFreeBeasongLimit-totitemcost>0 THEN 0 ELSE 1 END) ,defaultDeliverPay desc"
			''���ı��� ���� 2020/01/28
			''sqlStr = sqlStr + " 	(case when isupchebeasong = 'Y' then makerid else '' end) "

			rsget.CursorLocation = adUseClient
    		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

            t_recordCNT = rsget.RecordCount
			redim arrtotitemcost(t_recordCNT)
			redim arrmakerid(t_recordCNT)
			redim arrfreebeasongcount(t_recordCNT)
			redim arrupchebeasongcondcount(t_recordCNT)
			redim arrdefaultFreeBeasongLimit(t_recordCNT)
			redim arrdefaultDeliverPay(t_recordCNT)
			redim arrdefaultDeliveryType(t_recordCNT)
            redim arritemcostcouponnotApplied(t_recordCNT)

			if  not rsget.EOF  then
				i = 0
				do until rsget.eof
					arrtotitemcost(i)				= rsget("totitemcost")
					arrmakerid(i)					= rsget("makerid")
					arrfreebeasongcount(i)			= rsget("freebeasongcount")			'// �����ۻ�ǰ
					arrupchebeasongcondcount(i)		= rsget("upchebeasongcondcount")	'// ���ǹ�ۻ�ǰ

					arrdefaultFreeBeasongLimit(i)	= rsget("defaultFreeBeasongLimit")
					arrdefaultDeliverPay(i)			= rsget("defaultDeliverPay")

					arritemcostcouponnotApplied(i)  = rsget("totitemcostcouponnotApplied")

					''��ǰ�������� �Ѵ�.
					''arrdefaultDeliveryType(i)		= rsget("defaultDeliveryType")

					if (arrmakerid(i) = "") then
						IsTenbeaItemExist = True
					elseif (arrupchebeasongcondcount(i) > 0) then  ''��ü���ǹ���� �����Ѵ�.
						IsAllFreeUpcheBeasong = False
						IsUpchePaticleBeasongExists = True
					end if

					rsget.MoveNext
					i = i + 1
				loop
			end if
			rsget.close

			'// �����Ѿ׿� ��ۺ� �߰�
			buf_totcost = buf_totcost + otmpOrder.FItemList(0).ForderDlvPay
			buf_totvat = buf_totvat + 0
			buf_totCpnNotAppliedcost = buf_totCpnNotAppliedcost + otmpOrder.FItemList(0).ForderDlvPay

			'' ���� ��ۺ� ���� �Է����� ���� //2020/01/28
			Dim addTenDlvPay : addTenDlvPay = 0
			UpcheBeasongOptionIdx = 0
			for i = 0 to UBound(arrtotitemcost)
				if (arrmakerid(i) <> "") and (arrupchebeasongcondcount(i) > 0) then

					UpcheBeasongOptionIdx = UpcheBeasongOptionIdx + 1

					if (Not IsBeasongSellPriceInserted) then

						IsBeasongSellPriceInserted = True

						sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
						sqlStr = sqlStr + "itemoption, makerid, itemno, itemcost, reducedPrice, orgitemcost, itemcostCouponNotApplied, buycash, buycashCouponNotApplied, itemvat, mileage, itemname, itemoptionname)" & vbCrlf
						sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
						sqlStr = sqlStr + " '" & orderserial & "'," & vbCrlf
						sqlStr = sqlStr + " 0," & vbCrlf
						sqlStr = sqlStr + " '" & ("9" & Format00(3, UpcheBeasongOptionIdx)) & "'," & vbCrlf
						sqlStr = sqlStr + " '" + CStr(arrmakerid(i)) + "'," & vbCrlf
						sqlStr = sqlStr + " 1," & vbCrlf

                        if (IsForeignDLV) then  ''�ؿܹ�� �б�
                            if (arrfreebeasongcount(i) > 0) or (arritemcostcouponnotApplied(i) >= arrdefaultFreeBeasongLimit(i)) then
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcost
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''reducedPrice
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''orgitemcost
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcostCouponNotApplied
    							'// �����ۻ�ǰ�� �ְų� ���ǹ�ۺ� �̻�(���� �ʰ�����)
    							sqlStr = sqlStr + "	0," & vbCrlf
    							sqlStr = sqlStr + "	0," & vbCrlf
								sqlStr = sqlStr + "	" & CLng(otmpOrder.FItemList(0).ForderDlvPay*1/11) & ","
    						else
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcost
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''reducedPrice
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''orgitemcost
								sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcostCouponNotApplied
    							''���԰�
    							sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycash
    							sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycashCouponNotApplied
								sqlStr = sqlStr + "	" & CLng(otmpOrder.FItemList(0).ForderDlvPay*1/11) & ","
    						end if

                        else

    						if (arrfreebeasongcount(i) > 0) or (arrtotitemcost(i) >= arrdefaultFreeBeasongLimit(i)) then
								addTenDlvPay = otmpOrder.FItemList(0).ForderDlvPay-0

								sqlStr = sqlStr + "	" & 0 & "," & vbCrlf    ''itemcost
								sqlStr = sqlStr + "	" & 0 & "," & vbCrlf    ''reducedPrice
								sqlStr = sqlStr + "	" & 0 & "," & vbCrlf    ''orgitemcost
								sqlStr = sqlStr + "	" & 0 & "," & vbCrlf    ''itemcostCouponNotApplied
    							'// �����ۻ�ǰ�� �ְų� ���ǹ�ۺ� �̻�(���� �ʰ�����)
    							sqlStr = sqlStr + "	0," & vbCrlf
    							sqlStr = sqlStr + "	0," & vbCrlf
								sqlStr = sqlStr + "	" & CLng(0*1/11) & ","
    						else
								if (arrdefaultDeliverPay(i)>0) and (otmpOrder.FItemList(0).ForderDlvPay-arrdefaultDeliverPay(i)>0) then
									'' �űԹ��
									addTenDlvPay = otmpOrder.FItemList(0).ForderDlvPay-arrdefaultDeliverPay(i)  '' 3000�� �޾����� 2500 �������ִ� CASE // �ٹ�� �߰��� ����.

									sqlStr = sqlStr + "	" & arrdefaultDeliverPay(i) & "," & vbCrlf    ''itemcost
									sqlStr = sqlStr + "	" & arrdefaultDeliverPay(i) & "," & vbCrlf    ''reducedPrice
									sqlStr = sqlStr + "	" & arrdefaultDeliverPay(i) & "," & vbCrlf    ''orgitemcost
									sqlStr = sqlStr + "	" & arrdefaultDeliverPay(i) & "," & vbCrlf    ''itemcostCouponNotApplied
									''���԰�
									sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycash
									sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycashCouponNotApplied
									sqlStr = sqlStr + "	" & CLng(arrdefaultDeliverPay(i)*1/11) & ","

								else
									sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcost
									sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''reducedPrice
									sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''orgitemcost
									sqlStr = sqlStr + "	" & otmpOrder.FItemList(0).ForderDlvPay & "," & vbCrlf    ''itemcostCouponNotApplied
									''���԰�
									sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycash
									sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf           ''buycashCouponNotApplied
									sqlStr = sqlStr + "	" & CLng(otmpOrder.FItemList(0).ForderDlvPay*1/11) & ","
								end if
    						end if
                        end if

						sqlStr = sqlStr + "	0,"
						sqlStr = sqlStr + "	'',''"

						sqlStr = sqlStr + " )"
					'rw sqlStr
					'rw isNULL(arrdefaultDeliverPay(i))
						dbget.Execute sqlStr
					else
						sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
						sqlStr = sqlStr + "itemoption, makerid, itemno, itemcost, buycash, buycashCouponNotApplied, itemvat, mileage, itemname, itemoptionname, reducedPrice)" & vbCrlf  '', buycashCouponNotApplied �߰�.
						sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
						sqlStr = sqlStr + " '" & orderserial & "'," & vbCrlf
						sqlStr = sqlStr + " 0," & vbCrlf
						sqlStr = sqlStr + " '" & ("9" & Format00(3, UpcheBeasongOptionIdx)) & "'," & vbCrlf
						sqlStr = sqlStr + " '" + CStr(arrmakerid(i)) + "'," & vbCrlf
						sqlStr = sqlStr + " 1," & vbCrlf
						sqlStr = sqlStr + "	0," & vbCrlf

                        if (IsForeignDLV) then  ''�ؿܹ�� �б�
                            if (arrfreebeasongcount(i) > 0) or (arritemcostcouponnotApplied(i) >= arrdefaultFreeBeasongLimit(i)) then
    							'// �����ۻ�ǰ�� �ְų� ���ǹ�ۺ� �̻�(���� �ʰ�����)
    							sqlStr = sqlStr + "	0," & vbCrlf
								sqlStr = sqlStr + "	0," & vbCrlf
    						else
    							''���԰�
    							sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf
								sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf
    						end if
                        else
    						if (arrfreebeasongcount(i) > 0) or (arrtotitemcost(i) >= arrdefaultFreeBeasongLimit(i)) then
    							'// �����ۻ�ǰ�� �ְų� ���ǹ�ۺ� �̻�(���� �ʰ�����)
    							sqlStr = sqlStr + "	0," & vbCrlf
								sqlStr = sqlStr + "	0," & vbCrlf
    						else
    							''���԰�
    							sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf
								sqlStr = sqlStr + "	" &  arrdefaultDeliverPay(i) & "," & vbCrlf
    						end if
                        end if

						sqlStr = sqlStr + "	0,"
						sqlStr = sqlStr + "	0,"
						sqlStr = sqlStr + "	'','',"
						sqlStr = sqlStr + "	0" & vbCrlf
						sqlStr = sqlStr + " )"
						dbget.Execute sqlStr
					end if
				end if
			next

			Dim tmpTenBeaPay : tmpTenBeaPay=otmpOrder.FItemList(0).ForderDlvPay
			if (IsTenbeaItemExist or IsAllFreeUpcheBeasong or IsForeignDLV or (addTenDlvPay>0)) then  ''IsForeignDLV �߰� 2016/05/24 , addTenDlvPay �߰� 2020/01/28
				'// ��ۺ� �ǸŰ��� �ѹ��� �Է��Ѵ�.
				''IsBeasongSellPriceInserted = True
				''��ü���ǹ���� �ְ� ��ۺ� ���������� ������, ��ۺ��Ǹž��� ��ü���ǹ�ۿ� ����.
				' if (NOT IsForeignDLV) and (IsUpchePaticleBeasongExists) and (tmpTenBeaPay>0) then
				' 	tmpTenBeaPay = 0
				' 	IsBeasongSellPriceInserted = False
				' end if

				if (NOT IsForeignDLV) and (IsUpchePaticleBeasongExists) and (tmpTenBeaPay>0) then
				 	tmpTenBeaPay = 0
				end if

				if (IsBeasongSellPriceInserted) then
					tmpTenBeaPay = 0
				end if

				if (addTenDlvPay>0) then tmpTenBeaPay=addTenDlvPay

				sqlStr = "insert into [db_order].[dbo].tbl_order_detail(masteridx, orderserial,itemid," & vbCrlf
				sqlStr = sqlStr + "itemoption, makerid, itemno, itemcost, reducedprice, orgitemcost, itemcostCouponNotApplied, buycash, buycashCouponNotApplied, itemvat, mileage, itemname, itemoptionname)" & vbCrlf
				sqlStr = sqlStr + " values (" + CStr(iid) + "," & vbCrlf
				sqlStr = sqlStr + " '" & orderserial & "'," & vbCrlf
				sqlStr = sqlStr + " 0," & vbCrlf

				if ucase(countryCode)<>"KR" then
					sqlStr = sqlStr + " '0999'," & vbCrlf
				else
					IF tmpTenBeaPay="0" then
						sqlStr = sqlStr + " '0501'," & vbCrlf
					else
						sqlStr = sqlStr + " '0101'," & vbCrlf
					end if
				end if

				if (tmpTenBeaPay>0) and (NOT IsTenbeaItemExist) then
					sqlStr = sqlStr + " '10x10logistics'," & vbCrlf
				else
					sqlStr = sqlStr + " ''," & vbCrlf
				end if
				sqlStr = sqlStr + " 1," & vbCrlf
				sqlStr = sqlStr + "	"&tmpTenBeaPay&"," & vbCrlf
				sqlStr = sqlStr + "	"&tmpTenBeaPay&"," & vbCrlf
				sqlStr = sqlStr + "	"&tmpTenBeaPay&"," & vbCrlf
				sqlStr = sqlStr + "	"&tmpTenBeaPay&"," & vbCrlf
				sqlStr = sqlStr + "	0," & vbCrlf
				sqlStr = sqlStr + "	0," & vbCrlf                                        ''buycashCouponNotApplied
				sqlStr = sqlStr + "	0,"
				sqlStr = sqlStr + "	0,"
				sqlStr = sqlStr + "	'��ۺ�',''"
				sqlStr = sqlStr + " )"
				''rw sqlStr
				dbget.Execute sqlStr
			end if

		if Err then
			dbget.RollBackTrans
			response.write ErrMsg & Err.Description
			response.end
		else
			ErrMsg = "[004]"
		end if

		sqlStr = "update [db_order].[dbo].tbl_order_master" & vbCrlf
		sqlStr = sqlStr + " set totalvat = " + CStr(buf_totvat) + "" & vbCrlf
		sqlStr = sqlStr + " ,totalsum = " + CStr(buf_totCpnNotAppliedcost) + "" & vbCrlf  '' buf_totcost=>buf_totCpnNotAppliedcost
		sqlStr = sqlStr + " ,subtotalprice = " + CStr(buf_totcost) + "" & vbCrlf
		sqlStr = sqlStr + " ,subtotalPriceCouponNotApplied = " + CStr(buf_totCpnNotAppliedcost) + "" & vbCrlf ''����
		if (buf_totCpnNotAppliedcost>buf_totcost) then
		    sqlStr = sqlStr + " ,tencardspend="&buf_totCpnNotAppliedcost-buf_totcost
		    sqlStr = sqlStr + " ,bCpnIdx=-1"
	    end if
		sqlStr = sqlStr + " where orderserial='" + orderserial + "'"

		'response.write sqlStr & "<BR>"
		dbget.Execute sqlStr

	    ''�ؿܹ�� ���� ����
	    if ucase(countryCode)<>"KR" then
	        dim iUsDollor : iUsDollor = getEmsItemUsDollar(orderserial)

	        sqlStr = "insert into [db_order].[dbo].tbl_ems_orderInfo"
	        sqlStr = sqlStr + "(orderserial"
            sqlStr = sqlStr + ",countryCode"
            sqlStr = sqlStr + ",emsZipCode"
            sqlStr = sqlStr + ",itemGubunName"
            sqlStr = sqlStr + ",goodNames"
            sqlStr = sqlStr + ",itemWeigth"
            sqlStr = sqlStr + ",itemUsDollar"
            sqlStr = sqlStr + ",InsureYn"
            sqlStr = sqlStr + ",InsurePrice"
            sqlStr = sqlStr + ",emsDlvCost"
            sqlStr = sqlStr + ")"
            sqlStr = sqlStr + " values("
            sqlStr = sqlStr + " '" & orderserial + "'" & vbCrlf
            sqlStr = sqlStr + " ,'" & countryCode + "'" & vbCrlf
            sqlStr = sqlStr + " ,'" & otmpOrder.FItemList(0).FReceiveZipCode + "'" & vbCrlf
            sqlStr = sqlStr + " ,'" & getEmsItemGubunName & "'" & vbCrlf
            sqlStr = sqlStr + " ,'" & getEmsGoodNames & "'" & vbCrlf
            sqlStr = sqlStr + " ," & (getEmsTotalWeight(orderserial)-getEmsBoxWeight) & vbCrlf
            sqlStr = sqlStr + " ," & iUsDollor & vbCrlf

            if isEmsInsureRequire(orderserial) then
                sqlStr = sqlStr + " ,'Y'" & vbCrlf
                sqlStr = sqlStr + " ," & getEmsInsurePrice(orderserial) & vbCrlf
            else
                sqlStr = sqlStr + " ,'N'" & vbCrlf
                sqlStr = sqlStr + " ,0" & vbCrlf
            end if
            sqlStr = sqlStr + " ,"& otmpOrder.FItemList(0).ForderDlvPay &"" &vbCrlf
            sqlStr = sqlStr + " )"
            dbget.Execute sqlStr
	    end if

		if Err then
			dbget.RollBackTrans
			response.write ErrMsg & Err.Description
			response.end
		else
			dbget.CommitTrans
			rw "["&orderserial&"]"
		end if

		'' Flag update
		sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"
		sqlStr = sqlStr & " set matchState='O'"
		sqlStr = sqlStr & " ,OrderSerial='"&orderserial&"'"
		sqlStr = sqlStr & " where OutMallorderSerial='"&ixsiteOrderSerial&"'"
		sqlStr = sqlStr & " and matchState='I'"
		''rw sqlStr
		dbget.Execute sqlStr

		''province�� nvarchar�� update select�� ��ȯ // 2022-12-02 ������ ����
		If otmpOrder.FItemList(0).FSellSite = "shopify" and ucase(countryCode)<>"KR" Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R "
			sqlStr = sqlStr & " SET R.city = T.city "
			sqlStr = sqlStr & " , R.province = T.province "
			sqlStr = sqlStr & " , R.provinceCode = T.provinceCode "
			sqlStr = sqlStr & " FROM [db_order].[dbo].tbl_ems_orderInfo R "
			sqlStr = sqlStr & " JOIN db_temp.dbo.tbl_xsite_tmporder T on R.orderserial = T.orderserial "
			sqlStr = sqlStr & " WHERE R.orderserial = '"& orderserial &"' "
			dbget.Execute sqlStr
		End If

		''����� ������Ʈ
		sqlStr = "exec [db_summary].[dbo].sp_ten_RealtimeStock_regOrder '" & orderserial & "'"
		dbget.execute sqlStr

		'' ''����ǰ.//
		'' if (otmpOrder.FItemList(0).FSellSite="wizwid") and (Left(CStr(now()),10)<"2012-09-20") then
		'' 	sqlStr = "exec [db_order].[dbo].sp_Ten_order_gift '" & orderserial & "'"
		'' 	dbget.Execute(sqlStr)
		'' end if

		'' if (otmpOrder.FItemList(0).FSellSite="wconcept") and (Left(CStr(now()),10)<"2012-09-22") then
		'' 	sqlStr = "exec [db_order].[dbo].sp_Ten_order_gift '" & orderserial & "'"
		'' 	dbget.Execute(sqlStr)
		'' end if

		If (otmpOrder.FItemList(0).FSellSite="nvstorefarm") OR (otmpOrder.FItemList(0).FSellSite="kakaostore") Then
			sqlStr = "exec db_order.dbo.[sp_Ten_order_gift_Outmall_storefarm] '" & orderserial & "'"
			dbget.Execute(sqlStr)
		Else
			sqlStr = "exec db_order.dbo.[sp_Ten_order_gift_Outmall] '" & orderserial & "'"
			dbget.Execute(sqlStr)
		End If

    	Set otmpOrder = Nothing
    End if
Next

	''''''

elseif (mode = "MatchItemSeqChg") then
    Dim chgItemID
    outMallorderSeq     = requestCheckvar(request("outMallorderSeq"),20)
    orderItemID         = requestCheckvar(request("orderItemID"),32)
    chgItemID           = requestCheckvar(request("chgItemID"),32)

'    rw outMallorderSeq
'    rw orderItemID
'    rw chgItemID

    sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder"&VbCRLF
    sqlStr = sqlStr & " set matchitemid="&chgItemID&VbCRLF
    sqlStr = sqlStr & " ,orderitemid="&chgItemID&VbCRLF
    sqlStr = sqlStr & " where outmallorderseq="&outMallorderSeq&VbCRLF
    sqlStr = sqlStr & " and matchitemid="&orderItemID

    ''rw sqlStr
    dbget.Execute sqlStr,AssignedRow

    response.write "<script>alert('"&AssignedRow&" �� �ݿ��Ǿ����ϴ�.'); opener.location.reload();window.close();</script>"
    response.end
elseif (mode = "delpInputOrder") then
    Dim validDEL : validDEL = false
    outMallorderSeq     = requestCheckvar(request("outMallorderSeq"),20)
    OutMallOrderSerial  = requestCheckvar(request("OutMallOrderSerial"),32)
    orderItemID         = requestCheckvar(request("orderItemID"),32)
    orderItemOption     = requestCheckvar(request("orderItemOption"),32)

    sqlStr = "select IsNULL(count(*),0) as TTLCNT"
    sqlStr = sqlStr & " , IsNULL(Sum(CASE WHEN ORDERSERIAL is NULL and outmallorderseq="&outMallorderSeq&" THEN 1 ELSE 0 END),0) as NoInputCNT"
    sqlStr = sqlStr & " , IsNULL(sum(CASE WHEN ORDERSERIAL is Not NULL and outmallorderseq<>"&outMallorderSeq&" THEN 1 ELSE 0 END),0) as InputedCNT"
    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder"
    sqlStr = sqlStr & " where outmallorderserial='"&OutMallOrderSerial&"'"
    sqlStr = sqlStr & " and orderitemid="&orderItemID
    sqlStr = sqlStr & " and orderitemoption='"&orderItemOption&"'"
    sqlStr = sqlStr & " and matchstate<>'D'"
''rw sqlStr
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if  not rsget.EOF  then
        validDEL = (rsget("TTLCNT")=2 and rsget("NoInputCNT")=1 and rsget("InputedCNT")=1)
    end if
    rsget.close

    if (validDEL) then
        sqlStr = "update db_temp.dbo.tbl_xSite_TMPOrder"
        sqlStr = sqlStr & " set matchstate='D'"
        sqlStr = sqlStr & " where outMallorderSeq="&outMallorderSeq
        sqlStr = sqlStr & " and outmallorderserial='"&OutMallOrderSerial&"'"
        sqlStr = sqlStr & " and matchstate='I'"
        dbget.execute sqlStr

        response.write "<script>alert('���� �Ǿ����ϴ�.'); opener.location.reload();window.close();</script>"
        dbget.Close() : response.end
    else
        response.write "<script>alert('�� �Էµ� ������ ���� ���� �մϴ�. ���� �� �� �����ϴ�.\n\n������ ���� ���');</script>"
        dbget.Close() : response.end
    end if
elseif (mode="ltimalldel") then
    sqlStr = " update db_temp.dbo.tbl_LTiMall_OrdNoti"
    sqlStr = sqlStr & " set notistatus=9"
    sqlStr = sqlStr & " where notistatus=0"
    sqlStr = sqlStr & " and ORDER_NO='"&requestCheckvar(request("OutMallOrderSerial"),32)&"'"
    sqlStr = sqlStr & " and ORDER_SEQ='"&requestCheckvar(request("outMallorderSeq"),20)&"'"

    dbget.execute sqlStr,AssignedRow

    if (AssignedRow>0) then
        response.write "<script>alert('"&AssignedRow&" �� �Է��� ��ҵǾ����ϴ�..'); window.close();</script>"
        dbget.Close() : response.end
    end if

elseif (mode="ltimallreg") then
    sqlStr = "Insert Into db_temp.dbo.tbl_xSite_TMPOrder"
    sqlStr = sqlStr & " (SellSite,OutMallOrderSerial,SellDate,PayType,PayDate"
    sqlStr = sqlStr & " ,MatchItemID,matchItemoption,orderItemID,OrderItemName,orderItemoption,orderItemOptionName"
    sqlStr = sqlStr & " ,OrderName,OrderTelNo,OrderHpNo"
    sqlStr = sqlStr & " ,ReceiveName,ReceiveTelNo,ReceiveHpNo"
    sqlStr = sqlStr & " ,ReceiveZipCode,ReceiveAddr1,ReceiveAddr2"
    sqlStr = sqlStr & " ,SellPrice,RealSellPrice,vatInclude,ItemOrderCount,deliveryType,DeliveryPrice,deliveryMemo"
    sqlStr = sqlStr & " ,countryCode,matchstate,orderDlvPay,OrgDetailKey,outMallGoodsNo)"
    sqlStr = sqlStr & " select 'lotteimall',ORDER_NO, ORDER_DT,'50',ORDER_DT"
    sqlStr = sqlStr & " ,SubString(ENTP_DT_CODE,0,CHARINDEX('_',ENTP_DT_CODE))"
    sqlStr = sqlStr & " ,SubString(ENTP_DT_CODE,CHARINDEX('_',ENTP_DT_CODE)+1,4)"
    sqlStr = sqlStr & " ,SubString(ENTP_DT_CODE,0,CHARINDEX('_',ENTP_DT_CODE))"
    sqlStr = sqlStr & " ,GOODS_NAME,SubString(ENTP_DT_CODE,CHARINDEX('_',ENTP_DT_CODE)+1,4),GOODSDT_INFO"
    sqlStr = sqlStr & " ,O_NAME,O_TEL,O_HTEL"
    sqlStr = sqlStr & " ,S_NAME,S_TEL,S_HTEL"
    sqlStr = sqlStr & " ,S_POST"
    sqlStr = sqlStr & " ,SubString(S_ADDR,0,CHARINDEX(' ',S_ADDR)+CHARINDEX(' ',SubString(S_ADDR,CHARINDEX(' ',S_ADDR)+1,512)))"
    sqlStr = sqlStr & " ,SubString(S_ADDR,CHARINDEX(' ',S_ADDR)+CHARINDEX(' ',SubString(S_ADDR,CHARINDEX(' ',S_ADDR)+1,512))+1,512)"
    sqlStr = sqlStr & " ,SALE_PRICE,SALE_PRICE,'Y',QTY,0,0,CS_MSG"
    sqlStr = sqlStr & " ,'KR','I',0,ORDER_NO+'-'+ORDER_SEQ"
    sqlStr = sqlStr & " ,Goods_ID"
    sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_OrdNoti"
    sqlStr = sqlStr & " where notistatus=0"
    sqlStr = sqlStr & " order by ORDER_NO,ORDER_SEQ"

    dbget.execute sqlStr,AssignedRow

    if (AssignedRow>0) then
        sqlStr = " Update N"
        sqlStr = sqlStr & " set notistatus=3"
        sqlStr = sqlStr & " from db_temp.dbo.tbl_LTiMall_OrdNoti N"
        sqlStr = sqlStr & " 	Join db_temp.dbo.tbl_xSite_TMPOrder T"
        sqlStr = sqlStr & " 	on N.ORDER_NO=T.OutMallOrderSerial"
        sqlStr = sqlStr & " 	and N.ORDER_NO+'-'+N.ORDER_SEQ=T.OrgDetailKey"
        sqlStr = sqlStr & " where N.notistatus=0"

        dbget.execute sqlStr,AssignedRow

        response.write "<script>alert('"&AssignedRow&" �� �ԷµǾ����ϴ�.'); location.replace('"&refer&"');</script>"
        dbget.Close() : response.end
    else
        response.write "<script>alert('�Է��� �����Ͱ� �����ϴ�.'); location.replace('"&refer&"');</script>"
        dbget.Close() : response.end
    end if

elseif (mode = "dlpre") then
'    sqlStr = "update T
'    sqlStr = sqlStr & " set matchstate='D'"
'    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T"
'    sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_master m"
'    sqlStr = sqlStr & " 	on T.outmallorderserial=m.authcode"
'    sqlStr = sqlStr & " 	and T.sellsite=m.sitename"
'    sqlStr = sqlStr & " 	Join db_order.dbo.tbl_order_detail D"
'    sqlStr = sqlStr & " 	on m.orderserial=D.orderserial"
'    sqlStr = sqlStr & " 	and D.itemid=T.matchitemid"
'    sqlStr = sqlStr & " 	and D.itemoption=T.matchItemOption"
'    sqlStr = sqlStr & " where T.matchstate='I'"
elseif (mode="ssgupdate") then  ''SSG ��ǰ�ߺ� ������Ʈ �и�
    sitegbn = request("sitename")
    outMallorderSeq     = requestCheckvar(request("outMallorderSeq"),20)  '' ssg �����߰�

    sqlStr = ""
	sqlStr = sqlStr & " SELECT T.outmallorderseq, T.outmallorderserial, T.ItemOrderCount, T.Realsellprice, T.orgdetailkey ,T.matchItemid, T.matchItemoption, isNull(T.requireDetail, '') as requireDetail "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xsite_tmpOrder AS T "
	sqlStr = sqlStr & " JOIN ( "
	sqlStr = sqlStr & " 	SELECT TOP 1 P.matchItemid,P.matchItemoption, P.receiveaddr2, count(*) as cnt "
	sqlStr = sqlStr & " 	FROM db_temp.dbo.tbl_xSite_TMPOrder P "
	sqlStr = sqlStr & " 	    JOin (select matchItemid,matchItemoption from db_temp.dbo.tbl_xSite_TMPOrder where outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' and outMallorderSeq='"&outMallorderSeq&"') T1"
	sqlStr = sqlStr & " 	    on P.matchItemid=T1.matchItemid and P.matchItemoption=T1.matchItemoption"
	sqlStr = sqlStr & " 	WHERE P.outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' and P.matchstate='I'  "
	sqlStr = sqlStr & " 	GROUP BY P.matchItemid, P.matchItemoption, P.receiveaddr2 "
	sqlStr = sqlStr & " 	Having count(*) > 1 "
	sqlStr = sqlStr & "	) Dp on T.matchItemid = Dp.matchItemid and T.matchItemoption = Dp.matchItemoption  "
	sqlStr = sqlStr & " WHERE T.sellsite='"&sitegbn&"' "
	sqlStr = sqlStr & " and T.outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
	sqlStr = sqlStr & " ORDER BY T.orgdetailkey ASC "
'	rw "----------[TEST] �� ���� ���� �� �ϴ� UPDATE & DELETE ���� �� �� ---------------------------"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

	If rsget.RecordCount > 0 Then
	    If not rsget.EOF Then
			requireDetailAdd = ""
		    Do until rsget.Eof
				sumItemOrderCount	= sumItemOrderCount	+ rsget("ItemOrderCount")
				sumRealsellprice	= sumRealsellprice	+ (rsget("Realsellprice") * rsget("ItemOrderCount"))
				orgdetailkey 		= orgdetailkey & rsget("orgdetailkey") & ","
				If rsget("requireDetail") <> "" Then
					If rsget("ItemOrderCount") > 1 Then
						requireDetailAdd	= requireDetailAdd & rsget("requireDetail") & "/" & rsget("ItemOrderCount") & "��!{!{"
					Else
						requireDetailAdd	= requireDetailAdd & rsget("requireDetail") & "!{!{"
					End If
				End If
				matchItemid			= rsget("matchItemid")
				matchItemoption		= rsget("matchItemoption")
				orgdetailkeylength	= Len(rsget("orgdetailkey"))
			    rsget.moveNext
			Loop
	    End If
	    rsget.Close
	    avgRealsellprice = Clng(sumRealsellprice/sumItemOrderCount)
		If Right(orgdetailkey,1)="," then orgdetailkey=Left(orgdetailkey,Len(orgdetailkey)-1)
		If Right(requireDetailAdd,4)="!{!{" then requireDetailAdd=Left(requireDetailAdd,Len(requireDetailAdd)-4)
		requireDetailAdd = Replace(requireDetailAdd, "!{!{", CHR(3)&CHR(4))
		orgdetailkeyGRoup = Split(orgdetailkey, ",")
		orgdetailkeyMin = orgdetailkeyGRoup(0)
		For lp = 0 to Ubound(orgdetailkeyGRoup)
			If orgdetailkeyGRoup(0) <> orgdetailkeyGRoup(lp) Then
				orgdetailkeyNotMin = orgdetailkeyNotMin & orgdetailkeyGRoup(lp)&","
			End If
		Next
		If Right(orgdetailkeyNotMin,1)="," then orgdetailkeyNotMin=Left(orgdetailkeyNotMin,Len(orgdetailkeyNotMin)-1)

		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xsite_tmpOrder "
		sqlStr = sqlStr & " SET RealSellprice="&avgRealsellprice&" "
		sqlStr = sqlStr & " ,itemorderCount="&sumItemOrderCount&" "
		sqlStr = sqlStr & " ,requireDetail='"&html2db(requireDetailAdd)&"' "
		sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
		sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
		sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
		sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
		sqlStr = sqlStr & " and outMallorderSeq='"&outMallorderSeq&"' "

	'	rw sqlStr
	'	rw "-----------[���� ����� ���� �ƴմϴ�.--------------------------"
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xsite_mayDelOrder "
		sqlStr = sqlStr & " (OrderSerial,SellSite,SellSiteName,OutMallOrderSerial,SellDate,PayType,PayDate,matchItemID,matchitemoption,orderItemID,orderItemName,orderItemOption,orderItemOptionName,prdcode,locationidmaker,sellsiteUserID,OrderName,OrderEmail,OrderTelNo,OrderHpNo,ReceiveName,ReceiveTelNo,ReceiveHpNo,ReceiveZipCode,ReceiveAddr1,ReceiveAddr2,SellPrice,RealSellPrice,vatinclude,ItemOrderCount,DeliveryType,deliveryprice,RegDate,deliverymemo,countryCode,requireDetail,matchState,orderDlvPay,OrgDetailKey,sendState,sendReqCNT,outMallGoodsNo,orderCsGbn,ref_OutMallOrderSerial,ref_CSID,etcFinUser,changeitemid,changeitemoption,orgOrderCNT,recvSendState,recvSendReqCnt,shoplinkerOrderID,tenCpnUint,mallCpnUnit,PRE_USE_UNITCOST,outMallJMonth) "
		sqlStr = sqlStr & " SELECT OrderSerial,SellSite,SellSiteName,OutMallOrderSerial,SellDate,PayType,PayDate,matchItemID,matchitemoption,orderItemID,orderItemName,orderItemOption,orderItemOptionName,prdcode,locationidmaker,sellsiteUserID,OrderName,OrderEmail,OrderTelNo,OrderHpNo,ReceiveName,ReceiveTelNo,ReceiveHpNo,ReceiveZipCode,ReceiveAddr1,ReceiveAddr2,SellPrice,RealSellPrice,vatinclude,ItemOrderCount,DeliveryType,deliveryprice,RegDate,deliverymemo,countryCode,requireDetail,matchState,orderDlvPay,OrgDetailKey,sendState,sendReqCNT,outMallGoodsNo,orderCsGbn,ref_OutMallOrderSerial,ref_CSID,etcFinUser,changeitemid,changeitemoption,orgOrderCNT,recvSendState,recvSendReqCnt,shoplinkerOrderID,tenCpnUint,mallCpnUnit,PRE_USE_UNITCOST,outMallJMonth "
		sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xsite_tmpOrder "
		sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
		sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
		sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
		sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
		sqlStr = sqlStr & " and outMallorderSeq<>'"&outMallorderSeq&"' "
'		rw sqlStr
'		rw "-----------[���� ����� ���� �ƴմϴ�.--------------------------"
		dbget.execute sqlStr

        sqlStr = ""
        sqlStr = sqlStr & " Update  db_temp.dbo.tbl_xsite_tmpOrder "
        sqlStr = sqlStr & " set ref_outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
        sqlStr = sqlStr & " ,matchstate='O'"
        sqlStr = sqlStr & " where outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
		sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
		sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
        sqlStr = sqlStr & " and outMallorderSeq<>'"&outMallorderSeq&"' "
        dbget.execute sqlStr
		response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload(); window.close();</script>"
		dbget.Close() : response.end

	Else
		response.write "<script>alert('���̻� ��ĥ ��ǰ�� �����ϴ�.');opener.location.reload(); window.close();</script>"
		dbget.Close() : response.end
	End If

elseif (mode="gsshopupdate") then
	sitegbn = request("sitename")
	If sitegbn = "gsshop" Then
		sitegbn = "gseshop"
	ElseIf sitegbn = "auction1010" Then
		sitegbn = "auction1010"
	ElseIf sitegbn = "gmarket1010" Then
		sitegbn = "gmarket1010"
	ElseIf sitegbn = "hyundaihmall" or sitegbn = "Hmall" or sitegbn = "hmall1010" Then
		sitegbn = "hmall1010"
	ElseIf sitegbn = "interpark" Then
		sitegbn = "interpark"
	ElseIf sitegbn = "lotteimall" Then
		sitegbn = "lotteimall"
	ElseIf sitegbn = "ezwel" Then
		sitegbn = "ezwel"
	ElseIf sitegbn = "lotteon" Then
		sitegbn = "lotteon"
	ElseIf sitegbn = "shintvshopping" Then
		sitegbn = "shintvshopping"
	ElseIf sitegbn = "skstoa" Then
		sitegbn = "skstoa"
	ElseIf sitegbn = "LFmall" Then
		sitegbn = "lfmall"
	ElseIf sitegbn = "kakaostore" Then
		sitegbn = "kakaostore"
	ElseIf sitegbn = "boribori1010" Then
		sitegbn = "boribori1010"
	ElseIf sitegbn = "wconcept1010" Then
		sitegbn = "wconcept1010"
	ElseIf sitegbn = "withnature1010" Then
		sitegbn = "withnature1010"
	End If

    if (LCASE(sitegbn)="ssg") then
        response.write "����� �� �����ϴ�. SSG"
        response.end
        dbget.close()
    end if

    outMallorderSeq     = requestCheckvar(request("outMallorderSeq"),20)  '' ssg �����߰�
	Dim tmpSeq, tmpSeqGroup, tmpSeqMin, tmpSeqNotMin
	sqlStr = ""
	sqlStr = sqlStr & " SELECT T.outmallorderseq, T.outmallorderserial, T.ItemOrderCount, T.Realsellprice, T.orgdetailkey ,T.matchItemid, T.matchItemoption, isNull(T.requireDetail, '') as requireDetail "
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xsite_tmpOrder AS T "
	sqlStr = sqlStr & " JOIN ( "
	sqlStr = sqlStr & " 	SELECT TOP 1 P.matchItemid,P.matchItemoption, P.receiveaddr2, count(*) as cnt "
	sqlStr = sqlStr & " 	FROM db_temp.dbo.tbl_xSite_TMPOrder P "
	sqlStr = sqlStr & " 	    JOin (select matchItemid,matchItemoption from db_temp.dbo.tbl_xSite_TMPOrder where outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' and outMallorderSeq='"&outMallorderSeq&"') T1"
	sqlStr = sqlStr & " 	    on P.matchItemid=T1.matchItemid and P.matchItemoption=T1.matchItemoption"
	sqlStr = sqlStr & " 	WHERE P.outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' and P.matchstate='I'  "
	sqlStr = sqlStr & " 	GROUP BY P.matchItemid, P.matchItemoption, P.receiveaddr2 "
	sqlStr = sqlStr & " 	Having count(*) > 1 "
	sqlStr = sqlStr & "	) Dp on T.matchItemid = Dp.matchItemid and T.matchItemoption = Dp.matchItemoption  "
	sqlStr = sqlStr & " WHERE T.sellsite='"&sitegbn&"' "
	sqlStr = sqlStr & " and T.outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
	sqlStr = sqlStr & " ORDER BY T.orgdetailkey ASC "
'	rw "----------[TEST] �� ���� ���� �� �ϴ� UPDATE & DELETE ���� �� �� ---------------------------"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget.RecordCount > 0 Then
	    If not rsget.EOF Then
			requireDetailAdd = ""
		    Do until rsget.Eof
				sumItemOrderCount	= sumItemOrderCount	+ rsget("ItemOrderCount")
				sumRealsellprice	= sumRealsellprice	+ (rsget("Realsellprice") * rsget("ItemOrderCount"))
				orgdetailkey 		= orgdetailkey & rsget("orgdetailkey") & ","
				tmpSeq				= tmpSeq & rsget("outmallorderseq") & ","
				If rsget("requireDetail") <> "" Then
					If rsget("ItemOrderCount") > 1 Then
						requireDetailAdd	= requireDetailAdd & rsget("requireDetail") & "/" & rsget("ItemOrderCount") & "��!{!{"
					Else
						requireDetailAdd	= requireDetailAdd & rsget("requireDetail") & "!{!{"
					End If
				End If
				matchItemid			= rsget("matchItemid")
				matchItemoption		= rsget("matchItemoption")
				orgdetailkeylength	= Len(rsget("orgdetailkey"))
			    rsget.moveNext
			Loop
	    End If
	    rsget.Close
	    avgRealsellprice = Clng(sumRealsellprice/sumItemOrderCount)
		If Right(orgdetailkey,1)="," then orgdetailkey=Left(orgdetailkey,Len(orgdetailkey)-1)
		If Right(tmpSeq,1)="," then tmpSeq=Left(tmpSeq,Len(tmpSeq)-1)
		If Right(requireDetailAdd,4)="!{!{" then requireDetailAdd=Left(requireDetailAdd,Len(requireDetailAdd)-4)
		requireDetailAdd = Replace(requireDetailAdd, "!{!{", CHR(3)&CHR(4))
		orgdetailkeyGRoup = Split(orgdetailkey, ",")
		orgdetailkeyMin = orgdetailkeyGRoup(0)

		tmpSeqGroup = Split(tmpSeq, ",")
		tmpSeqMin = tmpSeqGroup(0)
		For lp = 0 to Ubound(orgdetailkeyGRoup)
			If orgdetailkeyGRoup(0) <> orgdetailkeyGRoup(lp) Then
				orgdetailkeyNotMin = orgdetailkeyNotMin & orgdetailkeyGRoup(lp)&","
			End If
		Next

		For lp = 0 to Ubound(tmpSeqGroup)
			If tmpSeqGroup(0) <> tmpSeqGroup(lp) Then
				tmpSeqNotMin = tmpSeqNotMin & tmpSeqGroup(lp)&","
			End If
		Next

		If Right(orgdetailkeyNotMin,1)="," then orgdetailkeyNotMin=Left(orgdetailkeyNotMin,Len(orgdetailkeyNotMin)-1)
		If Right(tmpSeqNotMin,1)="," then tmpSeqNotMin=Left(tmpSeqNotMin,Len(tmpSeqNotMin)-1)

		If orgdetailkeyNotMin <> "" Then
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xsite_tmpOrder "
			sqlStr = sqlStr & " SET RealSellprice="&avgRealsellprice&" "
			sqlStr = sqlStr & " ,itemorderCount="&sumItemOrderCount&" "
			sqlStr = sqlStr & " ,requireDetail='"&requireDetailAdd&"' "
			sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
			sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
			sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
			sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
			if (LCASE(sitegbn)="ssg") then  ''//2017/12/11 �����߰�
				sqlStr = sqlStr & " and outMallorderSeq='"&outMallorderSeq&"' "
			end if
'			sqlStr = sqlStr & " and orgdetailKey = "&orgdetailkeyMin
			sqlStr = sqlStr & " and outmallorderseq = "&tmpSeqMin
			dbget.execute sqlStr

			If sitegbn = "auction1010" OR sitegbn = "boribori1010" OR sitegbn = "shintvshopping" OR sitegbn = "skstoa" OR sitegbn = "lfmall" OR sitegbn = "kakaostore" OR sitegbn = "wconcept1010" OR sitegbn = "withnature1010" OR sitegbn = "interpark" OR sitegbn = "gmarket1010" OR sitegbn = "gseshop" OR sitegbn = "nvstorefarm" OR sitegbn = "nvstoremoonbangu" OR sitegbn = "Mylittlewhoopee" OR sitegbn = "nvstoregift" OR sitegbn = "lotteimall" OR sitegbn = "lotteCom" OR sitegbn = "ezwel" OR sitegbn = "hmall1010" OR sitegbn = "WMP" OR sitegbn = "wmpfashion" OR LCASE(sitegbn)="ssg" OR LCASE(sitegbn)="lotteon" Then
				sqlStr = ""
				sqlStr = sqlStr & " INSERT INTO db_temp.dbo.tbl_xsite_mayDelOrder "
				sqlStr = sqlStr & " (OrderSerial,SellSite,SellSiteName,OutMallOrderSerial,SellDate,PayType,PayDate,matchItemID,matchitemoption,orderItemID,orderItemName,orderItemOption,orderItemOptionName,prdcode,locationidmaker,sellsiteUserID,OrderName,OrderEmail,OrderTelNo,OrderHpNo,ReceiveName,ReceiveTelNo,ReceiveHpNo,ReceiveZipCode,ReceiveAddr1,ReceiveAddr2,SellPrice,RealSellPrice,vatinclude,ItemOrderCount,DeliveryType,deliveryprice,RegDate,deliverymemo,countryCode,requireDetail,matchState,orderDlvPay,OrgDetailKey,sendState,sendReqCNT,outMallGoodsNo,orderCsGbn,ref_OutMallOrderSerial,ref_CSID,etcFinUser,changeitemid,changeitemoption,orgOrderCNT,recvSendState,recvSendReqCnt,shoplinkerOrderID,tenCpnUint,mallCpnUnit,PRE_USE_UNITCOST,outMallJMonth) "
				sqlStr = sqlStr & " SELECT OrderSerial,SellSite,SellSiteName,OutMallOrderSerial,SellDate,PayType,PayDate,matchItemID,matchitemoption,orderItemID,orderItemName,orderItemOption,orderItemOptionName,prdcode,locationidmaker,sellsiteUserID,OrderName,OrderEmail,OrderTelNo,OrderHpNo,ReceiveName,ReceiveTelNo,ReceiveHpNo,ReceiveZipCode,ReceiveAddr1,ReceiveAddr2,SellPrice,RealSellPrice,vatinclude,ItemOrderCount,DeliveryType,deliveryprice,RegDate,deliverymemo,countryCode,requireDetail,matchState,orderDlvPay,OrgDetailKey,sendState,sendReqCNT,outMallGoodsNo,orderCsGbn,ref_OutMallOrderSerial,ref_CSID,etcFinUser,changeitemid,changeitemoption,orgOrderCNT,recvSendState,recvSendReqCnt,shoplinkerOrderID,tenCpnUint,mallCpnUnit,PRE_USE_UNITCOST,outMallJMonth "
				sqlStr = sqlStr & " FROM db_temp.dbo.tbl_xsite_tmpOrder "
				sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
				sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
				sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
				sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
'				sqlStr = sqlStr & " and orgdetailKey in ("&orgdetailkeyNotMin&")"
				sqlStr = sqlStr & " and outMallorderSeq in ("&tmpSeqNotMin&")"
		'		rw sqlStr
		'		rw "-----------[���� ����� ���� �ƴմϴ�.--------------------------"
				dbget.execute sqlStr
			End If

			if (LCASE(sitegbn)="ssg") then
				sqlStr = ""
				sqlStr = sqlStr & " Update  db_temp.dbo.tbl_xsite_tmpOrder "
				sqlStr = sqlStr & " set ref_outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
				sqlStr = sqlStr & " ,matchstate='O'"
				sqlStr = sqlStr & " where outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
				sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
				sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
				sqlStr = sqlStr & " and outMallorderSeq<>'"&outMallorderSeq&"' "
				dbget.execute sqlStr
				response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload(); window.close();</script>"
				dbget.Close() : response.end
			else
				' sqlStr = ""
				' sqlStr = sqlStr & " DELETE FROM db_temp.dbo.tbl_xsite_tmpOrder "
				' sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
				' sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
				' sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
				' sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
				' sqlStr = sqlStr & " and orgdetailKey in ("&orgdetailkeyNotMin&")"
				' dbget.execute sqlStr

				sqlStr = ""
				sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xsite_tmpOrder "
				sqlStr = sqlStr & " SET itemorderCount = 0 "
				sqlStr = sqlStr & " WHERE sellsite='"&sitegbn&"' "
				sqlStr = sqlStr & " and outmallorderserial='"&requestCheckvar(request("OutMallOrderSerial"),32)&"' "
				sqlStr = sqlStr & " and matchItemid='"&matchItemid&"' "
				sqlStr = sqlStr & " and matchItemoption='"&matchItemoption&"' "
'				sqlStr = sqlStr & " and orgdetailKey in ("&orgdetailkeyNotMin&")"
				sqlStr = sqlStr & " and outMallorderSeq in ("&tmpSeqNotMin&")"
				dbget.execute sqlStr

				'response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload(); window.close();</script>"
				response.write "�����Ǿ����ϴ�.<script>setTimeout('opener.location.reload();window.close();', 500);</script>"
				dbget.Close() : response.end
			end if
		Else
			'response.write "<script>alert('�����߻�.orgdetailkeyNotMin�� �����ϴ�.');opener.location.reload(); window.close();</script>"
			response.write "�����߻�.orgdetailkeyNotMin�� �����ϴ�.<script>setTimeout('opener.location.reload();window.close();', 2000);</script>"
			dbget.Close() : response.end
		End If
	Else
		'response.write "<script>alert('����� Ȯ�����ּ���.\n���� : ������ �븮');opener.location.reload(); window.close();</script>"
		response.write "����� Ȯ�����ּ���.<br>���� : ������ �븮<script>setTimeout('opener.location.reload();window.close();', 2000);</script>"
		dbget.Close() : response.end
	End If
ElseIf mode = "realDel" Then
	Dim delordCnt
	dummyseqarr = cksel
	dummyseqarr = Replace(dummyseqarr, ", ", ",")
	dummyseqarr = Replace(dummyseqarr, ",", "','")
	dummyseqarr = "'"&dummyseqarr&"'"

	sqlStr = ""
	sqlStr = sqlStr & " SELECT COUNT(*) as CNT FROM db_temp.dbo.tbl_xsite_tmpOrder "
	sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
	sqlStr = sqlStr & " and Orderserial is NULL "
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    	delordCnt = rsget("CNT")
    rsget.Close

	If delordCnt < 100 Then
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_temp.dbo.tbl_xsite_tmpOrder "
		sqlStr = sqlStr & " WHERE OutMallOrderSerial in (" & dummyseqarr & ") "
		sqlStr = sqlStr & " and Orderserial is NULL "
		dbget.execute sqlStr
		response.write "<script>alert('�����Ͽ����ϴ�.');</script>"
		dbget.Close() : response.end
	Else
		response.write "<script>alert('�ѹ��� ������ ���� 100�� �̻��Դϴ�.\n\n������ ���� ���');</script>"
	End If
ElseIf mode = "realPriceUpd" Then
	splitedSeq = split(cksel,",")
	For j = LBound(splitedSeq) to UBound(splitedSeq)
		sqlStr = ""
		sqlStr = sqlStr & " IF EXISTS(SELECT * FROM db_temp.dbo.tbl_xsite_tmporder WHERE sellsite in ('interpark', 'gseshop', 'alphamall', 'ohou1010', 'wadsmartstore', 'aboutpet', 'shintvshopping', 'goodwearmall10') and orderserial is null and OutMallOrderSerial = '"& Trim(splitedSeq(j)) &"' and RealSellPrice < 1) "& VbCRLF
		sqlStr = sqlStr & " 	UPDATE db_temp.dbo.tbl_xsite_tmporder "& VbCRLF
		sqlStr = sqlStr & " 	SET RealSellPrice = 1 "& VbCRLF
		sqlStr = sqlStr & " 	WHERE sellsite in ('interpark', 'gseshop', 'alphamall', 'ohou1010', 'wadsmartstore', 'aboutpet', 'shintvshopping', 'goodwearmall10')  "& VbCRLF
		sqlStr = sqlStr & " 	and orderserial is null "& VbCRLF
		sqlStr = sqlStr & " 	and OutMallOrderSerial = '"& Trim(splitedSeq(j)) &"' "& VbCRLF
		sqlStr = sqlStr & " 	and RealSellPrice < 1 "
		dbget.execute sqlStr
	Next
	response.write "<script>alert('���� �Ͽ����ϴ�.');</script>"
	dbget.Close() : response.end
ElseIf mode = "rateCal" Then
	sitegbn = request("sitename")
    If (LCASE(sitegbn) <> "shopify") Then
        response.write "����� �� �����ϴ�."
        response.end
        dbget.close()
    End If

	outMallorderSeq     = requestCheckvar(request("outMallorderSeq"),20)
	Dim USDrate
	sqlStr = ""
	sqlStr = sqlStr & " SELECT isNull(USD, '') as USD "
	sqlStr = sqlStr & " FROM db_item.dbo.tbl_dayexchageRate "
	sqlStr = sqlStr & " WHERE yyyymmdd = '"& request("paydate") &"' "
    rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If NOT rsget.Eof Then
		USDrate = rsget("USD")
	End If
	rsget.Close

	If USDrate = "" Then
		response.write "<script>alert('ȯ�� ������ �Էµ��� �ʾҽ��ϴ�');</script>"
		response.write "<script>window.close();</script>"
		dbget.close()	:	response.End
	Else
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_temp.dbo.tbl_xsite_tmporder "
		sqlStr = sqlStr & " SET SellPrice = SellPrice * "& USDrate &" "
		sqlStr = sqlStr & " ,RealSellPrice = RealSellPrice * "& USDrate &" "
		sqlStr = sqlStr & " ,orderDlvPay = orderDlvPay * "& USDrate &" "
		sqlStr = sqlStr & " WHERE OutMallOrderSeq = '"& outMallorderSeq &"' "
		sqlStr = sqlStr & " and sellsite='shopify' "
		dbget.execute sqlStr
		response.write "<script>opener.location.reload();window.close();</script>"
		dbget.close()	:	response.End
	End If
End If

''on error Goto 0

''response.end
%>

<% if (isbatchMode) then %>
<% response.write "<script>parent.addResultLog("&oseq&",'"&orderserial&"');parent.fnNextOrderInputProc();</script>" %>
<% else %>
<script>alert('����Ǿ����ϴ�.');</script>
<script>//location.replace('<%= refer %>');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
