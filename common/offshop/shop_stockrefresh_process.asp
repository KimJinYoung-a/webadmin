<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ���
' History : 2009.04.07 �̻� ����
'			2017.04.11 �ѿ�� ����(���Ȱ���ó��)
'###########################################################
%>
<!-- #include virtual="/common/incSessionBctID.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim refer, alertStr
refer = request.ServerVariables("HTTP_REFERER")

dim mode,itemgubun,itemid,itemoption, makerid, realstock, errrealcheckno, shopid
dim refreshstartdate
dim yyyymmdd, stockdate
dim cksel, Arritemgubun, Arritemid, Arritemoption, Arrrealstock
dim samplestock, stTakingIdx, stStatus, SType
dim yyyymm
mode	    = requestCheckvar(request.form("mode"),32)
itemgubun   = requestCheckvar(request.form("itemgubun"),2)
itemid      = requestCheckvar(request.form("itemid"),9)
itemoption  = requestCheckvar(request.form("itemoption"),4)
makerid     = requestCheckvar(request.form("makerid"),32)
realstock   = requestCheckvar(request.form("realstock"),9)
shopid      = requestCheckvar(request.form("shopid"),32)
yyyymmdd    = requestCheckvar(request.form("yyyymmdd"),10)
stockdate   = requestCheckvar(request.form("stockdate"),10)
samplestock = requestCheckvar(request.form("samplestock"),10)
SType       = requestCheckvar(request.form("SType"),32)
cksel           = request.form("cksel")
Arritemgubun    = request.form("Arritemgubun")
Arritemid       = request.form("Arritemid")
Arritemoption   = request.form("Arritemoption")
Arrrealstock    = request.form("Arrrealstock")
stTakingIdx     = requestCheckvar(request.form("stTakingIdx"),10)
stStatus        = requestCheckvar(request.form("stStatus"),10)
yyyymm      = requestCheckvar(request.form("yyyymm"),7)

dim BasicMonth, ThisDate
BasicMonth  = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)
ThisDate    = Left(CStr(now()),10)

dim sqlStr, AssignedRows, i, chkVal
AssignedRows =0

'''rw mode

if (mode="OFFStockitemRecentRefresh") then
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_daily_shopstock_maker] "&CHKIIF(shopid="","NULL","'"&shopid&"'")&",'" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

elseif (mode="itemAccStockShop") then
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_daily_shopstock_maker] "&CHKIIF(shopid="","NULL","'"&shopid&"'")&",'" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

    ''������̵� ���ʿ�..
    sqlStr = "exec db_summary.[dbo].[usp_STOCK_ITEM_monthly_acc_shopstock_maker] '"&yyyymm&"','" & itemgubun & "'," & itemid & "," & CHKIIF(itemoption="" or itemoption="0000","NULL","'"&itemoption&"'") & ""
    dbget.Execute sqlStr

elseif (mode="OFFitemAllRefresh") then
    ''-1 ���� ������Ʈ
    sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    'dbget.Execute sqlStr

''rw sqlStr
    ''-1 �Ϻ� ������Ʈ
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

	'// �����
    sqlStr = "exec [db_summary].[dbo].usp_Ten_ShopChulgo_Update_One '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

	'// ��ǰ��
    sqlStr = "exec [db_summary].[dbo].usp_Ten_ShopReturn_Update_One '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

''rw sqlStr

    ''response.end
elseif (mode="Offerrcheckupdate") then
    ''���� �ǻ� ��� ����.
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekErr_Input_By_CurrentStock '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & realstock & ",'" & stockdate & "','" & session("ssBctID") & "'"
    dbget.Execute sqlStr
elseif (mode="OffSampleCheckupdate") then
    ''���� ���� ��� ����.
    samplestock = samplestock *-1
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekSample_Input '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'," & samplestock & ",'" & session("ssBctID") & "'"
    dbget.Execute sqlStr
elseif (mode="ArrOfferrcheckupdate") then
'    rw "cksel::"&cksel
'    rw "Arritemgubun::"&Arritemgubun
'    rw "Arritemid::"&Arritemid
'    rw "Arritemoption::"&Arritemoption
'    rw "Arrrealstock::"&Arrrealstock

    cksel           = split(cksel,",")
    Arritemgubun    = split(Arritemgubun,",")
    Arritemid       = split(Arritemid,",")
    Arritemoption   = split(Arritemoption,",")
    Arrrealstock    = split(Arrrealstock,",")

    for i=LBound(cksel) to UBound(cksel)
        chkVal = Trim(cksel(i))
        if (chkVal<>"") then
            if (Trim(Arritemgubun(chkVal))<>"") and (Trim(Arritemid(chkVal))<>"") and (Trim(Arritemoption(chkVal))<>"") and (Trim(Arrrealstock(chkVal))<>"") then
                sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_realchekErr_Input_By_CurrentStock '" & shopid & "','" & requestCheckVar(Trim(Arritemgubun(chkVal)),2) & "'," & requestCheckVar(Trim(Arritemid(chkVal)),10) & ",'" & requestCheckVar(Trim(Arritemoption(chkVal)),4) & "'," & requestCheckVar(Trim(Arrrealstock(chkVal)),10) & ",'" & stockdate & "','" & session("ssBctID") & "'"

                'response.write sqlStr & "<br>"
                dbget.Execute sqlStr
            end if
        end if
    next

    ''����ľ� �Ͻ� ����. 2011-08 eastone �߰�
    sqlStr = "update db_shop.dbo.tbl_shop_designer"
    sqlStr = sqlStr + " set lastStockdate=getdate()"
    sqlStr = sqlStr + " where shopid='"&shopid&"'"
    sqlStr = sqlStr + " and makerid='"&makerid&"'"

    dbget.Execute sqlStr

    if (SType="stTaking") and (stTakingIdx<>"") then
        sqlStr = "update db_shop.dbo.tbl_shop_stockTaking_master"&VbCRLF
        sqlStr = sqlStr + " set finishuserid='"&session("ssBctID")&"'"&VbCRLF
        sqlStr = sqlStr + " ,ststatus=7"&VbCRLF
        sqlStr = sqlStr + " ,inputFinishdate=getdate()"&VbCRLF
        sqlStr = sqlStr + " where stTakingIdx="&stTakingIdx&VbCRLF

        dbget.Execute sqlStr
    end if
    alertStr = "��� �Է� �Ϸ� �Ǿ����ϴ�."
    response.write "<script type='text/javascript'>alert('"&alertStr&"');opener.location.reload();window.close();</script>"
elseif (mode="ArrOffStockTakingupdate") then
'    rw "cksel::"&cksel
'    rw "Arritemgubun::"&Arritemgubun
'    rw "Arritemid::"&Arritemid
'    rw "Arritemoption::"&Arritemoption
'    rw "Arrrealstock::"&Arrrealstock

    cksel           = split(cksel,",")
    Arritemgubun    = split(Arritemgubun,",")
    Arritemid       = split(Arritemid,",")
    Arritemoption   = split(Arritemoption,",")
    Arrrealstock    = split(Arrrealstock,",")

    for i=LBound(cksel) to UBound(cksel)
        chkVal = Trim(cksel(i))
        if (chkVal<>"") then
            if (stTakingIdx<>"") and (Trim(Arritemgubun(chkVal))<>"") and (Trim(Arritemid(chkVal))<>"") and (Trim(Arritemoption(chkVal))<>"") and (Trim(Arrrealstock(chkVal))<>"") then
                sqlStr = "update db_shop.dbo.tbl_shop_stockTaking_Detail"&VbCRLF
                sqlStr = sqlStr & " set stNo="& requestCheckVar(Trim(Arrrealstock(chkVal)),10) &VbCRLF
                sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx&VbCRLF
                sqlStr = sqlStr & " and itemgubun='"& requestCheckVar(Trim(Arritemgubun(chkVal)),2) &"'"
                sqlStr = sqlStr & " and itemid="& requestCheckVar(Trim(Arritemid(chkVal)),10) &""
                sqlStr = sqlStr & " and itemoption='"& requestCheckVar(Trim(Arritemoption(chkVal)),4) &"'"

                dbget.Execute sqlStr, AssignedRows

                if (AssignedRows=0) then
                    sqlStr = "Insert Into db_shop.dbo.tbl_shop_stockTaking_Detail"&VbCRLF
                    sqlStr = sqlStr & " (stTakingIdx,itemgubun,itemid,itemoption,stNo)"&VbCRLF
                    sqlStr = sqlStr & " values("&stTakingIdx&VbCRLF
                    sqlStr = sqlStr & " ,'"& requestCheckVar(Trim(Arritemgubun(chkVal)),2) &"'"&VbCRLF
                    sqlStr = sqlStr & " ,"& requestCheckVar(Trim(Arritemid(chkVal)),10) &VbCRLF
                    sqlStr = sqlStr & " ,'"& requestCheckVar(Trim(Arritemoption(chkVal)),4) &"'"&VbCRLF
                    sqlStr = sqlStr & " ,"& requestCheckVar(Trim(Arrrealstock(chkVal)),10) &VbCRLF
                    sqlStr = sqlStr & " )"

                    dbget.Execute sqlStr
                end if

                ''������ ����..?
                sqlStr = "delete from  db_shop.dbo.tbl_shop_stockTaking_Detail"&VbCRLF
                sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx&VbCRLF
                sqlStr = sqlStr & " and stNo=0"
                dbget.Execute sqlStr

            end if
        end if
    next

elseif (mode="stockTakingNext") then
    if (stStatus="3") and (stockdate<>"") Then
        sqlStr = "update db_shop.dbo.tbl_shop_stockTaking_Master"&VbCRLF
        sqlStr = sqlStr & " set stStatus="&stStatus&VbCRLF
        sqlStr = sqlStr & " ,stockdate='"&stockdate&"'"&VbCRLF
        sqlStr = sqlStr & " ,inputReqdate=getdate()"&VbCRLF
        sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx&VbCRLF
        sqlStr = sqlStr & " and stStatus=0"

        dbget.Execute sqlStr

        alertStr = "����ľ� �Է��� ��û �Ǿ����ϴ�."

    elseif  (stStatus="0") Then
        sqlStr = "update db_shop.dbo.tbl_shop_stockTaking_Master"&VbCRLF
        sqlStr = sqlStr & " set stStatus="&stStatus&VbCRLF
        sqlStr = sqlStr & " ,stockdate=NULL"&VbCRLF
        sqlStr = sqlStr & " ,inputReqdate=NULL"&VbCRLF
        sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx&VbCRLF
        sqlStr = sqlStr & " and stStatus=3"

        dbget.Execute sqlStr, AssignedRows

        if (AssignedRows>0) then
           alertStr = "���� �Ǿ����ϴ�."
        else
           alertStr = "������ ó�� �� ������ �߻� �Ͽ����ϴ�."
        end if
    else
        alertStr = "����� ������ �����ϴ�."
    end if

elseif (mode="OffErrDelete") then
    sqlStr = "delete from [db_summary].[dbo].tbl_erritem_shop_summary" + VbCrlf
    sqlStr = sqlStr + " where yyyymmdd='" + yyyymmdd + "'" + VbCrlf
    sqlStr = sqlStr + " and itemgubun='" + itemgubun + "'" + VbCrlf
    sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + "" + VbCrlf
    sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
    sqlStr = sqlStr + " and shopid='" + shopid + "'"

    dbget.Execute sqlStr

    if (CDate(BasicMonth+"-01")>CDate(yyyymmdd)) then
        ''-1 ���� ������Ʈ
        sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_UpdateALL] '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
        dbget.Execute sqlStr

        response.write "."
    end if
    ''-1 �Ϻ� ������Ʈ
    sqlStr = "exec [db_summary].[dbo].sp_Ten_Shop_Stock_RecentUpdateByItem '" & shopid & "','" & itemgubun & "'," & itemid & ",'" & itemoption & "'"
    dbget.Execute sqlStr

elseif (mode="stockTakingDel") then
    sqlStr = " delete from db_shop.dbo.tbl_shop_stockTaking_Detail"
    sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx
    dbget.Execute sqlStr

    sqlStr = " delete from db_shop.dbo.tbl_shop_stockTaking_Master"
    sqlStr = sqlStr & " where stTakingIdx="&stTakingIdx
    dbget.Execute sqlStr

    alertStr = "���� �Ǿ����ϴ�."

else
    response.write "<script type='text/javascript'>alert('���� ���� �ʾҽ��ϴ�. - " & mode & "');</script>"
end if
%>

<script type='text/javascript'>
	alert('<%=CHKIIF(alertStr<>"",alertStr,"���� �Ǿ����ϴ�.")%>');

	<% if (mode="stockTakingNext")  then %>
	    opener.location.reload();
	    window.close();
	<% else %>
		location.replace('<%= refer %>');
	<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
