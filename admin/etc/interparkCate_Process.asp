<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/classes/items/extsiteitemcls.asp"-->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")


dim mode, tecdl, tecdm, tecdn
dim interparkdispcategory, SupplyCtrtSeq, interparkstorecategory
mode = request("mode")
tecdl = request("tecdl")
tecdm = request("tecdm")
tecdn = request("tecdn")

interparkdispcategory   = request("interparkdispcategory")
SupplyCtrtSeq           = request("SupplyCtrtSeq")
interparkstorecategory  = request("interparkstorecategory")

dim sqlStr
dim oldDispCate
if (mode="cateedit") then
    ''ī�װ��� ����� ��� �����ؾ��� -> ������ ����
    oldDispCate = ""
    
    sqlStr = "select interparkdispcategory from [db_item].[dbo].tbl_interpark_dspcategory_mapping"
    sqlStr = sqlStr + " where tencdl='" + tecdl + "'"
    sqlStr = sqlStr + " and tencdm='" + tecdm + "'"
    sqlStr = sqlStr + " and tencdn='" + tecdn + "'"
    
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        oldDispCate = rsget("interparkdispcategory")
    end if    
    rsget.Close
    
    sqlStr = "If Exists(select * from [db_item].[dbo].tbl_interpark_dspcategory_mapping "
    sqlStr = sqlStr + " where tencdl='" + tecdl + "'"
    sqlStr = sqlStr + " and tencdm='" + tecdm + "'"
    sqlStr = sqlStr + " and tencdn='" + tecdn + "'"
    sqlStr = sqlStr + " )"
    sqlStr = sqlStr + " BEGIN"
    sqlStr = sqlStr + "     update [db_item].[dbo].tbl_interpark_dspcategory_mapping "
    sqlStr = sqlStr + "     set interparkdispcategory='" + interparkdispcategory + "'"
    sqlStr = sqlStr + "     ,SupplyCtrtSeq=" + SupplyCtrtSeq + ""
    sqlStr = sqlStr + "     ,interparkstorecategory='" + interparkstorecategory + "'"
    sqlStr = sqlStr + "     where tencdl='" + tecdl + "'"
    sqlStr = sqlStr + "     and tencdm='" + tecdm + "'"
    sqlStr = sqlStr + "     and tencdn='" + tecdn + "'"
    sqlStr = sqlStr + " END"
    sqlStr = sqlStr + " ELSE"
    sqlStr = sqlStr + " BEGIN"
    sqlStr = sqlStr + "     insert into [db_item].[dbo].tbl_interpark_dspcategory_mapping "
    sqlStr = sqlStr + "     (tencdl, tencdm, tencdn, interparkdispcategory, SupplyCtrtSeq, interparkstorecategory) "
    sqlStr = sqlStr + "     values("
    sqlStr = sqlStr + "     '" + tecdl + "'"
    sqlStr = sqlStr + "     ,'" + tecdm + "'"
    sqlStr = sqlStr + "     ,'" + tecdn + "'"
    sqlStr = sqlStr + "     ,'" + interparkdispcategory + "'"
    sqlStr = sqlStr + "     ," + SupplyCtrtSeq + ""
    sqlStr = sqlStr + "     ,'" + interparkstorecategory + "'"
    sqlStr = sqlStr + "     )"
    sqlStr = sqlStr + " END"
    
    dbget.Execute sqlStr
    
    ''���� ī�װ��� ����� ���
    if (oldDispCate<>"") and (oldDispCate<>interparkdispcategory) then
        sqlStr = "update [db_item].[dbo].tbl_interpark_reg_item"
        sqlStr = sqlStr + " set interparklastupdate='2008-01-01'"
        sqlStr = sqlStr + " where itemid in ("
        sqlStr = sqlStr + "	select top 500 r.itemid from [db_item].[dbo].tbl_interpark_reg_item r,"
        sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
        sqlStr = sqlStr + "	[db_item].[dbo].tbl_interpark_dspcategory_mapping p"
        sqlStr = sqlStr + "	where r.itemid=i.itemid"
        sqlStr = sqlStr + "	and p.interparkdispcategory='" & interparkdispcategory & "'"
        sqlStr = sqlStr + "	and p.tencdl=i.cate_large"
        sqlStr = sqlStr + "	and p.tencdm=i.cate_mid"
        sqlStr = sqlStr + "	and p.tencdn=i.cate_small"
        sqlStr = sqlStr + ")"
        
        dbget.Execute sqlStr
        
        '''ī�װ��� ����Ǿ .. ������ �ǵ���..
        ''''��ǰ�� ���� - ipark��ǰ�ʿ� ����ī�װ� ����. interparkSupplyCtrtSeq �� �����ؾ���.. �ٲ�� ��ǰ������Ʈ �ȵ�.
        ''' 2011-04-21  ��ǰ ��Ͻÿ��� �ʿ��ҵ�.. => ��� ���� ���μ����ʿ�..
        sqlStr = " update R"
        sqlStr = sqlStr + " set interparkSupplyCtrtSeq=D.SupplyCtrtSeq"
        sqlStr = sqlStr + " ,interparkStoreCategory=D.interparkStoreCategory"
        sqlStr = sqlStr + " , Pinterparkdispcategory=D.interparkdispcategory"
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_interpark_reg_item R"
        sqlStr = sqlStr + " 	Join [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr + " 	on R.itemid=i.itemid"
        sqlStr = sqlStr + " 	left join [db_item].[dbo].tbl_interpark_dspcategory_mapping D"
        sqlStr = sqlStr + " 	on D.tencdl=i.cate_large"
        sqlStr = sqlStr + " 	and D.tencdm=i.cate_mid"
        sqlStr = sqlStr + " 	and D.tencdn=i.cate_small"
        sqlStr = sqlStr + " where IsNULL(R.interparkSupplyCtrtSeq,D.SupplyCtrtSeq)=D.SupplyCtrtSeq"
        sqlStr = sqlStr + " and D.SupplyCtrtSeq is Not NULL"
        sqlStr = sqlStr + " and i.cate_large='" + tecdl + "'"
        sqlStr = sqlStr + " and i.cate_mid='" + tecdm + "'"
        sqlStr = sqlStr + " and i.cate_small='" + tecdn + "'"
        sqlStr = sqlStr + " and R.interParkPrdNo is Not NULL"

        dbget.Execute sqlStr
    end if

end if

%>
<script language='javascript'>
alert('����Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->