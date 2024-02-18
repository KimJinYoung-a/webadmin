<%@ language=vbscript %>
<% option explicit %>
<% Response.CharSet = "EUC-KR" %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��ǰ�ڵ� üũ
' History : 2022.05.31 �̻� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim sqlStr
dim itemid, makerid, sellcash

itemid = RequestCheckVar(Request("itemid"), 32)
makerid = RequestCheckVar(Request("makerid"), 32)


sqlStr = " select top 1 i.itemid "
sqlStr = sqlStr & " from "
sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_item] i "
sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o on i.itemid = o.itemid "
sqlStr = sqlStr & " where "
sqlStr = sqlStr & " 	1 = 1 "
sqlStr = sqlStr & " 	and IsNull(o.optaddprice, 0) <> 0 "
sqlStr = sqlStr & " 	and i.itemid = " & itemid

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    Response.write "{""result"" :""err"",""message"":""�ɼǰ� �ִ� ��ǰ�Դϴ�.""}"
    rsget.Close
else
    rsget.Close

    sqlStr = " select top 1 i.makerid, i.orgprice as  sellcash "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_item] i "
    sqlStr = sqlStr & " where "
    sqlStr = sqlStr & " 	1 = 1 "
    sqlStr = sqlStr & " 	and i.itemid = " & itemid
    rsget.Open sqlStr, dbget, 1

    if rsget.Eof then
        Response.write "{""result"" :""err"",""message"":""�߸��� ��ǰ�ڵ��Դϴ�.""}"
    elseif makerid <> "" and (makerid <> rsget("makerid")) then
        Response.write "{""result"" :""err"",""message"":""�귣�尡 ��ġ���� �ʽ��ϴ�.""}"
    else
        sellcash = rsget("sellcash")
        rsget.Close

        '��ǰ �ߺ� üũ (2022-12-21 ����)
        sqlStr = "SELECT top 1 B.buy_benefit_idx, B.benefit_title"
        sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_buy_benefit AS B"
        sqlStr = sqlStr & " LEFT JOIN db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group AS G ON B.buy_benefit_idx = G.buy_benefit_idx"
        sqlStr = sqlStr & " LEFT JOIN db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item AS I ON G.benefit_group_no=I.benefit_group_no"
        sqlStr = sqlStr & " WHERE B.benefit_start_dt <= GETDATE()"
        sqlStr = sqlStr & " AND B.benefit_end_dt >= GETDATE()"
        sqlStr = sqlStr & " AND B.use_yn='Y'"
        sqlStr = sqlStr & " AND G.use_yn='Y'"
        sqlStr = sqlStr & " AND I.use_yn='Y'"
        sqlStr = sqlStr & " AND I.itemid =" & itemid
        sqlStr = sqlStr & " ORDER BY B.show_rank, G.sort_no, I.sort_no ASC"
        rsget.Open sqlStr, dbget, 1

        if Not rsget.Eof then
            Response.write "{""result"" :""err"",""message"":""�ε��� ��ȣ " & rsget("buy_benefit_idx") & " " + rsget("benefit_title") + "�� ��ϵ� ��ǰ�� �ߺ��˴ϴ�.""}"
        else
            Response.write "{""result"" :""ok"",""message"":"""",""sellcash"":""" & sellcash &  """}"
        end if
        rsget.Close
    end if
    
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
