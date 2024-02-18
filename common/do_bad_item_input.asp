<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,itemgubun,itemid,itemoption, itemcount
dim sqlStr, found, i

dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr


mode = requestCheckVar(request("mode"),32)
if (Len(request("itemid")) = 12) then
        itemgubun = Mid(request("itemid"), 1, 2)
        itemid = Mid(request("itemid"), 3, 6)
        itemoption = Mid(request("itemid"), 9, 4)
        itemcount = request("itemcount")
end if
if (Len(request("itemid")) = 14) then
        itemgubun = Mid(request("itemid"), 1, 2)
        itemid = Mid(request("itemid"), 3, 8)
        itemoption = Mid(request("itemid"), 11, 4)
        itemcount = request("itemcount")
end if

itemgubunarr = request.form("itemgubunarr")
itemidarr = request.form("itemidarr")
itemoptionarr = request.form("itemoptionarr")
itemnoarr = request.form("itemnoarr")

itemgubunarr = split(itemgubunarr, "|")
itemidarr = split(itemidarr, "|")
itemoptionarr = split(itemoptionarr, "|")
itemnoarr = split(itemnoarr, "|")

if (mode = "insert") then
        '상품존재여부체크
        if (CStr(itemgubun) = "10") then

			sqlStr = " select count(i.itemid) as cnt from "
			sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
			sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
			sqlStr = sqlStr + " on i.itemid=o.itemid"
			sqlStr = sqlStr + " where i.itemid=" + CStr(itemid)
			sqlStr = sqlStr + " and IsNULL(o.itemoption,'0000')='"+ itemoption + "'"

	        rsget.Open sqlStr,dbget,1
            	found = rsget("cnt")>0
        	rsget.close
        else
            sqlStr = " select count(shopitemid) as cnt from [db_shop].[dbo].tbl_shop_item "
            sqlStr = sqlStr + " where shopitemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "

            rsget.Open sqlStr,dbget,1
        		found = rsget("cnt")>0
        	rsget.close
        end if

        if (found) then
            sqlStr = " select isnull(sum(itemno), 0) as itemno from [db_summary].[dbo].tbl_temp_baditem "
            sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "

            rsget.Open sqlStr,dbget,1
        	if  not rsget.EOF  then
        		itemcount = itemcount + rsget("itemno")
        	end if
        	rsget.close

        	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem "
        	sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "
        	rsget.Open sqlStr,dbget,1

        	sqlStr = "insert into [db_summary].[dbo].tbl_temp_baditem(itemgubun, itemid, itemoption, itemno) "
        	sqlStr = sqlStr + " values('" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "', " + CStr(itemcount) + ") "
        	rsget.Open sqlStr,dbget,1
        else
                response.write "<script>alert('존재하지 않는 상품입니다.');</script>"
        end if
end if

if (mode = "arrinsert") then
        for i = 0 to ubound(itemgubunarr)
                if (itemgubunarr(i) <> "") then
                        '상품존재여부체크
                        if (CStr(itemgubunarr(i)) = "10") then

                        	sqlStr = " select count(i.itemid) as cnt from "
							sqlStr = sqlStr + " [db_item].[dbo].tbl_item i"
							sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option o"
							sqlStr = sqlStr + " on i.itemid=o.itemid"
							sqlStr = sqlStr + " where i.itemid=" + CStr(itemidarr(i))
							sqlStr = sqlStr + " and IsNULL(o.itemoption,'0000')='"+ itemoptionarr(i) + "'"

                            rsget.Open sqlStr,dbget,1
                            	found = rsget("cnt")>0
        					rsget.close
                        else
                            sqlStr = " select count(shopitemid) as cnt from [db_shop].[dbo].tbl_shop_item "
                            sqlStr = sqlStr + " where shopitemid = " + CStr(itemidarr(i)) + " and itemoption = '" + CStr(itemoptionarr(i)) + "' and itemgubun = '" + CStr(itemgubunarr(i)) + "' "
                            rsget.Open sqlStr,dbget,1
                            	found = rsget("cnt")>0
        					rsget.close
                        end if

                        if (found) then
                            sqlStr = " select isnull(sum(itemno), 0) as itemno from [db_summary].[dbo].tbl_temp_baditem "
                            sqlStr = sqlStr + " where itemid = " + CStr(itemidarr(i)) + " and itemoption = '" + CStr(itemoptionarr(i)) + "' and itemgubun = '" + CStr(itemgubunarr(i)) + "' "
                            rsget.Open sqlStr,dbget,1
                        	if  not rsget.EOF  then
                        		itemnoarr(i) = itemnoarr(i) + rsget("itemno")
                        	end if
                        	rsget.close

                        	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem "
                        	sqlStr = sqlStr + " where itemid = " + CStr(itemidarr(i)) + " and itemoption = '" + CStr(itemoptionarr(i)) + "' and itemgubun = '" + CStr(itemgubunarr(i)) + "' "
                        	rsget.Open sqlStr,dbget,1

                        	sqlStr = "insert into [db_summary].[dbo].tbl_temp_baditem(itemgubun, itemid, itemoption, itemno) "
                        	sqlStr = sqlStr + " values('" + CStr(itemgubunarr(i)) + "', " + CStr(itemidarr(i)) + ", '" + CStr(itemoptionarr(i)) + "', " + CStr(itemnoarr(i)) + ") "
                        	rsget.Open sqlStr,dbget,1
                        else
                                response.write "<script>alert('존재하지 않는 상품입니다.');</script>"
                        end if
                end if
        next
end if

if (mode = "modify") then
	sqlStr = "update [db_summary].[dbo].tbl_temp_baditem set itemno = " + CStr(itemcount) + " "
	sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "
	rsget.Open sqlStr,dbget,1
elseif (mode = "delete") then
	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem "
	sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and itemgubun = '" + CStr(itemgubun) + "' "
	rsget.Open sqlStr,dbget,1
end if

%>
<script language="javascript">
// alert('저장 되었습니다.');
<% if refer<>"" then %>
location.replace('<%= refer %>');
<% else %>
location.replace('/common/pop_baditem_input.asp');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
