<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode,itemgubun,itemid,itemoption, itemcount
dim sqlStr, found, i
dim itembarcode

dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr

itembarcode = requestCheckVar(request("itembarcode"),20)
itemcount   = requestCheckVar(request("itemcount"),10)

mode = request("mode")
if (Len(itembarcode) = 12) then
    itemgubun   = Mid(getNumeric(request("itembarcode")), 1, 2)
    itemid      = Mid(getNumeric(request("itembarcode")), 3, 6)
    itemoption  = Mid(request("itembarcode"), 9, 4)
elseif (Len(itembarcode) = 14) then
    itemgubun   = Mid(getNumeric(request("itembarcode")), 1, 2)
    itemid      = Mid(getNumeric(request("itembarcode")), 3, 8)
    itemoption  = Mid(request("itembarcode"), 11, 4)
elseif (Len(itembarcode)>6) then
    '''바코드인경우 검색후 상품코드 가져옴.
    call fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
end if

if (mode="insert") then
    if (Len(itemgubun)<2) or (Len(itemid)<3) or (Len(itemoption)<4) then
        response.write "<script type='text/javascript'>alert('상품 바코드가 유효하지 않습니다.');</script>"
        response.write "<script type='text/javascript'>location.replace('" & refer & "');</script>"
        dbget.close()	:	response.End
    end if
end if


itemgubunarr    = request.form("itemgubunarr")
itemidarr       = request.form("itemidarr")
itemoptionarr   = request.form("itemoptionarr")
itemnoarr       = request.form("itemnoarr")

itemgubunarr    = split(itemgubunarr, "|")
itemidarr       = split(itemidarr, "|")
itemoptionarr   = split(itemoptionarr, "|")
itemnoarr       = split(itemnoarr, "|")

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

        '''재고디비에 있으면 추가가능.
        if (not found) then
    	    sqlStr = " select count(*) as cnt from "
    	    sqlStr = sqlStr + " db_summary.dbo.tbl_current_logisstock_summary"
    	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)
		    sqlStr = sqlStr + " and itemoption='"+ itemoption + "'"
		    sqlStr = sqlStr + " and itemgubun='"+ itemgubun + "'"

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
							sqlStr = sqlStr + " where i.itemid=" + CStr(requestCheckVar(itemidarr(i),10))
							sqlStr = sqlStr + " and IsNULL(o.itemoption,'0000')='"+ requestCheckVar(itemoptionarr(i),4) + "'"

                            rsget.Open sqlStr,dbget,1
                            	found = rsget("cnt")>0
        					rsget.close
                        else
                            sqlStr = " select count(shopitemid) as cnt from [db_shop].[dbo].tbl_shop_item "
                            sqlStr = sqlStr + " where shopitemid = " + CStr(requestCheckVar(itemidarr(i),10)) + " and itemoption = '" + CStr(requestCheckVar(itemoptionarr(i),4)) + "' and itemgubun = '" + CStr(requestCheckVar(itemgubunarr(i),2)) + "' "
                            rsget.Open sqlStr,dbget,1
                            	found = rsget("cnt")>0
        					rsget.close
                        end if

                        if (found) then
                            sqlStr = " select isnull(sum(itemno), 0) as itemno from [db_summary].[dbo].tbl_temp_baditem "
                            sqlStr = sqlStr + " where itemid = " + CStr(requestCheckVar(itemidarr(i),10)) + " and itemoption = '" + CStr(requestCheckVar(itemoptionarr(i),4)) + "' and itemgubun = '" + CStr(requestCheckVar(itemgubunarr(i),2)) + "' "
                            rsget.Open sqlStr,dbget,1
                        	if  not rsget.EOF  then
                        		itemnoarr(i) = itemnoarr(i) + rsget("itemno")
                        	end if
                        	rsget.close

                        	sqlStr = "delete from [db_summary].[dbo].tbl_temp_baditem "
                        	sqlStr = sqlStr + " where itemid = " + CStr(requestCheckVar(itemidarr(i),10)) + " and itemoption = '" + CStr(requestCheckVar(itemoptionarr(i),4)) + "' and itemgubun = '" + CStr(requestCheckVar(itemgubunarr(i),2)) + "' "
                        	rsget.Open sqlStr,dbget,1

                        	sqlStr = "insert into [db_summary].[dbo].tbl_temp_baditem(itemgubun, itemid, itemoption, itemno) "
                        	sqlStr = sqlStr + " values('" + CStr(requestCheckVar(itemgubunarr(i),2)) + "', " + CStr(requestCheckVar(itemidarr(i),10)) + ", '" + CStr(requestCheckVar(itemoptionarr(i),4)) + "', " + CStr(requestCheckVar(itemnoarr(i),10)) + ") "
                        	rsget.Open sqlStr,dbget,1
                        else
                                response.write "<script type='text/javascript'>alert('존재하지 않는 상품입니다.');</script>"
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
<script type='text/javascript'>
	// alert('저장 되었습니다.');
	<% if refer<>"" then %>
		location.replace('<%= refer %>');
	<% else %>
		location.replace('/common/pop_baditem_input.asp');
	<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
