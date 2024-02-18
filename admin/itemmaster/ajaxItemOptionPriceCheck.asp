<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitemoption, oOptionMultiple
dim k, i, mwdiv, deliverOverseas

itemid = requestCheckvar(getNumeric(request("itemid")),10)
mwdiv = requestCheckvar(request("mwdiv"),1)
deliverOverseas = requestCheckvar(request("deliverOverseas"),1)

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

set oOptionMultiple = new CitemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" then
    oOptionMultiple.GetOptionMultipleInfo
end if

if oitemoption.FResultCount<1 then
    response.write "3"
    dbget.close()	:	response.End
else
    if (oitemoption.IsMultipleOption) then
        for k=0 to oOptionMultiple.FResultCount -1
            if mwdiv<>"U" then
                if oOptionMultiple.FItemList(k).Foptaddprice > 0 then
                    response.write "1"
                    dbget.close()	:	response.End
                end if
            else
                if oOptionMultiple.FItemList(k).Foptaddprice > 0 then
                    if deliverOverseas="Y" then
                        response.write "2"
                        dbget.close()	:	response.End
                    end if
                end if
            end if
        next
        response.write "3"
    else
        for k=0 to oitemoption.FResultCount -1
            if mwdiv<>"U" then
                if oitemoption.FItemList(k).Foptaddprice > 0 then
                    response.write "1"
                    dbget.close()	:	response.End
                end if
            else
                if oitemoption.FItemList(k).Foptaddprice > 0 then
                    if deliverOverseas="Y" then
                        response.write "2"
                        dbget.close()	:	response.End
                    end if
                end if
            end if
        next
        response.write "3"
    end if
end if
set oOptionMultiple = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->