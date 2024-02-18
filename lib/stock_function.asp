<%
public function IsValidShopItem(itemgubun, itemid, itemoption, byref imakerid)
    dim sqlStr
    IsValidShopItem = false
    sqlStr = "select top 1  makerid from db_shop.dbo.tbl_shop_item"
    sqlStr = sqlStr + " where itemgubun='"&itemgubun&"'"
    sqlStr = sqlStr + " and shopitemid="&itemid&""
    sqlStr = sqlStr + " and itemoption='"&itemoption&"'"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        IsValidShopItem = true
        imakerid = rsget("makerid")
    end if
    rsget.close
end function

public function getItemCodeByBarcode(byval barcode,byref iitemgubun, byref iitemid, byref iitemoption)
    dim sqlStr
    getItemCodeByBarcode = False

    if (Len(barcode)<8) then Exit function

    sqlStr = "select top 1 b.* " + VbCrlf
    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
    sqlStr = sqlStr + " where b.barcode='" + CStr(barcode) + "' " + VbCrlf

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
    	iitemgubun = rsget("itemgubun")
    	iitemid = rsget("itemid")
    	iitemoption = rsget("itemoption")

    end if
    rsget.Close

    if (Len(CStr(iitemgubun))=2) and (Len(CStr(iitemid))>0) and (Len(CStr(iitemoption))=4) then
        getItemCodeByBarcode = True
        Exit function
    end if

    if (Len(barcode)=12) then
        iitemgubun  = Left(barcode,2)
    	iitemid     = Mid(barcode,3,6)
    	iitemoption = Right(barcode,4)

    	getItemCodeByBarcode = True
    elseif (Len(barcode)=14) then
        iitemgubun  = Left(barcode,2)
    	iitemid     = Mid(barcode,3,8)
    	iitemoption = Right(barcode,4)

    	getItemCodeByBarcode = True

    end if

end function

%>