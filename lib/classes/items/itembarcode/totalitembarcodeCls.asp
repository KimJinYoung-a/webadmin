<%
class CTotalItemBarCodeSubItem
	public FsiteSeq
	public FsiteItemGubun
    public FsiteItemid
    public FsiteItemOption
	public Fmakerid
	public FbrandName
    public FsiteItemName
    public FsiteItemOptionName
	public FsiteItemOptionName1
	public FsiteItemOptionName2
    public Fpublicbarcode
    public Flocalbarcode
    public FitemRackcode
    public ForgSellPrice
	public FImageSmall
	public FImageList
	public FImageBasic
	public FOptionUseYN
    public Fregdate
    public Flastupdate
    public FOptaddprice
    public Fupchemanagecode


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CTotalItemBarCode
	public FItemList()
	public FOneItem
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectSiteSeq
	public FRectCompanyid3PL
	public FRectItemGubun
	public FRectItemID
	public FRectItemoption
	public FRectBarcode



	public Sub getTotalItemCodeSearch()
		Dim i, sqlStr, sqlsearch

		sqlStr = "select S.itemgubun, S.itemid, S.itemoption"
		sqlStr = sqlstr + " FROM [db_item].[dbo].[tbl_item_option_stock] AS S"
		sqlStr = sqlstr + " WHERE S.barcode = '" & FRectBarcode & "'"
		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1
			set FOneItem = new CTotalItemBarCodeSubItem
						if Not rsget.Eof then
				FOneItem.FsiteSeq        		= rsget("itemgubun")
				FOneItem.FsiteItemGubun 		= rsget("itemgubun")
				FOneItem.FsiteItemid    		= rsget("itemid")
				FOneItem.FsiteItemOption   		= rsget("itemoption")

		end if

		rsget.Close
	end Sub


	public Sub getTotalItemBarcodeON()
		dim sqlStr,i
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash, i.brandname, i.smallimage, i.listimage, i.basicimage, i.itemrackcode," + VbCrlf
		sqlStr = sqlStr + " IsNULL(v.itemoption,'0000') as itemoption, '' as opt1name," + VbCrlf
		sqlStr = sqlStr + " IsNULL(v.optionname,'') as opt2name, IsNULL(v.isusing,'N') as optusing, " + VbCrlf
		sqlStr = sqlStr + " s.barcode, IsNull(v.optaddprice,0) as optaddprice"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s on (s.itemgubun='10') and i.itemid=s.itemid  and IsNULL(s.itemoption,'0000')=IsNULL(v.itemoption,'0000')" + VbCrlf
		sqlStr = sqlStr + " where i.itemid=" + CStr(FRectItemID) + VbCrlf
'rw sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim  FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTotalItemBarCodeSubItem

				FItemList(i).FsiteSeq  				= "10"

				FItemList(i).FsiteItemGubun  		= "10"
				FItemList(i).FsiteItemid     		= rsget("itemid")
				FItemList(i).FsiteItemOption 		= rsget("itemoption")

				FItemList(i).Fmakerid        		= rsget("makerid")
				FItemList(i).FbrandName      		= db2html(rsget("brandname"))

				FItemList(i).FsiteItemName      	= db2html(rsget("itemname"))
				FItemList(i).FsiteItemOptionName 	= rsget("opt2name")

				FItemList(i).FsiteItemOptionName1 	= rsget("opt1name")
				FItemList(i).FsiteItemOptionName2 	= rsget("opt2name")

                FItemList(i).Fpublicbarcode    		= rsget("barcode")
                if (FItemList(i).FsiteItemid>=1000000) then
                    FItemList(i).Flocalbarcode     		= CStr(FItemList(i).FsiteItemGubun) & Right(("100000000" & FItemList(i).FsiteItemid), 8) & FItemList(i).FsiteItemOption
                else
                    FItemList(i).Flocalbarcode     		= CStr(FItemList(i).FsiteItemGubun) & Right(("1000000" & FItemList(i).FsiteItemid), 6) & FItemList(i).FsiteItemOption
                end if
                FItemList(i).FitemRackcode     		= rsget("itemRackcode")

				FItemList(i).ForgSellPrice       	= rsget("sellcash")


				FItemList(i).FImageSmall     		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("smallimage")
				FItemList(i).FImageList      		= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("listimage")
				FItemList(i).FImageBasic      		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("basicimage")

				FItemList(i).FOptionUseYN 			= rsget("optusing")
                FItemList(i).Fregdate          		= ""
                FItemList(i).Flastupdate       		= ""
                FItemList(i).FOptaddprice			= rsget("optaddprice")
                If FItemList(i).FOptaddprice = "0" Then
                	FItemList(i).FOptaddprice = ""
                Else
                	FItemList(i).FOptaddprice = "(+" & FItemList(i).FOptaddprice & ")"
                End If

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	public Sub getTotalItemBarcodeOFF()
		Dim i, sqlStr, sqlsearch

		If FRectItemGubun <> "" Then
			sqlsearch = sqlsearch & " and i.itemgubun = '" & FRectItemGubun & "' "
		End If

		If FRectItemID <> "" Then
			sqlsearch = sqlsearch & " and i.shopitemid = '" & FRectItemID & "' "
		End If

		If FRectItemoption <> "" Then
			sqlsearch = sqlsearch & " and i.itemoption = '" & FRectItemoption & "' "
		End If

		sqlStr = "select top 1"
		sqlStr = sqlstr + " i.itemgubun, i.shopitemid, i.shopitemname, i.makerid, i.shopitemprice"
		sqlStr = sqlstr + " , i.isusing, i.regdate, i.centermwdiv as mwdiv"
		sqlStr = sqlstr + " ,i.offimgmain ,i.offimglist ,i.offimgsmall , i.itemoption, i.shopitemoptionname"
		sqlstr = sqlstr + " ,IsNull(sm.realstock,0) as realstock"
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv5,0) as ipkumdiv5"
		sqlstr = sqlstr + " ,IsNull(sm.offconfirmno,0) as offconfirmno"
		sqlstr = sqlstr + " ,sm.lastupdate, i.extbarcode, i.orgsellprice"
		sqlStr = sqlstr + " from db_shop.dbo.tbl_shop_item i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " 	on i.itemgubun = sm.itemgubun"
		sqlstr = sqlstr + " 	and i.shopitemid=sm.itemid"
		sqlstr = sqlstr + " 	and i.itemoption=sm.itemoption"
		sqlStr = sqlstr + " where 1=1 " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FItemList(0) = new CTotalItemBarCodeSubItem

				FItemList(0).FsiteSeq        		= rsget("itemgubun")
				FItemList(0).FsiteItemGubun 		= rsget("itemgubun")
				FItemList(0).FsiteItemid    		= rsget("shopitemid")
				FItemList(0).FsiteItemOption   		= rsget("itemoption")
				FItemList(0).FMakerID				= rsget("makerid")
				FItemList(0).FbrandName      		= "" 'db2html(rsget("brandname"))
				FItemList(0).FsiteItemName      	= db2html(rsget("shopitemname"))
				FItemList(0).FsiteItemOptionName 	= db2html(rsget("shopitemoptionname"))
				FItemList(0).FsiteItemOptionName1 	= ""
				FItemList(0).FsiteItemOptionName2 	= ""
				FItemList(0).Fpublicbarcode    		= rsget("extbarcode")
				FItemList(0).ForgSellPrice       	= rsget("shopitemprice")
				FItemList(0).FImageList	= rsget("offimglist")
				if FItemList(0).FImageList<>"" then FItemList(0).FImageList = webImgUrl + "/offimage/offlist/i" + FItemList(0).FsiteItemGubun + "/" + GetImageSubFolderByItemid(FItemList(0).FsiteItemid) + "/" + FItemList(0).FImageList
				FItemList(0).FOptionUseYN 			= ""
				FItemList(0).Fregdate				= rsget("regdate")
				FItemList(0).Flastupdate			= rsget("lastupdate")
				FItemList(0).FOptaddprice			= ""

		end if

		rsget.Close
	end Sub


	public Sub getTotalUpcheManagecodeON()
		dim sqlStr,i
		sqlStr = "select i.itemid, i.itemname, i.makerid, i.sellcash, i.brandname, i.smallimage, i.listimage, i.basicimage, i.itemrackcode," + VbCrlf
		sqlStr = sqlStr + " IsNULL(v.itemoption,'0000') as itemoption, '' as opt1name," + VbCrlf
		sqlStr = sqlStr + " IsNULL(v.optionname,'') as opt2name, IsNULL(v.isusing,'N') as optusing, " + VbCrlf
		sqlStr = sqlStr + " s.upchemanagecode, IsNull(v.optaddprice,0) as optaddprice"
		sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i" + VbCrlf
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v on i.itemid=v.itemid"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option_stock s on (s.itemgubun='10') and i.itemid=s.itemid  and IsNULL(s.itemoption,'0000')=IsNULL(v.itemoption,'0000')" + VbCrlf
		sqlStr = sqlStr + " where i.itemid=" + CStr(FRectItemID) + VbCrlf
'rw sqlStr
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		redim  FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTotalItemBarCodeSubItem

				FItemList(i).FsiteSeq  				= "10"

				FItemList(i).FsiteItemGubun  		= "10"
				FItemList(i).FsiteItemid     		= rsget("itemid")
				FItemList(i).FsiteItemOption 		= rsget("itemoption")

				FItemList(i).Fmakerid        		= rsget("makerid")
				FItemList(i).FbrandName      		= db2html(rsget("brandname"))

				FItemList(i).FsiteItemName      	= db2html(rsget("itemname"))
				FItemList(i).FsiteItemOptionName 	= rsget("opt2name")

				FItemList(i).FsiteItemOptionName1 	= rsget("opt1name")
				FItemList(i).FsiteItemOptionName2 	= rsget("opt2name")

                FItemList(i).Fupchemanagecode  		= rsget("upchemanagecode")

                FItemList(i).FitemRackcode     		= rsget("itemRackcode")

				FItemList(i).ForgSellPrice       	= rsget("sellcash")


				FItemList(i).FImageSmall     		= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("smallimage")
				FItemList(i).FImageList      		= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("listimage")
				FItemList(i).FImageBasic      		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FRectItemID) + "/" + rsget("basicimage")

				FItemList(i).FOptionUseYN 			= rsget("optusing")
                FItemList(i).Fregdate          		= ""
                FItemList(i).Flastupdate       		= ""
                FItemList(i).FOptaddprice			= rsget("optaddprice")
                If FItemList(i).FOptaddprice = "0" Then
                	FItemList(i).FOptaddprice = ""
                Else
                	FItemList(i).FOptaddprice = "(+" & FItemList(i).FOptaddprice & ")"
                End If

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Sub


	public Sub getTotalUpcheManagecodeOFF()
		Dim i, sqlStr, sqlsearch

		If FRectItemGubun <> "" Then
			sqlsearch = sqlsearch & " and i.itemgubun = '" & FRectItemGubun & "' "
		End If

		If FRectItemID <> "" Then
			sqlsearch = sqlsearch & " and i.shopitemid = '" & FRectItemID & "' "
		End If

		If FRectItemoption <> "" Then
			sqlsearch = sqlsearch & " and i.itemoption = '" & FRectItemoption & "' "
		End If

		sqlStr = "select top 1"
		sqlStr = sqlstr + " i.itemgubun, i.shopitemid, i.shopitemname, i.makerid, i.shopitemprice"
		sqlStr = sqlstr + " , i.isusing, i.regdate, i.centermwdiv as mwdiv"
		sqlStr = sqlstr + " ,i.offimgmain ,i.offimglist ,i.offimgsmall , i.itemoption, i.shopitemoptionname"
		sqlstr = sqlstr + " ,IsNull(sm.realstock,0) as realstock"
		sqlstr = sqlstr + " ,IsNull(sm.ipkumdiv5,0) as ipkumdiv5"
		sqlstr = sqlstr + " ,IsNull(sm.offconfirmno,0) as offconfirmno"
		sqlstr = sqlstr + " ,sm.lastupdate, i.extbarcode, i.orgsellprice"
		sqlStr = sqlstr + " from db_shop.dbo.tbl_shop_item i"
		sqlstr = sqlstr + " left join [db_summary].[dbo].tbl_current_logisstock_summary sm"
		sqlstr = sqlstr + " 	on i.itemgubun = sm.itemgubun"
		sqlstr = sqlstr + " 	and i.shopitemid=sm.itemid"
		sqlstr = sqlstr + " 	and i.itemoption=sm.itemoption"
		sqlStr = sqlstr + " where 1=1 " & sqlsearch

		'response.write sqlStr & "<Br>"
		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount
		if Not rsget.Eof then
			set FItemList(0) = new CTotalItemBarCodeSubItem

				FItemList(0).FsiteSeq        		= rsget("itemgubun")
				FItemList(0).FsiteItemGubun 		= rsget("itemgubun")
				FItemList(0).FsiteItemid    		= rsget("shopitemid")
				FItemList(0).FsiteItemOption   		= rsget("itemoption")
				FItemList(0).FMakerID				= rsget("makerid")
				FItemList(0).FbrandName      		= "" 'db2html(rsget("brandname"))
				FItemList(0).FsiteItemName      	= db2html(rsget("shopitemname"))
				FItemList(0).FsiteItemOptionName 	= db2html(rsget("shopitemoptionname"))
				FItemList(0).FsiteItemOptionName1 	= ""
				FItemList(0).FsiteItemOptionName2 	= ""
				FItemList(0).Fpublicbarcode    		= rsget("extbarcode")
				FItemList(0).ForgSellPrice       	= rsget("shopitemprice")
				FItemList(0).FImageList	= rsget("offimglist")
				if FItemList(0).FImageList<>"" then FItemList(0).FImageList = webImgUrl + "/offimage/offlist/i" + FItemList(0).FsiteItemGubun + "/" + GetImageSubFolderByItemid(FItemList(0).FsiteItemid) + "/" + FItemList(0).FImageList
				FItemList(0).FOptionUseYN 			= ""
				FItemList(0).Fregdate				= rsget("regdate")
				FItemList(0).Flastupdate			= rsget("lastupdate")
				FItemList(0).FOptaddprice			= ""

		end if

		rsget.Close
	end Sub



	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class


Function AutoItemIDSetting(itemid)
	'// 사용하지 말것
	'// (참조 : /lib/BarcodeFunction.asp)
	Dim i, vTmp
	If itemid <> "" Then
		If Len(itemid) < 6 Then
			For i = 1 To (6-(Len(itemid)))
				vTmp = vTmp & "0"
			Next
			AutoItemIDSetting = vTmp & itemid
		Else
			AutoItemIDSetting = itemid
		End If
	End If
End Function


Function AutoItemIDSettingReturn(itemid)
	'// 사용하지 말것
	'// (참조 : /lib/BarcodeFunction.asp)
	Dim i, vTmp
	vTmp = itemid

	For i = 1 To 6
		If Mid(itemid, i, 1) = "0" Then
			vTmp = Right(itemid, Len(vTmp-1))
		End IF
	Next
	AutoItemIDSettingReturn = vTmp
End Function
%>
