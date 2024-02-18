<%

'10년동안 한클래스 쓰기운동!!

'==============================================================================
class CItemDetail
    public Fitemid

    public Flarge,  Flarge_nm
    public Fmid,    Fmid_nm
    public Fsmall,  Fsmall_nm

    public Fitemdiv
    public Fitemname
    public Fitemcontent
    public Fdesignercomment
    public Fitemsource
    public Fitemsize
    public Fsellcash
    public Fsellvat
    public Fbuycash
    public Fbuyvat
    public Fdeliverytype
    public Fsourcearea
    public Fmakerid
    public Fmakername
    public Foptioncnt
    public Flimityn
    public Flimitno
    public Flimitsold
    public Fvatinclude
    public Fpojangok

    public FMargin
    public FMileage
    public Fsellyn
    public Fisusing

    public Fitemgubun
    public Fstylegubun
    public Fitemstyle
    public Fusinghtml
    public Fkeywords
    public Fmwdiv
    public Fordercomment
    public Fdeliverarea
    public Fdeliverfixday
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fsailyn

    public Fregdate
    public FLinkitemid

    public Fimgtitle
    public Fimgmain
    public Fimgbasic
    public Fimgsmall
    public Fimglist
    public Flistimage120
    public Ficon1
    public Ficon2
    public Fimgadd
    public Fitemaddcontent
    public FImgStory
    public Finfoimage

	public FitemCouponYn
	public FitemCouponType
	public FitemCouponValue
	public Fcouponbuyprice
	public Freipgoitemyn

	public function IsSoldOut()
		ISsoldOut = (FDispyn<>"Y") or (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

    public function getMwDiv()
        if (IsNull(Fmaeipdiv) or (Fmaeipdiv="")) then
            getMwDiv = Fmaeipdiv
        else
            getMwDiv = Fmaeipdiv
        end if
    end function
end Class

'==============================================================================
class CItem
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectMakerid
    public FRectItemid
    public FRectItemName
    public FRectSellYN
    public FRectIsUsing
    public FRectDanjongyn
    public FRectMWDiv
    public FRectLimityn

    Private Sub Class_Initialize()
        redim FItemList(0)

        FCurrPage       = 1
        FPageSize       = 50
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub

    Private Sub Class_Terminate()

    End Sub

	public function GetImageFolerName(byval i)
	    GetImageFolerName = GetImageSubFolderByItemid(FItemList(i).FItemID)
		''GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public function GetImageFolerNameByItemid(byval itemid)
	    GetImageFolerNameByItemid = GetImageSubFolderByItemid(itemid)
		''GetImageFolerNameByItemid = "0" + CStr(Clng(itemid\10000))
	end function

	public function GetImageAddByIndex(byval index)
		dim arr
		arr = Split(FOneItem.Fimgadd, ",")
		if (UBound(arr) < (index - 1)) then
		    GetImageAddByIndex = ""
		elseif (index < 1) then
		    GetImageAddByIndex = ""
		elseif (arr(index - 1) = "") then
		    GetImageAddByIndex = ""
		else
		    GetImageAddByIndex = "http://webimage.10x10.co.kr/image/add" + CStr(index) + "/" + GetImageFolerNameByItemid(FOneItem.Fitemid) + "/" + arr(index - 1)
		end if
	end function

	public function GetImageContentByIndex(byval index)
		dim arr
		arr = Split(FOneItem.Fitemaddcontent, "|")
		if (UBound(arr) < (index - 1)) then
		    GetImageContentByIndex = ""
		elseif (index < 1) then
		    GetImageContentByIndex = ""
		elseif (arr(index - 1) = "") then
		    GetImageContentByIndex = ""
		else
		    GetImageContentByIndex = db2html(arr(index - 1))
		end if
	end function

    public sub GetProductList()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectIsUsing + "'"
        end if
        
        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if        

        

		'// 결과수 카운트
		sqlStr = "select count(i.itemid) as cnt"
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
        sqlStr = sqlStr & " where i.itemid<>0" & addSql

        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close


        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " i.itemid, i.makerid, i.itemname, i.sellcash, i.buycash, i.sailprice, i.sailsuplycash, i.orgprice, i.orgsuplycash, "
        sqlStr = sqlStr & " i.sailyn, i.sellyn, i.isusing, i.danjongyn, i.mwdiv, i.limityn, i.limitno, i.limitsold, i.vatinclude, "
        sqlStr = sqlStr & " i.itemCouponyn, i.itemCoupontype, i.itemCouponvalue, "
        sqlStr = sqlStr & " IsNull(s.makername,'') as makername, i.regdate, IsNull(i.smallimage,'') as imgsmall, "
        sqlStr = sqlStr & " Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i "
        sqlStr = sqlStr & "    left join [db_item].[dbo].tbl_item_Contents s"
        sqlStr = sqlStr & "    on i.itemid=s.itemid"
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & " and i.itemid<>0" & addSql

        sqlStr = sqlStr & " order by i.itemid desc"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FResultCount =  rsget.RecordCount - (FPageSize*(FCurrPage-1))

        FTotalPage = CInt(FTotalCount\FPageSize) + 1


        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new CItemDetail

                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fmakerid       = db2html(rsget("makerid"))
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fsellcash      = rsget("sellcash")
                FItemList(i).Fbuycash       = rsget("buycash")
	            FItemList(i).Fsailprice     = rsget("sailprice")
	            FItemList(i).Fsailsuplycash = rsget("sailsuplycash")
				FItemList(i).Forgprice     = rsget("orgprice")
	            FItemList(i).Forgsuplycash = rsget("orgsuplycash")

                FItemList(i).Fmakername     = db2html(rsget("makername"))
                FItemList(i).Fregdate       = rsget("regdate")

                FItemList(i).Fsailyn		= rsget("sailyn")		'할인여부
                FItemList(i).Fsellyn        = rsget("sellyn")		'판매
                FItemList(i).Fisusing       = rsget("isusing")		'사용
                FItemList(i).Fmwdiv         = rsget("mwdiv")		'계약구분
                FItemList(i).Flimityn       = rsget("limityn")		'한정
                FItemList(i).Flimitno       = rsget("limitno")
                FItemList(i).Flimitsold     = rsget("limitsold")
                FItemList(i).Fvatinclude	= rsget("vatinclude")	'과세부분

				FItemList(i).FitemCouponYn	  = rsget("itemCouponyn")		'쿠폰적용여부
				FItemList(i).FitemCouponType  = rsget("itemCoupontype")	'쿠폰종류(1:퍼센트,2:원,3:배송비무료)
				FItemList(i).FitemCouponValue = rsget("itemCouponvalue")	'쿠폰값
				
				FItemList(i).Fcouponbuyprice  = rsget("couponbuyprice")	'쿠폰적용 매입가
                if (FItemList(i).Fcouponbuyprice=0) then FItemList(i).Fcouponbuyprice= FItemList(i).Fbuycash
                
                FItemList(i).Fimgsmall      = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("imgsmall")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end sub

    public sub GetProductOne()
        dim sqlStr, i



        sqlStr =    "select top 1 t1.itemid, t1.cate_large, t1.cate_mid, t1.cate_small, t1.itemdiv, t1.itemname " &_
                    "   , t2.code_nm as large_nm, t3.code_nm as mid_nm, t4.code_nm as small_nm " &_
                    "   , t1.itemcontent, t1.designercomment, t1.itemsource, t1.itemsize " &_
                    "   , t1.sellcash, t1.sellvat, t1.buycash, t1.buyvat, t1.margin, t1.mileage, t1.sellyn, t1.dispyn, t1.isusing, t1.reipgoitemyn " &_
                    "   , t1.deliverytype, t1.sourcearea, t1.makerid, t1.makername, t1.limityn, t1.limitno, t1.limitsold " &_
                    "   , t1.vatinclude, t1.pojangok, t1.usinghtml, t1.keywords, t1.orgprice, t1.orgsuplycash " &_
                    "   , t1.sailprice, t1.sailsuplycash, t1.sailyn, t1.mwdiv, t1.ordercomment, t1.deliverarea, t1.deliverfixday, t1.optioncnt " &_
                    "   , t1.titleimage,t1.mainimage,t1.smallimage,t1.listimage,t1.listimage120,t1.basicimage,t1.icon1image,t1.icon2image,t1.addimage,IsNULL(t1.imagecontent,'||||') as imagecontent,t1.storyimage" &_
                    " from [db_item].[dbo].tbl_item as t1 " &_
                    "       Join db_item.[dbo].tbl_item_large as t2 on t1.cate_large=t2.code_large " &_
                    "       Join db_item.[dbo].tbl_item_mid as t3 on t1.cate_mid=t3.code_mid " &_
                    "       Join db_item.[dbo].tbl_item_small as t4 on t1.cate_small=t4.code_small " &_
                    " where 1 = 1 " &_
                    "       and t1.itemid='" + FRectItemid + "' " &_
                    "       and t1.cate_large=t3.code_large " &_
                    "       and t1.cate_large=t4.code_large " &_
                    "       and t1.cate_mid=t4.code_mid "

        if (FRectMakerid <> "") then
            sqlStr = sqlStr & " and t1.makerid='" + FRectMakerid + "'"
        end if
				
				'response.write sqlStr
        rsget.Open sqlStr,dbget,1

        if  not rsget.EOF  then
            FTotalCount = 1
            FResultCount = 1

            set FOneItem = new CItemDetail

            FOneItem.Fitemid        = rsget("itemid")
            FOneItem.Flarge         = rsget("cate_large")
            FOneItem.Fmid           = rsget("cate_mid")
            FOneItem.Fsmall         = rsget("cate_small")
            FOneItem.Flarge_nm      = db2html(rsget("large_nm"))
            FOneItem.Fmid_nm        = db2html(rsget("mid_nm"))
            FOneItem.Fsmall_nm      = db2html(rsget("small_nm"))
            FOneItem.Fitemdiv       = rsget("itemdiv")
            FOneItem.Fitemname      = db2html(rsget("itemname"))
            FOneItem.Fitemcontent   = db2html(rsget("itemcontent"))
            FOneItem.Fdesignercomment= db2html(rsget("designercomment"))
            FOneItem.Fitemsource    = db2html(rsget("itemsource"))
            FOneItem.Fitemsize      = db2html(rsget("itemsize"))
            FOneItem.Fsellcash      = db2html(rsget("sellcash"))
            FOneItem.Fbuycash       = db2html(rsget("buycash"))
            FOneItem.Fsellvat       = rsget("sellvat")
            FOneItem.Fbuyvat        = rsget("buyvat")

            FOneItem.FMargin        = rsget("margin")
            ''수정
            if (FOneItem.Fsellcash<>0) then
            	FOneItem.FMargin		= 100-CLng(FOneItem.Fbuycash/FOneItem.Fsellcash*100*100)/100
            end if

            FOneItem.FMileage       = rsget("mileage")
            FOneItem.Fsellyn        = rsget("sellyn")
            FOneItem.Fdispyn        = rsget("dispyn")
            FOneItem.Fisusing       = rsget("isusing")
            FOneItem.Fdeliverytype  = rsget("deliverytype")
            FOneItem.Fsourcearea    = db2html(rsget("sourcearea"))
            FOneItem.Fmakerid       = db2html(rsget("makerid"))
            FOneItem.Fmakername     = db2html(rsget("makername"))
            FOneItem.Foptioncnt     = rsget("optioncnt")
            FOneItem.Flimityn       = rsget("limityn")
            FOneItem.Flimitno       = rsget("limitno")
            FOneItem.Flimitsold     = rsget("limitsold")
            FOneItem.Fvatinclude    = rsget("vatinclude")
            FOneItem.Fpojangok      = rsget("pojangok")
            FOneItem.Fusinghtml     = rsget("usinghtml")
            FOneItem.Fkeywords      = db2html(rsget("keywords"))
            FOneItem.Fmwdiv         = rsget("mwdiv")
            FOneItem.Fordercomment  = db2html(rsget("ordercomment"))
            FOneItem.Fdeliverarea   = rsget("deliverarea")
            FOneItem.Fdeliverfixday = rsget("deliverfixday")
            FOneItem.Forgprice      = rsget("orgprice")
            FOneItem.Forgsuplycash  = rsget("orgsuplycash")
            FOneItem.Fsailprice     = rsget("sailprice")
            FOneItem.Fsailsuplycash = rsget("sailsuplycash")
            FOneItem.Fsailyn        = rsget("sailyn")
            FOneItem.Freipgoitemyn	= rsget("reipgoitemyn")

            if (rsget("titleimage") = "") then
                FOneItem.Fimgtitle = ""
            else
                FOneItem.Fimgtitle      = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("titleimage")
            end if

            if (rsget("mainimage") = "") then
                FOneItem.Fimgmain = ""
            else
                FOneItem.Fimgmain       = "http://webimage.10x10.co.kr/image/main/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("mainimage")
            end if

            if (rsget("basicimage") = "") then
                FOneItem.Fimgbasic = ""
            else
                FOneItem.Fimgbasic      = "http://webimage.10x10.co.kr/image/basic/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("basicimage")
            end if

            if (rsget("icon1image") = "") then
                FOneItem.Ficon1 = ""
            else
                FOneItem.Ficon1         = "http://webimage.10x10.co.kr/image/icon1/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("icon1image")
            end if
            
            if (rsget("listimage120") = "") then
                FOneItem.Flistimage120 = ""
            else
                FOneItem.Flistimage120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("listimage120")
            end if
            
            if (rsget("icon2image") = "") then
                FOneItem.Ficon2 = ""
            else
                FOneItem.Ficon2         = "http://webimage.10x10.co.kr/image/icon2/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("icon2image")
            end if

            if (rsget("smallimage") = "") then
                FOneItem.Fimgsmall = ""
            else
                FOneItem.Fimgsmall      = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("smallimage")
            end if

            if (rsget("listimage") = "") then
                FOneItem.Fimglist = ""
            else
                FOneItem.Fimglist       = "http://webimage.10x10.co.kr/image/list/" + GetImageFolerNameByItemid(rsget("itemid")) + "/" + rsget("listimage")
            end if

            FOneItem.Fimgadd        = rsget("addimage")
            FOneItem.Fitemaddcontent= rsget("imagecontent")
        else
            FTotalCount = 0
            FResultCount = 0
        end if
        rsget.close
    end sub

    public Function HasPreScroll()
        HasPreScroll = StartScrollPage > 1
    end Function

    public Function HasNextScroll()
        HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
    end Function

    public Function StartScrollPage()
        StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function
end Class

'// 구분에 따른 문자열 색상 지정
function fnColor(str, div)
	Select Case div
		Case "yn"
			if str="N" or isNull(str) then
				fnColor = "<Font color=#F08050>" & str & "</font>"
			else
				fnColor = "<Font color=#5080F0>" & str & "</font>"
			end if
		Case "mw"
			Select Case str
				Case "M"
					fnColor = "<Font color=#F08050>매입</font>"
				Case "W"
					fnColor = "<Font color=#808080>위탁</font>"
				Case "U"
					fnColor = "<Font color=#5080F0>업체</font>"
			end Select
		Case "tx"
			if str="Y" then
				fnColor = "<Font color=#808080>과세</font>"
			else
				fnColor = "<Font color=#F08050>면세</font>"
			end if
	end Select
end Function

'// 마진률계산
Function fnPercent(oup,inp,pnt)
	if oup=0 or isNull(oup) then exit function
	if inp=0 or isNull(inp) then exit function
	fnPercent = FormatNumber((1-(clng(oup)/clng(inp)))*100,pnt) & "%"
End Function
%>