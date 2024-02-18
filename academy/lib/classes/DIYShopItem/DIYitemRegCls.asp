<%
Class CItemListItems
	public Fitemid
	public Fitemname
	public Fsellcash
	public FSuplyCash
	public Fmakername
	public Fregdate
	public FrejectMsg
	public FrejectDate
	public FreRegMsg
	public FreRegDate

	public Fmakerid

	public FCurrState
	public FLinkitemid
	public FImgSmall
	public FSellyn
    
    public Fupchemanagecode
    
	public function GetCurrStateColor()
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="3" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="8" then
			GetCurrStateColor = "#AAAAAA"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		elseif FCurrState="9" then
			GetCurrStateColor = "#996600"
		elseif FCurrState="0" then
			GetCurrStateColor = "#FF0000"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

	public function GetCurrStateName()
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "등록대기"
		elseif FCurrState="2" then
			GetCurrStateName = "등록보류"
		elseif FCurrState="3" then
			GetCurrStateName = "XXX"            '' 임시저장 전단계 mobileApp 2016/12/08
		elseif FCurrState="7" then
			GetCurrStateName = "등록완료"
		elseif FCurrState="8" then              '' mobileApp 2016/12/08
			GetCurrStateName = "임시저장"
		elseif FCurrState="5" then
			GetCurrStateName = "등록재요청"
		elseif FCurrState="0" then
			GetCurrStateName = "등록불가" ''등록거부
		elseif FCurrState="9" then
			GetCurrStateName = "업체취소"
		else
			GetCurrStateName = ""
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemlist
	public FItemList()

	public FTotalCount
	public FResultCount
	public FRectDesignerID
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount

	public FRectCurrState
	public FRectSellyn
	public FRectItemID
	public FRectLectureYN
	Public FRectitemname

	Private Sub Class_Initialize()
	redim FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	public sub WaitProductList()
		dim sqlStr,i,wheredetail

		if (FRectDesignerID<>"") then
			wheredetail = wheredetail + " and makerid='" + FRectDesignerID + "'"
		end if

		if (FRectCurrState="notreg") then
			wheredetail = wheredetail + " and currstate='1'"
		end if

		if (FRectCurrState="notregwithgubu") then
			wheredetail = wheredetail + " and currstate in ('1','2')"
		end If
        
        if (FRectCurrState="junstnotreged") then
			wheredetail = wheredetail + " and currstate in ('1','2','0','5','8')"  ''2016/12/08 8번추가
		end If
        
		if (FRectitemname<>"") then
			wheredetail = wheredetail + " and itemname like '%" + FRectitemname + "%'"
		end if

		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(itemid) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		sqlStr = sqlStr & wheredetail

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " itemid,makerid,itemname,sellcash,buycash,"
		sqlStr = sqlStr & " linkitemid, currstate, IsNull(makername,'')as maker,regdate, upchemanagecode, rejectmsg, rejectDate"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid<>0"
		sqlStr = sqlStr & " and currstate<9"
		sqlStr = sqlStr & wheredetail
		sqlStr = sqlStr & " order by regdate Desc"


		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount =  rsACADEMYget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CItemListItems
				FItemList(i).Fitemid = rsACADEMYget("itemid")
				FItemList(i).Fmakerid = rsACADEMYget("makerid")
			    FItemList(i).Fitemname = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fsellcash = rsACADEMYget("sellcash")
				FItemList(i).FSuplyCash = rsACADEMYget("buycash")
				FItemList(i).Fmakername = rsACADEMYget("maker")
				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).Frejectmsg = rsACADEMYget("rejectmsg")
				FItemList(i).FrejectDate = rsACADEMYget("rejectDate")

				FItemList(i).FLinkitemid = rsACADEMYget("linkitemid")
				FItemList(i).FCurrState = rsACADEMYget("currstate")
				
				FItemList(i).Fupchemanagecode = db2html(rsACADEMYget("upchemanagecode"))
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	public sub UpdateProductList()
		dim sqlStr,i,wheredetail

		if (FRectDesignerID<>"") then
			wheredetail = wheredetail + " and i.makerid='" + FRectDesignerID + "'"
		end if


		if (FRectSellyn<>"") then
			wheredetail = wheredetail + " and i.sellyn='Y'"
		end if

		if (FRectItemID<>"") then
			wheredetail = wheredetail + " and i.itemid='" + FRectItemID + "'"
		end if



		'###########################################################################
		'등록대기 상품 총 갯수 구하기
		'###########################################################################
		sqlStr = "select count(i.itemid) as cnt"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & wheredetail

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		'###########################################################################
		'등록대기 상품 데이터
		'###########################################################################

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.itemid,i.makerid,i.itemname,i.sellcash,i.buycash,i.sellyn,"
		sqlStr = sqlStr & " IsNull(i.makername,'')as maker, regdate, i.smallimage as imgsmall"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_item i"
		sqlStr = sqlStr & " where i.itemid<>0"
		sqlStr = sqlStr & wheredetail
		sqlStr = sqlStr & " order by regdate Desc"

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FResultCount =  rsACADEMYget.RecordCount - (FPageSize*(FCurrPage-1))

		FTotalPage = CInt(FTotalCount\FPageSize) + 1


		redim preserve FItemList(FResultCount)

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new CItemListItems

				FItemList(i).Fitemid = rsACADEMYget("itemid")
				FItemList(i).Fmakerid = rsACADEMYget("makerid")
			    FItemList(i).Fitemname = db2html(rsACADEMYget("itemname"))
				FItemList(i).Fsellcash = rsACADEMYget("sellcash")
				FItemList(i).FSuplyCash = rsACADEMYget("buycash")
				FItemList(i).Fmakername = rsACADEMYget("maker")
				FItemList(i).Fregdate = rsACADEMYget("regdate")
				FItemList(i).Fsellyn = rsACADEMYget("sellyn")
				FItemList(i).FImgSmall = imgFingers & "/diyItem/waitimage/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsACADEMYget("imgsmall")

				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
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




class CItemOptionItem
	public Fitemoption
	public Fitemoptionname
	public Fisusing
	public Foptsellyn
	public Foptlimityn
	public Foptlimitno
	public Foptlimitsold
	public Fcodeview                'deprecated( Fitemoptionname 으로 변경한다. )

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CWaitItemDetail
'########################################
'임시데이터
'########################################
	public FItemList()
	public FResultCount
	
	public FWaitItemID
	public FMakerid
	public Flarge
	public Fmid
	public Fsmall
	public Fitemdiv
	public Fitemname
	public Fitemcontent
	public Fdesignercomment
	public Fitemsource
	public Fitemsize
	public FitemWeight
	public Fsellcash
	public Fsellvat
	public Fbuycash
	public Fbuyvat
	public Fdeilverytype
	public Fsourcearea
	public Fmakername
	public Flimityn
	public Flimitno

	public FvatYn
	public FMargin
	public FMileage
	public Fsellyn

	public Fusinghtml
	public Fkeywords
	public Fmwdiv
	public Fmaeipdiv
	public Fordercomment
	public Foptioncnt
    
    public FCurrState
    public Frejectmsg
    public FrejectDate
    public FreRegMsg
    public FreRegDate
    
    public FsellEndDate
    public Fupchemanagecode
    
    public FRectDesignerID

	public Fimgtitle
	public Fimgmain
	public Fimgsmall
	public Fimglist
	public Fimgbasic
	public Ficon1
	public Ficon2
	public Fimgadd

	public Fcstodr
	public FrequireMakeDay
	public Frequirecontents
	public Frefundpolicy
	public FinfoDiv
	public FsafetyYn
	public FsafetyDiv
	public FsafetyNum
	public Ffreight_mine
	public Ffreight_max

	Public Frequirechk	'//주문제작 이미지 체크
	Public FrequireEmail'//주문제작 이메일

	public function getMwDiv()
		if (IsNull(Fmaeipdiv) or (Fmaeipdiv="")) then
			getMwDiv = Fmaeipdiv
		else
			getMwDiv = Fmaeipdiv
		end if
	end function

	public function getMwDivName()
		if (Fmaeipdiv = "U") then
		    getMwDivName = "업체"
		elseif (Fmaeipdiv = "W") then
		    getMwDivName = "특정"
		else
		    getMwDivName = "매입"
		end if
	end function

	Private Sub Class_Initialize()
		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function getDesignerDefaultMargin()
		dim sqlStr
		sqlStr = "select top 1 diy_margin from db_academy.dbo.tbl_lec_user "
		sqlStr = sqlStr & " where lecturer_id='" & FRectDesignerID & "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof then
			getDesignerDefaultMargin = rsACADEMYget("diy_margin")
		end if
		rsACADEMYget.close
	end function

	public sub WaitProductDetail(byval itemid)
		dim sqlStr
		sqlStr = "select top 1  IsNULL(i.Cate_large,'') as Cate_large, IsNULL(i.Cate_mid,'') as Cate_mid, IsNULL(i.Cate_small,'') as Cate_small, i.itemdiv, i.itemname,"
		sqlStr = sqlStr & " i.itemid, i.makerid, i.itemcontent,i.designercomment,i.itemsource,i.itemsize,i.itemWeight,"
		sqlStr = sqlStr & " i.sellcash,i.buycash,i.mileage,i.sellyn,"
		sqlStr = sqlStr & " i.deliverytype,i.sourcearea,i.makername,i.limityn,i.limitno,"
		sqlStr = sqlStr & " i.vatYn, i.usinghtml,"
		sqlStr = sqlStr & " i.keywords, i.mwdiv, i.ordercomment, i.optioncnt, i.currstate, "
		sqlStr = sqlStr & " i.rejectmsg, i.rejectDate, i.reRegMsg, i.reRegDate, i.sellEndDate, i.upchemanagecode, "
		sqlStr = sqlStr & " titleimage,mainimage,smallimage,listimage,basicimage,icon1image,icon2image,imgadd "
		sqlStr = sqlStr & " ,i.cstodr,i.requireMakeDay,i.requirecontents,i.refundpolicy,i.infoDiv,i.safetyYn,i.safetyDiv,i.safetyNum,i.freight_min,i.freight_max ,i.requireimgchk , i.requireMakeEmail "
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item i"
		sqlStr = sqlStr & " where i.makerid='" & FRectDesignerID & "'"
		sqlStr = sqlStr & " and i.itemid='" & itemid & "'"

		'response.write sqlStr
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			Flarge      = rsACADEMYget("Cate_large")
			Fmid        = rsACADEMYget("Cate_mid")
			Fsmall      = rsACADEMYget("Cate_small")
			Fitemdiv    = rsACADEMYget("itemdiv")
			FWaitItemID     = rsACADEMYget("itemid")
			FMakerid    = rsACADEMYget("makerid")
			Fitemname   = db2html(rsACADEMYget("itemname"))
			Fitemcontent        = db2html(rsACADEMYget("itemcontent"))
			Fdesignercomment    = db2html(rsACADEMYget("designercomment"))
			Fitemsource     = db2html(rsACADEMYget("itemsource"))
			Fitemsize   =	db2html(db2html(rsACADEMYget("itemsize")))
			FitemWeight =	db2html(db2html(rsACADEMYget("itemWeight")))
			Fsellcash   = db2html(rsACADEMYget("sellcash"))
			Fbuycash    = db2html(rsACADEMYget("buycash"))
			FMileage    = rsACADEMYget("mileage")
			Fsellyn     = rsACADEMYget("sellyn")
			Fdeilverytype = rsACADEMYget("deliverytype")
			Fsourcearea = db2html(rsACADEMYget("sourcearea"))
			Fmakername  = db2html(rsACADEMYget("makername"))
			Flimityn    = rsACADEMYget("limityn")
			Flimitno    = rsACADEMYget("limitno")

			FvatYn		= rsACADEMYget("vatYn")
			Fusinghtml	= rsACADEMYget("usinghtml")
			Fkeywords	= db2html(rsACADEMYget("keywords"))
			Fmwdiv		= rsACADEMYget("mwdiv")
			Fmaeipdiv       = "U"	'DIY상품은 업체배송
			Fordercomment   = db2html(rsACADEMYget("ordercomment"))
            
            FsellEndDate     = rsACADEMYget("sellEndDate")
            Fupchemanagecode = rsACADEMYget("upchemanagecode")
            
			Foptioncnt   = rsACADEMYget("optioncnt")
            Fcurrstate   = rsACADEMYget("currstate")
            Frejectmsg	= rsACADEMYget("rejectmsg")
            FrejectDate	= rsACADEMYget("rejectDate")
            FreRegMsg	= rsACADEMYget("reRegMsg")
            FreRegDate	= rsACADEMYget("reRegDate")

			Fimgtitle = rsACADEMYget("titleimage")
			Fimgmain = rsACADEMYget("mainimage")
			Fimgbasic = rsACADEMYget("basicimage")
			Ficon1 = rsACADEMYget("icon1image")
			Ficon2 = rsACADEMYget("icon2image")
			Fimgsmall = rsACADEMYget("smallimage")
			Fimglist = rsACADEMYget("listimage")
			Fimgadd = rsACADEMYget("imgadd")

			Fcstodr				= rsACADEMYget("cstodr")
			FrequireMakeDay		= rsACADEMYget("requireMakeDay")
			Frequirecontents	= rsACADEMYget("requirecontents")
			Frefundpolicy		= rsACADEMYget("refundpolicy")
			FinfoDiv			= rsACADEMYget("infoDiv")
			FsafetyYn			= rsACADEMYget("safetyYn")
			FsafetyDiv			= rsACADEMYget("safetyDiv")
			FsafetyNum			= rsACADEMYget("safetyNum")
			Ffreight_mine		= rsACADEMYget("freight_min")
			Ffreight_max		= rsACADEMYget("freight_max")
			Frequirechk			= rsACADEMYget("requireimgchk")
			FrequireEmail		= rsACADEMYget("requireMakeEmail")
            
            if (Fsellcash<>0) then
                FMargin     =  100-CLng(Fbuycash/Fsellcash*100)
            end if
		rsACADEMYget.close
	end sub

	public sub WaitProductDetailOption(byval itemid)
		dim sqlStr,i

        sqlStr = " select top 100 o.itemoption, o.optionname as itemoptionname,"
        sqlStr = sqlStr + " isusing, optsellyn, optlimityn, optlimitno, optlimitsold "
        sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_wait_item_option o "
        sqlStr = sqlStr + " where o.itemid = " + CStr(itemid) + " "
        sqlStr = sqlStr + " and o.itemoption<>''"
        sqlStr = sqlStr + " order by itemoption "

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount

		redim preserve FItemList(FResultCount)

			do until rsACADEMYget.Eof
				set FItemList(i) = new CItemOptionItem

				FItemList(i).Fitemoption    = rsACADEMYget("itemoption")
				FItemList(i).Fitemoptionname= db2html(rsACADEMYget("itemoptionname"))
				FItemList(i).Fisusing       = rsACADEMYget("isusing")
				FItemList(i).Foptsellyn     = rsACADEMYget("optsellyn")
				FItemList(i).Foptlimityn    = rsACADEMYget("optlimityn")
				FItemList(i).Foptlimitno    = rsACADEMYget("optlimitno")
				FItemList(i).Foptlimitsold  = rsACADEMYget("optlimitsold")
				FItemList(i).Fcodeview      = db2html(rsACADEMYget("itemoptionname"))

				rsACADEMYget.movenext
				i=i+1
			loop

		rsACADEMYget.Close
	end sub


end Class

class CItemReg

	public largename
	public midname
	public smallname
	public optionbigname
	public optionbigno
	public FRectDesignerID
	public Fitemid
	public FItemoption
	public FMainImage
	public optioncodename


	Private Sub Class_Initialize()


	End Sub

	Private Sub Class_Terminate()

	End Sub


	function CheckFiles(ifile)
		dim file1_size, file1_name
		dim extension

		if (ifile="") then
			CheckFiles =0
			exit function
		end if

		file1_size = ifile.FileLen
	    file1_name = ifile.FileName
	    extension = LCase(Mid(file1_name, InStrRev(file1_name, ".")))

	    if (file1_size>100000) then
	    	response.write "<script language='javascript'>alert('파일사이즈 100,000Byte 까지 지원됩니다.'); history.go(-1);</script>"
	        dbACADEMYget.close()	:	response.End
	    	exit function
	    end if

	    if ((extension <> ".gif") and (extension <> ".jpg") and (extension <> ".png")) then
	    	response.write "<script language='javascript'>alert('이미지(gif,jpg,png) 화일만 지원됩니다.'); history.go(-1);</script>"
	        dbACADEMYget.close()	:	response.End
	    	exit function
	    end if
	    CheckFiles =0
	end function

	public sub SearchCategoryNameLarge(byval largeno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_item_cate_large"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof then
			largename = db2html(rsACADEMYget("code_nm"))
		end if
		rsACADEMYget.close
	end sub

	public sub SearchCategoryNameMid(byval largeno,midno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_item_cate_mid"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"
		sqlStr = sqlStr + " and code_mid='" + Cstr(midno) + "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof then
			midname = db2html(rsACADEMYget("code_nm"))
		end if
		rsACADEMYget.close
	end sub

	public sub SearchCategoryNameSmall(byval largeno,midno,smallno)
		dim sqlStr

		sqlStr = "select code_nm"
		sqlStr = sqlStr + " from db_academy.dbo.tbl_diy_item_cate_small"
		sqlStr = sqlStr + " where code_large='" + Cstr(largeno) + "'"
		sqlStr = sqlStr + " and code_mid='" + Cstr(midno) + "'"
		sqlStr = sqlStr + " and code_small='" + Cstr(smallno) + "'"

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if Not rsACADEMYget.Eof then
			smallname = db2html(rsACADEMYget("code_nm"))
		end if
		rsACADEMYget.close
	end sub

	public sub SearchOptionName(byval bigno)
		dim sqlStr

		sqlStr = "select optioncode01,codename"
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_option_div01"
		sqlStr = sqlStr & " where optioncode01='" & Cstr(bigno) & "'"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			optionbigname = rsACADEMYget("codename")
			optionbigno = rsACADEMYget("optioncode01")
		rsACADEMYget.close
	end sub

	function FormatStr(n,orgData)
			dim tmp
			if (n-Len(CStr(orgData))) < 0 then
				FormatStr = CStr(orgData)
				Exit Function
			end if

			tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
			FormatStr = tmp
	end Function


end Class




%>