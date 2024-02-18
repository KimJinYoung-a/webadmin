<%

Class CExtJungsanProfitSiteItem
    public Fsellsite
    public Fmakerid
    public Fextitemno
    public FextTenMeachulPrice
    public FextTenJungsanPrice
    public Ftenbuycash
    public Fjungsangain
    public FU_extTenJungsanPrice
    public FU_buycash
    public FU_jungsangain
    public FW_extTenJungsanPrice
    public FW_buycash
    public FW_jungsangain
    public FM_extTenJungsanPrice
    public FM_buycash
    public FM_jungsangain
    public FN_extTenJungsanPrice
    public FN_buycash
    public FN_jungsangain

    public FdefaultFreeBeasongLimit
    public FdefaultDeliverPay
    public FdefaultDeliveryType

    public function getDlvTypeHtml()
        dim ret, ret2
        if (FdefaultDeliveryType="7") then
            ret = "업체착불배송"
        elseif (FdefaultDeliveryType="9") then
            ret = "업체조건배송"
        end if

        if FdefaultFreeBeasongLimit<>0 then
            ret2 = FormatNumber(FdefaultFreeBeasongLimit,0)&"원<br>"
            ret2 = ret2 & FormatNumber(FdefaultDeliverPay,0)&"원"
        end if

        getDlvTypeHtml = ret &CHKIIF(ret<>"","<br>","")& ret2
    end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

End Class


Class CExtJungsanProfitSiteDtlItem
    public Fsellsite
    public Fmakerid
    public Fextitemno
    public FextTenMeachulPrice
    public FextTenJungsanPrice
    public Ftenbuycash
    public Fjungsangain

    public Fitemid
    public Fomwdiv
    public Fitemname
    public Fsmallimage
    public Fsellcash
    public Fbuycash
    public Fsellyn
    public Flimityn
    public Flimitno
    public Flimitsold

    Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function

    public function getItemStatHtml()
        dim ret 
        
        if IsSoldOut then ret = "품절"

        if FLimitYn="Y" then ret = CHKIIF(ret<>"","<br>","")&"<font color=blue'>한정("&GetLimitEa&")</font>"

    end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

End Class

Class CExtJungsanProfit
    public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	public FRectSellSite
    public FRectMakerid
    public FRectItemid
    public FRectStartDate
	public FRectEndDate

    public FRectReturnExcept
    public FRectMinusGainOnly

    public Sub GetExtUpcheDlvProfit()
        Dim sqlStr, i

        sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_UpcheDlvGainSum] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "', " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ""

        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsSTSget.RecordCount
		FTotalCount = FResultCount

        redim preserve FItemList(FResultCount)
		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.eof
				set FItemList(i) = new CExtJungsanProfitSiteItem
                FItemList(i).Fsellsite  = rsSTSget("sellsite")
                FItemList(i).Fextitemno = rsSTSget("extitemno")
                FItemList(i).FextTenMeachulPrice = rsSTSget("extTenMeachulPrice")
                ''FItemList(i).FextTenJungsanPrice = rsSTSget("extTenJungsanPrice")
                FItemList(i).Ftenbuycash  = rsSTSget("tenbuycash")
                FItemList(i).Fjungsangain = rsSTSget("jungsangain")

                FItemList(i).Fmakerid  = rsSTSget("makerid")

                FItemList(i).FdefaultFreeBeasongLimit   = rsSTSget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsSTSget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsSTSget("defaultDeliveryType")
				i=i+1
				rsSTSget.moveNext
			loop
		end if
		rsSTSget.Close
    end sub


    public Sub GetExtJungsanProfit()
        Dim sqlStr, i
        '' @styyyymmdd ,@edyyyymmdd ,@sellsite varchar(32) = NULL,@makerid varchar(32) = NULL
		sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSum] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "', " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ""

        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsSTSget.RecordCount
		FTotalCount = FResultCount
	
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.eof
				set FItemList(i) = new CExtJungsanProfitSiteItem

				FItemList(i).Fsellsite  = rsSTSget("sellsite")
                FItemList(i).Fextitemno = rsSTSget("extitemno")
                FItemList(i).FextTenMeachulPrice = rsSTSget("extTenMeachulPrice")
                FItemList(i).FextTenJungsanPrice = rsSTSget("extTenJungsanPrice")
                FItemList(i).Ftenbuycash  = rsSTSget("tenbuycash")
                FItemList(i).Fjungsangain = rsSTSget("jungsangain")

                
                FItemList(i).FU_extTenJungsanPrice = rsSTSget("U_extTenJungsanPrice")
                FItemList(i).FU_buycash = rsSTSget("U_buycash")
                FItemList(i).FU_jungsangain = rsSTSget("U_jungsangain")
                FItemList(i).FW_extTenJungsanPrice = rsSTSget("W_extTenJungsanPrice")
                FItemList(i).FW_buycash = rsSTSget("W_buycash")
                FItemList(i).FW_jungsangain = rsSTSget("W_jungsangain")
                FItemList(i).FM_extTenJungsanPrice = rsSTSget("M_extTenJungsanPrice")
                FItemList(i).FM_buycash = rsSTSget("M_buycash")
                FItemList(i).FM_jungsangain = rsSTSget("M_jungsangain")
                FItemList(i).FN_extTenJungsanPrice = rsSTSget("N_extTenJungsanPrice")
                FItemList(i).FN_buycash = rsSTSget("N_buycash")
                FItemList(i).FN_jungsangain = rsSTSget("N_jungsangain")
                

                'if (FRectSellSite<>"") then
                    FItemList(i).Fmakerid  = rsSTSget("makerid")
                'end if
				i=i+1
				rsSTSget.moveNext
			loop
		end if
		rsSTSget.Close
    end Sub

    public Sub GetExtJungsanProfitDetail()
        Dim sqlStr, i
        '' @styyyymmdd ,@edyyyymmdd ,@sellsite varchar(32) = NULL,@makerid varchar(32) = NULL
		sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSumDtl] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "'," & CHKIIF(FRectItemid="","NULL",FRectItemid) & ", " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ""
''rw sqlStr
        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsSTSget.RecordCount
		FTotalCount = FResultCount
	
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.eof
				set FItemList(i) = new CExtJungsanProfitSiteDtlItem

				FItemList(i).Fsellsite  = rsSTSget("sellsite")
                FItemList(i).Fmakerid  = rsSTSget("makerid")
                FItemList(i).Fitemid  = rsSTSget("itemid")
                FItemList(i).Fextitemno = rsSTSget("extitemno")
                FItemList(i).FextTenMeachulPrice = rsSTSget("extTenMeachulPrice")
                FItemList(i).FextTenJungsanPrice = rsSTSget("extTenJungsanPrice")
                FItemList(i).Ftenbuycash  = rsSTSget("tenbuycash")
                FItemList(i).Fjungsangain = rsSTSget("jungsangain")

                FItemList(i).Fomwdiv        = rsSTSget("omwdiv")
                FItemList(i).Fitemname      = rsSTSget("itemname")
                FItemList(i).Fsmallimage    = rsSTSget("smallimage")
                FItemList(i).Fsellcash      = rsSTSget("sellcash")
                FItemList(i).Fbuycash       = rsSTSget("buycash")
                FItemList(i).Fsellyn        = rsSTSget("sellyn")
                FItemList(i).Flimityn       = rsSTSget("limityn")
                FItemList(i).Flimitno       = rsSTSget("limitno")
                FItemList(i).Flimitsold     = rsSTSget("limitsold")
                
                If Not(FItemList(i).FsmallImage = "" OR isNull(FItemList(i).FsmallImage)) Then
					FItemList(i).FsmallImage = "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).Fitemid) & "/" & FItemList(i).FsmallImage
				Else
					FItemList(i).FsmallImage = "http://fiximage.10x10.co.kr/images/spacer.gif"
				End If

				i=i+1
				rsSTSget.moveNext
			loop
		end if
		rsSTSget.Close
    end Sub

    

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class
%>