<%
'###########################################################
' Description : 텐바이텐 대량구매 사이트
' Hieditor : 2013.07.15 한용민 생성
'###########################################################

class CwholesaleLoginItem
	public Fuserid
	public Fuserpass
	public Fshopname
	public Fshopphone
	public Fshopzipcode
	public Fshopaddr1
	public Fshopaddr2
	public Fmanname
	public Fmanhp
	public Fmanphone
	public Fmanemail
	public Fisusing
	public FShopdiv
	public Fstockbasedate
	public Fshopsocname
	public Fshopsocno
	public Fshopceoname
	public Fvieworder
	public Fgroupid
	public fcurrencyUnit
	public fexchangeRate
	public fbasedate
	public fcurrencyChar
	public fregdate
	public freguserid
	public flastuserid
	public fidx
	public fmultipleRate
    public fpyeong
	public FshopCountryCode
    public FcountryNamekr
	public fshopid
    public FdecimalPointLen
    public FdecimalPointCut
	public fothershopid
	public fsiteseq	
	public flastupdate
	public flastadminuserid	
    public Fismobileusing
    public Fmobileshopname
    public Fmobileshopimage
    public Fmobileworkhour
    public Fmobileclosedate
    public Fmobiletel
    public Fmobileaddr
    public Fmobilemapimage
    public Fmobilebysubway
    public Fmobilebybus
    public Fmobilelatitude
    public Fmobilelongitude
	public fadmindisplang
	public fsitename
	public fcountrylangcd
	
	public FResult
end Class

class CwholesaleLogin
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount
	public FRectUserID
	public FRectUserPW
	
	'### 로그인프로세스
	public Sub GetLoginData()
		dim sqlStr , i , sqlsearch

		sqlStr = "EXECUTE [db_shop].[dbo].[sp_Ten_UserLoginProc_wholesale] '" & FRectUserID & "', '" & FRectUserPW & "'"
		
		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr,dbget,1
		
		If not rsget.EOF Then
			set FOneItem = new CwholesaleLoginItem
			FOneItem.FResult = rsget("result")
			
			If FOneItem.FResult = "0" Then
				FOneItem.FUserID 		= rsget("userid")
				FOneItem.FShopName		= db2html(rsget("shopname"))
				FOneItem.FcurrencyUnit	= rsget("currencyUnit")
				FOneItem.FShopdiv		= rsget("shopdiv")
				FOneItem.Fgroupid		= rsget("groupid")
				FOneItem.FcountryNamekr	= db2html(rsget("countryNamekr"))
				FOneItem.Fismobileusing	= rsget("ismobileusing")
				FOneItem.FcurrencyChar	= rsget("currencyChar")
				FOneItem.fcountrylangcd	= rsget("countrylangcd")
				FOneItem.Fmanemail		= rsget("manemail")
			End If
		End If
		rsget.close
	end Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage = 1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>