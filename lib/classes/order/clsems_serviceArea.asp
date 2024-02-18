<%
'==========================================================================
'	Description: EMS 서비스지역 클래스, 서동석
'	History: 2009.04.07
'==========================================================================
' EMS 중량/지역별 요금

Class clsEms_weightPriceItem
    public FcompanyCode
    public FemsAreaCode
    public FWeightLimit
    public FemsPrice


    ' 초기화
    Private Sub Class_initialize()
		FemsAreaCode	= ""
		FWeightLimit    = 0
		FemsPrice       = 0

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class


' EMS 서비스지역 아이템
Class clsEms_serviceAreaItem
    public FcompanyCode
	Public FcountryCode	' 국가코드
	Public FcountryNameKr' 국가명(한글)
	Public FcountryNameEn' 국가명(영문)
	Public FemsAreaCode	' EMS요금적용지역
	Public FemsMaxWeight' EMS최대중량
	Public FreceiverPay	' 수취인부담여부
	Public Fisusing		' 사용여부
	Public FetcContents	' 기타사항



	' 초기화
    Private Sub Class_initialize()
		FcountryCode	= ""
		FcountryNameKr  = ""
		FcountryNameEn  = ""
		FemsAreaCode	= ""
		FemsMaxWeight   = 0
		FreceiverPay	= "N"
		Fisusing		= "Y"
		FetcContents	= ""



	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

' EMS 서비스지역 클래스
Class CEms

    public FOneItem
    public FItemList()



	'// 검색조건
	public FRectCurrPage
	public FRectPageSize
	public FRectCountryCode
	public FRectisUsing
	public FRectCountryNameKr
	public FRectCountryNameEn
	public FRectEmsAreaCode
    public FRectCompanyCode

	public FRectWeightLimit
	public FRectWeight

	' 페이징
	Dim FTotalCount
	Dim FTotalPage
	Dim FResultCount

	public function GetWeightPriceListByWeight
	    Dim i, strSql
		Dim paramInfo
		    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
				,Array("@weight"		, adInteger	, adParamInput	,		, FRectWeight)	_
            	,Array("@companyCode"	, adVarchar	, adParamInput	, 3		, FRectCompanyCode) _
			)

		strSql = "db_order.dbo.sp_Ten_Ems_priceListByWeight"
		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		FTotalCount = rsget.RecordCount
		FResultCount= FTotalCount

		ReDim FItemList(FResultCount)
	    If Not rsget.EOF Then
			i = 0
			Do Until rsget.EOF

				Set FItemList(i) = new clsEms_weightPriceItem

                FItemList(i).FcompanyCode  = null2blank(rsget("companyCode"))
                FItemList(i).FemsAreaCode  = null2blank(rsget("emsAreaCode"))
                FItemList(i).FWeightLimit  = null2blank(rsget("WeightLimit"))
                FItemList(i).FemsPrice     = null2blank(rsget("emsPrice"))

				i = i + 1
				rsget.MoveNext
			Loop
		End If

		rsget.close()

	end function

	Public Function GetWeightPriceList
	    Dim i, strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FRectPageSize)	_
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FRectCurrPage) _
			,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
			,Array("@emsAreaCode"	, adVarChar	, adParamInput	, 2	    , FRectEmsAreaCode) _
			,Array("@WeightLimit"	, adInteger	, adParamInput	, 	    , FRectWeightLimit) _
            ,Array("@companyCode"	, adVarchar	, adParamInput	, 3		, FRectCompanyCode) _
		)

		strSql = "db_order.dbo.sp_Ten_Ems_weightPrice_GetList"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		FTotalCount = GetValue(paramInfo, "@TotalCount")							' Output 리턴
		FTotalCount = CInt(FTotalCount)

		FTotalPage = Int((FTotalCount-1) / FRectPageSize) + 1
		FResultCount = FRectPageSize
		If FTotalCount = 0 Or FTotalPage < FRectCurrPage Then
			FResultCount = 0
		ElseIf FTotalPage = FRectCurrPage Then	' 마지막 페이지이면
			FResultCount = FTotalCount Mod FRectPageSize
			If FResultCount = 0 Then			' 나누어 떨어지면 페이지사이즈와 같음
				FResultCount = FRectPageSize
			End If
		End If
		ReDim FItemList(FResultCount)

		If Not rsget.EOF Then
			i = 0
			Do Until rsget.EOF

				Set FItemList(i) = new clsEms_weightPriceItem

                FItemList(i).FcompanyCode  = null2blank(rsget("companyCode"))
                FItemList(i).FemsAreaCode  = null2blank(rsget("emsAreaCode"))
                FItemList(i).FWeightLimit  = null2blank(rsget("WeightLimit"))
                FItemList(i).FemsPrice     = null2blank(rsget("emsPrice"))

				i = i + 1
				rsget.MoveNext
			Loop
		End If

		rsget.close()

    end Function

	' 리스트
	Public Function GetServiceAreaList

		Dim i, strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FRectPageSize)	_
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FRectCurrPage) _
			,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
			,Array("@countryCode"	, adChar	, adParamInput	, 2		, FRectCountryCode) _
			,Array("@isUsing"		, adChar	, adParamInput	, 1		, FRectisUsing) _
			,Array("@CountryNameKr"	, adVarchar	, adParamInput	, 50	, FRectCountryNameKr) _
			,Array("@CountryNameEn"	, adVarchar	, adParamInput	, 50	, FRectCountryNameEn) _
			,Array("@EmsAreaCode"	, adVarchar	, adParamInput	, 2		, FRectEmsAreaCode) _
            ,Array("@companyCode"	, adVarchar	, adParamInput	, 3		, FRectCompanyCode) _
		)

		strSql = "db_order.dbo.sp_Ten_Ems_serviceArea_GetList"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		FTotalCount = GetValue(paramInfo, "@TotalCount")							' Output 리턴
		FTotalCount = CInt(FTotalCount)

		FTotalPage = Int((FTotalCount-1) / FRectPageSize) + 1
		FResultCount = FRectPageSize
		If FTotalCount = 0 Or FTotalPage < FRectCurrPage Then
			FResultCount = 0
		ElseIf FTotalPage = FRectCurrPage Then	' 마지막 페이지이면
			FResultCount = FTotalCount Mod FRectPageSize
			If FResultCount = 0 Then			' 나누어 떨어지면 페이지사이즈와 같음
				FResultCount = FRectPageSize
			End If
		End If
		ReDim FItemList(FResultCount)

		If Not rsget.EOF Then
			i = 0
			Do Until rsget.EOF

				Set FItemList(i) = new clsEms_serviceAreaItem

                FItemList(i).FcompanyCode		= null2blank(rsget("companyCode"))
				FItemList(i).FcountryCode		= null2blank(rsget("countryCode"))
				FItemList(i).FcountryNameKr		= null2blank(rsget("countryNameKr"))
				FItemList(i).FcountryNameEn		= null2blank(rsget("countryNameEn"))
				FItemList(i).FemsAreaCode		= null2blank(rsget("emsAreaCode"))
				FItemList(i).FemsMaxWeight		= null2blank(rsget("emsMaxWeight"))
				FItemList(i).FreceiverPay		= null2blank(rsget("receiverPay"))
				FItemList(i).Fisusing			= null2blank(rsget("isusing"))
				FItemList(i).FetcContents		= null2blank(rsget("etcContents"))



				i = i + 1
				rsget.MoveNext
			Loop
		End If

		rsget.close()

	End Function


    ' 데이터
	Public Function GetWeightPriceData()
		Set FOneItem = new clsEms_weightPriceItem

		If FRectEmsAreaCode <> "" and FRectWeightLimit<>"" Then
			Dim i, strSql
			Dim paramInfo
			paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
				,Array("@EmsAreaCode"			, adVarChar	, adParamInput	,2, FRectEmsAreaCode)	_
				,Array("@WeightLimit"			, adInteger	, adParamInput	,, FRectWeightLimit)	_
                ,Array("@companyCode"			, adVarchar	, adParamInput	, 3		, FRectCompanyCode) _
			)

			strSql = "db_order.dbo.sp_Ten_Ems_weightPrice_GetData"
			call fnExecSPReturnRSOutput(strSql, paramInfo)


			If Not rsget.EOF Then

                FOneItem.FcompanyCode		= null2blank(rsget("companyCode"))
				FOneItem.FEmsAreaCode		= null2blank(rsget("EmsAreaCode"))
				FOneItem.FWeightLimit		= null2blank(rsget("WeightLimit"))
				FOneItem.FemsPrice			= null2blank(rsget("emsPrice"))

			End If

			rsget.close()

		End If

    End Function


	' 데이터
	Public Function GetServiceAreaData()
		Set FOneItem = new clsEms_serviceAreaItem

		If FRectCountryCode <> "" Then
			Dim i, strSql
			Dim paramInfo
			paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			    ,Array("@PKID"			, adChar	, adParamInput	,2, FRectCountryCode)	_
                ,Array("@companyCode"	, adChar	, adParamInput	,3, FRectCompanyCode)	_
			)

			strSql = "db_order.dbo.sp_Ten_Ems_serviceArea_GetData"
			call fnExecSPReturnRSOutput(strSql, paramInfo)


			If Not rsget.EOF Then

                FOneItem.FCompanyCode		= null2blank(rsget("CompanyCode"))
				FOneItem.FcountryCode		= null2blank(rsget("countryCode"))
				FOneItem.FcountryNameKr		= null2blank(rsget("countryNameKr"))
				FOneItem.FcountryNameEn		= null2blank(rsget("countryNameEn"))
				FOneItem.FemsAreaCode		= null2blank(rsget("emsAreaCode"))
				FOneItem.FemsMaxWeight		= null2blank(rsget("emsMaxWeight"))
				FOneItem.FreceiverPay		= null2blank(rsget("receiverPay"))
				FOneItem.Fisusing			= null2blank(rsget("isusing"))
				FOneItem.FetcContents		= null2blank(rsget("etcContents"))

			End If

			rsget.close()

		End If

	End Function

    Public Function ProcWeightPrice(ByVal mode)

		Dim ErrCode, ErrMsg

		Dim strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@mode"			, adVarchar	, adParamInput	, 10 , mode)	_
			,Array("@emsAreaCode"		, adVarChar	, adParamInput	, 2 , FOneItem.FemsAreaCode)	_
			,Array("@weightLimit"		, adInteger	, adParamInput	,   , FOneItem.FweightLimit)	_
			,Array("@emsPrice"		, adCurrency	, adParamInput	,   , FOneItem.FemsPrice)	_
            ,Array("@companyCode"	, adChar		, adParamInput	,3	, FOneItem.FCompanyCode)	_
		)


		strSql = "db_order.dbo.sp_Ten_Ems_weightPrice_Proc"
'rw strSql
		Call fnExecSP(strSql, paramInfo)

		ErrCode  = GetValue(paramInfo, "@RETURN_VALUE")	  ' 에러코드
		ErrCode  = CInt(ErrCode)
'rw ErrCode
		If ErrCode <> 0 Then
			ProcWeightPrice = False
			sbAlertMessage "오류발생", "", "back"
		Else
			ProcWeightPrice = True
		End If

    End Function

    Public Function ProcServiceArea(ByVal mode)

		Dim ErrCode, ErrMsg

		Dim strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@mode"			, adVarchar	, adParamInput	, 10 , mode)	_
			,Array("@countryCode"		, adChar	, adParamInput	, 2  , FOneItem.FcountryCode)	_
			,Array("@countryNameKr"		, adVarchar	, adParamInput	, 50 , FOneItem.FcountryNameKr)	_
			,Array("@countryNameEn"		, adVarchar	, adParamInput	, 50 , FOneItem.FcountryNameEn)	_
			,Array("@emsAreaCode"		, adVarChar	, adParamInput	, 2  , FOneItem.FemsAreaCode)	_
			,Array("@emsMaxWeight"		, adInteger	, adParamInput	,    , FOneItem.FemsMaxWeight)	_
			,Array("@receiverPay"		, adChar	, adParamInput	, 1  , FOneItem.FreceiverPay)	_
			,Array("@isusing"		    , adChar	, adParamInput	, 1  , FOneItem.Fisusing)	_
			,Array("@etcContents"	    , adVarchar	, adParamInput	,500 , FOneItem.FetcContents)	_
            ,Array("@companyCode"	    , adVarchar	, adParamInput	,3   , FOneItem.FcompanyCode)	_
		)

		strSql = "db_order.dbo.sp_Ten_Ems_serviceArea_Proc"
'rw strSql
'response.End
		Call fnExecSP(strSql, paramInfo)

		ErrCode  = GetValue(paramInfo, "@RETURN_VALUE")	  ' 에러코드
		ErrCode  = CInt(ErrCode)
'rw ErrCode
		If ErrCode <> 0 Then
			ProcServiceArea = False
			sbAlertMessage "오류발생", "", "back"
		Else
			ProcServiceArea = True
		End If

	End Function

End Class
%>
