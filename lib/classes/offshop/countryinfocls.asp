<%

' EMS �������� ������
Class clsCountryInfoItem
	Public FcountryCode	' �����ڵ�
	Public FcountryNameKr' ������(�ѱ�)
	Public FcountryNameEn' ������(����)
	Public Fisusing		' ��뿩��

	' �ʱ�ȭ
    Private Sub Class_initialize()
		FcountryCode	= ""
		FcountryNameKr  = ""
		FcountryNameEn  = ""
		Fisusing		= "Y"
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub

End Class

' EMS �������� Ŭ����
Class CCountryInfo

    public FOneItem
    public FItemList()

	'// �˻�����
	public FRectCurrPage
	public FRectPageSize
	public FRectCountryCode
	public FRectisUsing

	' ����¡
	Dim FTotalCount
	Dim FTotalPage
	Dim FResultCount

	' ����Ʈ
	Public Function GetCountryInfoList

		Dim i, strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
		,Array("@PageSize"		, adInteger	, adParamInput	,		, FRectPageSize)	_
		,Array("@CurrPage"		, adInteger	, adParamInput	,		, FRectCurrPage) _
		,Array("@TotalCount"	, adBigInt	, adParamOutput	,		, 0) _
		,Array("@countryCode"	, adChar	, adParamInput	, 2	, FRectCountryCode) _
		,Array("@isUsing"	, adChar	, adParamInput	, 1	, FRectisUsing) _
		)

		strSql = "[db_shop].[dbo].[sp_Ten_shop_get_country_info_LIST]"

		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		FTotalCount = GetValue(paramInfo, "@TotalCount")							' Output ����
		FTotalCount = CInt(FTotalCount)

		FTotalPage = Int((FTotalCount-1) / FRectPageSize) + 1
		FResultCount = FRectPageSize
		If FTotalCount = 0 Or FTotalPage < FRectCurrPage Then
			FResultCount = 0
		ElseIf FTotalPage = FRectCurrPage Then	' ������ �������̸�
			FResultCount = FTotalCount Mod FRectPageSize
			If FResultCount = 0 Then			' ������ �������� ������������� ����
				FResultCount = FRectPageSize
			End If
		End If
		ReDim FItemList(FResultCount)

		If Not rsget.EOF Then
			i = 0
			Do Until rsget.EOF

				Set FItemList(i) = new clsCountryInfoItem

				FItemList(i).FcountryCode		= null2blank(rsget("countryCode"))
				FItemList(i).FcountryNameKr		= null2blank(rsget("countryNameKr"))
				FItemList(i).FcountryNameEn		= null2blank(rsget("countryNameEn"))
				FItemList(i).Fisusing			= null2blank(rsget("isusing"))

				i = i + 1
				rsget.MoveNext
			Loop
		End If

		rsget.close()

	End Function

End Class
%>
