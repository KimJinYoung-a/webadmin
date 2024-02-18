<%

Class ClsDispCateItem
	public FdispCateCode
	public FdispCateName
	public Fcount

	Private Sub Class_Initialize()
		'
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
End Class

Class ClsDispCateArray
	Public FItemList()

	Private Sub Class_Initialize()
		dim i
		dim dispCateCodeArr, dispCateNameArr

		dispCateCodeArr = "101,102,103,104,114,106,112,113,115,110,111"
		dispCateNameArr = "디자인문구,디지털/핸드폰,캠핑/트래블,토이,가구,홈인테리어,키친/푸드,패션/뷰티,베이비/키즈,Cat & Dog,BOOK"

		dispCateCodeArr = Split(dispCateCodeArr, ",")
		dispCateNameArr = Split(dispCateNameArr, ",")

		redim FItemList(UBound(dispCateNameArr))

		for i = 0 to UBound(dispCateCodeArr) - 1
			set FItemList(i) = new ClsDispCateItem

			FItemList(i).FdispCateCode = dispCateCodeArr(i)
			FItemList(i).FdispCateName = dispCateNameArr(i)
		next
	End Sub

	Private Sub Class_Terminate()
		'
	End Sub
End Class

'// 업데이트 필요한지
Class ClsIsUpdateNeedItem
	public FwillFinishEvent
	public FEventCount
	public FupcheRequest
	public FitemRequest
	public FItemSellRequest
	public FBrandRequest
	public FEventPrize

	Private Sub Class_Initialize()
		FwillFinishEvent = False
		FEventCount = False
		FupcheRequest = False
		FitemRequest = False
		FItemSellRequest = False
		FBrandRequest = False
		FEventPrize = False
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

'// 종료임박 이벤트
Class ClsWillFinishEventItem
	public FNormalCnt
	public FAppCnt

	public FNormal101Cnt
	public FNormal102Cnt
	public FNormal103Cnt
	public FNormal104Cnt
	public FNormal114Cnt
	public FNormal106Cnt
	public FNormal112Cnt
	public FNormal113Cnt
	public FNormal115Cnt
	public FNormal110Cnt
	public FNormal111Cnt

	public FApp101Cnt
	public FApp102Cnt
	public FApp103Cnt
	public FApp104Cnt
	public FApp114Cnt
	public FApp106Cnt
	public FApp112Cnt
	public FApp113Cnt
	public FApp115Cnt
	public FApp110Cnt
	public FApp111Cnt

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class ClsEventCountItem
	public FtotCount
	public Fstate0
	public Fstate1
	public Fstate2
	public Fstate3
	public Fstate5
	public Fstate7
	public Fstate6
	public Fstate9

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class ClsCompanyRequestItem
	public FdispCate
	public FCateName
	public Fcount

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class ClsCompanyContractItem
	public FsendUserID
	public Fusername
	public Fcount

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class ClsCompanyInfoModifyReqItem
	public FuserID
	public Fusername
	public Fcount

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class ClsItemRegRequestCountItem
	public FcateCode
	public FcateName
	public Fcount1
	public Fcount5

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CEventPrizeItem
	public FeventCode
	public FeventName
	public FeventKind
	public FuserID
	public FuserName
	public FdDay
	public FprizeDay

	public function GetDDayStr()
		if (FdDay = 0) then
			GetDDayStr = "<font color='red'>D-DAY</font>"
		elseif (FdDay < 0) then
			GetDDayStr = "D" & FdDay
		else
			GetDDayStr = "D+" & FdDay
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


'==============================================================================
'// 함수들
function GetWeekDayName(dt)
	dim s, ret
	s = DatePart("w", CDate(dt))

	ret = s
	Select Case s
		Case "1"
			ret = "<font color='red'>일</font>"
		Case "2"
			ret = "월"
		Case "3"
			ret = "화"
		Case "4"
			ret = "수"
		Case "5"
			ret = "목"
		Case "6"
			ret = "금"
		Case "7"
			ret = "<font color='blue'>토</font>"
		Case Else
			''
	End Select

	GetWeekDayName = ret
end function

Function GetIndexFromDispCateCode(dispCateCode)
	dim i
	dim cateArr

	GetIndexFromDispCateCode = -1

	set cateArr = new ClsDispCateArray

	for i = 0 to UBound(cateArr.FItemList) - 1
		if (cateArr.FItemList(i).FdispCateCode = dispCateCode) then
			GetIndexFromDispCateCode = i
			exit for
		end if
	next
End Function

Function GetNameFromDispCateCode(dispCateCode)
	dim i
	dim cateArr

	GetNameFromDispCateCode = "ERR"

	set cateArr = new ClsDispCateArray

	for i = 0 to UBound(cateArr.FItemList) - 1
		if (cateArr.FItemList(i).FdispCateCode = dispCateCode) then
			GetNameFromDispCateCode = cateArr.FItemList(i).FdispCateName
			exit for
		end if
	next
End Function

%>
