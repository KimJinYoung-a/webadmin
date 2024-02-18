<%

class CWaitItemlist2014
	public FListType
	public Fcurrstate
	public FSort

	public FTotCnt
	public FSPageNo
	public FEPageNo
	public FPageSize
	public FCurrPage

	public Fcatecode
	public Fmakerid
	public Fitemname
 	public Fitemid

	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectctrState

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FTotCnt =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// 승인대기상품리스트
	public Function fnGetSummaryList
		dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_getSummrayList]('"&FListType&"', '"&FcurrState&"','"&FSort&"','"&Fmakerid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetSummaryList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//승인대기 상품 상세리스트
	' /admin/itemmaster/item_confirm.asp
	public Function fnGetWaitItemList
		Dim strSql

		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_getItemListCnt] '"&Fcatecode&"','"&Fmakerid&"','"&Fitemname&"','"&Fcurrstate&"','"&FItemid&"', '" + CStr(FRectCate_Large) + "', '" + CStr(FRectCate_Mid) + "', '" + CStr(FRectCate_Small) + "','"&FRectctrState&"'"

		'response.write strSql & "<Br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_temp].[dbo].sp_Ten_wait_item_getItemList '"&Fcatecode&"','"&Fmakerid&"','"&Fitemname&"','"&Fcurrstate&"','"&FSort&"','"&FItemid&"',"&FSPageNo&","&FEPageNo&", '" + CStr(FRectCate_Large) + "', '" + CStr(FRectCate_Mid) + "', '" + CStr(FRectCate_Small) + "','"&FRectctrState&"'"

		'response.write strSql & "<Br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'//승인 진행일자 로그
	public Function fnGetWaitItemLog
		Dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_log_getItemList]("&Fitemid&")"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemLog = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//API 진행일자 로그
	public Function fnGetWaitItemApiLog
		Dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_log_getItemList]("&Fitemid&", 'Y')"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemApiLog = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetOldWaitItemLog
		Dim strSql
		strSql =" select  rejectdate, rejectmsg, reregdate, reregmsg, currstate From db_temp.dbo.tbl_wait_item where itemid =" &Fitemid&" and currstate in (2, 0, 5 ) "
		rsget.Open strSql,dbget,1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOldWaitItemLog = rsget.getRows()
		END IF
		rsget.close
	End Function
end Class


'//진행상태 함수
Sub sbOptItemWaitStatus(currstate)
	%>
	<option value="A" <%IF currstate="A" THEN%>selected<%END IF%>>전체</option>
	<option value="1" <%IF currstate="1" THEN%>selected<%END IF%>>승인대기</option>
	<option value="5" <%IF currstate="5" THEN%>selected<%END IF%>>승인대기 (재등록)</option>
	<option value="2" <%IF currstate="2" THEN%>selected<%END IF%>>승인보류 (재등록요청)</option>
	<option value="0" <%IF currstate="0" THEN%>selected<%END IF%>>승인반려 (재등록불가)</option>
	<option value="7" <%IF currstate="7" THEN%>selected<%END IF%>>승인완료</option>
	<%
End Sub

	function GetCurrStateColor(ByVal FCurrState)
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="3" then
			GetCurrStateColor = "#DD0000"
		elseif FCurrState="4" then
			GetCurrStateColor = "#DD0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

 function GetCurrStateName(ByVal FCurrState)
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "승인대기"
		elseif FCurrState="2" then
			GetCurrStateName = "승인보류<Br>(재등록요청)"
		elseif FCurrState="3" then
			GetCurrStateName = "처리대기<Br>(재등록요청)"
		elseif FCurrState="4" then
			GetCurrStateName = "처리실패<Br>(재등록요청)"
		elseif FCurrState="7" then
			GetCurrStateName = "승인완료"
		elseif FCurrState="5" then
			GetCurrStateName = "승인대기<Br>(재등록)"
		elseif FCurrState="0" then
			GetCurrStateName = "승인반려<Br>(재등록불가)"
		elseif FCurrState="9" then
			GetCurrStateName = "업체삭제"
		else
			GetCurrStateName = ""
		end if
	end function

	 function GetCurrStateContsName(ByVal FCurrState)
		GetCurrStateContsName = ""
		if FCurrState="1" then
			GetCurrStateContsName = "승인대기"
		elseif FCurrState="2" then
			GetCurrStateContsName = "승인보류(재등록요청)"
		elseif FCurrState="7" then
			GetCurrStateContsName = "승인완료"
		elseif FCurrState="5" then
			GetCurrStateContsName = "승인대기(재등록)"
		elseif FCurrState="0" then
			GetCurrStateContsName = "승인반려(재등록불가)"
		elseif FCurrState="9" then
			GetCurrStateContsName = "업체삭제"
		else
			GetCurrStateContsName = ""
		end if
	end function

	function fnGetCurrStateShortName(ByVal FCurrState)
			fnGetCurrStateShortName = ""
		if FCurrState="1" then
			fnGetCurrStateShortName = "등록"
		elseif FCurrState="2" then
			fnGetCurrStateShortName = "보류"
		elseif FCurrState="7" then
			fnGetCurrStateShortName = "완료"
		elseif FCurrState="5" then
			fnGetCurrStateShortName = "재등록"
		elseif FCurrState="0" then
			fnGetCurrStateShortName = "반려"
		elseif FCurrState="9" then
			fnGetCurrStateShortName = "삭제"
		elseif FCurrState="S" then
			fnGetCurrStateShortName = "안전정보처리"
		elseif FCurrState="I" then
			fnGetCurrStateShortName = "이미지처리"
		else
			fnGetCurrStateShortName = ""
		end if
	End Function
%>
