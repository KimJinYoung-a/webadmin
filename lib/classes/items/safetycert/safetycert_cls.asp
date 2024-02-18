<%
'###########################################################
' Description : 안전인증품목관리 클래스
' History : 2018.01.16 한용민 생성
'###########################################################

Class Csafetycert_oneitem
	public finfoDiv
	public finfoDivName
	public finfoValidCnt
	public fSafetyTargetYN
	public fSafetyCertYN
	public fSafetyConfirmYN
	public fSafetySupplyYN
	public fSafetyComply
	public fregdate
	public flastupdate
	public flastadminid

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Csafetycert
	Public FItemList()
	public foneitem
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FPageCount

	'//admin/itemmaster/safetycert/safetycert.asp
	public Function fsafetycert()
		dim sqlStr, i

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " infoDiv, infoDivName, infoValidCnt, SafetyTargetYN, SafetyCertYN, SafetyConfirmYN, SafetySupplyYN, SafetyComply"
		sqlStr = sqlStr & " , regdate, lastupdate, lastadminid"
		sqlStr = sqlStr & " from db_item.dbo.tbl_item_infoDiv"
		sqlStr = sqlStr & " order by infoDiv asc"

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		i=0		
		if  not rsget.EOF  then
			redim preserve FItemList(FResultCount)

			do until rsget.eof
				set FItemList(i) = new Csafetycert_oneitem
					FItemList(i).finfoDiv = rsget("infoDiv")
					FItemList(i).finfoDivName = db2html(rsget("infoDivName"))
					FItemList(i).finfoValidCnt = rsget("infoValidCnt")
					FItemList(i).fSafetyTargetYN = rsget("SafetyTargetYN")
					FItemList(i).fSafetyCertYN = rsget("SafetyCertYN")
					FItemList(i).fSafetyConfirmYN = rsget("SafetyConfirmYN")
					FItemList(i).fSafetySupplyYN = rsget("SafetySupplyYN")
					FItemList(i).fSafetyComply = rsget("SafetyComply")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).flastupdate = rsget("lastupdate")
					FItemList(i).flastadminid = rsget("lastadminid")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end Function

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function
	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function
	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

    Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10
	End Sub
	Private Sub Class_Terminate()
    End Sub
end Class
%>


