<%

class CFlashMasterItem

	public Feachcnt
	public Fqadiv

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CFlashItemImg

	public FMasterItemList()
	public FResultCount


    Private Sub Class_Initialize()
			redim FMasterItemList(0)
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub ItemList()
		dim sqlStr
		dim i

		''#################################################
		''µ¥ÀÌÅ¸
		''#################################################
		sqlStr = "select qadiv,count(qadiv) as count from [db_cs].[10x10].tbl_myqna" + vbcrlf
		sqlStr = sqlStr + " where regdate >= '2004-03-01'" + vbcrlf
		sqlStr = sqlStr + " and regdate < '2004-03-03'" + vbcrlf
		sqlStr = sqlStr + " group by qadiv" + vbcrlf
		sqlStr = sqlStr + " order by qadiv asc"

'response.write sqlStr
		rsget.Open sqlStr,dbget,1

			FResultCount = rsget.recordcount

		redim preserve FMasterItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
		do until rsget.EOF
					set FMasterItemList(i) = new CFlashMasterItem
					FMasterItemList(i).Fqadiv = rsget("qadiv")
					FMasterItemList(i).Feachcnt = rsget("count")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

end Class

%>