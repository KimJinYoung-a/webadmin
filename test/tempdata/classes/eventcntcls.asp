<%

class Ceventuserlist

	public flist

	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

'##################################################################
public sub Feventuserlist3
	dim sql , i
	'mydate = "2013-02-13"

		sql="select count(distinct(userid))as totalcount from db_event.dbo.tbl_event_subscript where evt_code='40245'"
		rsget.open sql,dbget,1

		FTotalcount = rsget("totalcount")

		rsget.close




if Ftotalcount > 0 Then

		sql= "select count(userid) as cnt, convert(varchar(10),regdate,120)"
		sql= sql & "from db_event.dbo.tbl_event_subscript "
		sql= sql & "where evt_code='40245' and regdate between '2013-02-13 00:00:00' and '2013-02-24 23:59:59'"
		sql= sql & "group by convert(varchar(10),regdate,120)"
		sql= sql & "order by convert(varchar(10),regdate,120)"
		'response.write sql
		rsget.open sql,dbget,1

	FResultCount = rsget.recordcount
		redim flist(FResultCount)
		i = 0

	if not rsget.eof then
		do until rsget.eof
			flist(i)= rsget("cnt")
			rsget.movenext
			i = i+1
		loop
	end if
		rsget.close
End If
	end sub

'##################################################################
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class
%>