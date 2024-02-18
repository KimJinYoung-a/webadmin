<%
'###########################################################
' Description :  비회원메일진 클래스
' History : 2009.10.08 한용민 생성
'###########################################################
%>
<%
class CMailzineListSubItem

	public Fidx
	public Fusername
	public Fusermail
	public Fregdate
	public Fisusing

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end class

class CMailzineList
	public FItemList()
	public FTotalCount
	public FOneItem
	public FResultCount
	public FRectDesignerID
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FPCount

	public frectidx
	public frectisusing
	public frectusermail
	public frectusername
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

    public Sub Mailzine_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx,username,usermail,regdate,isusing" & vbcrlf
		sqlStr = sqlStr & " from db_user.dbo.tbl_mailzine_notmember with (nolock)" & vbcrlf
        sqlStr = sqlStr & " where idx = "& frectidx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new CMailzineListSubItem
        
        if Not rsget.Eof then

			FOneItem.Fidx = rsget("idx")
			FOneItem.fusername = db2html(rsget("username"))
			FOneItem.fusermail = db2html(rsget("usermail"))
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fisusing = rsget("isusing")
			           
        end if
        rsget.Close
    end Sub

	' /cscenter/mailzine/cs_mailzine.asp
	public sub MailzineList()
		dim sqlStr,i				
		
		if frectusermail="" and frectusername="" then exit sub

		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_user.dbo.tbl_mailzine_notmember with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
		
		if frectisusing <> "" then
		sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if
		if frectusermail <> "" then
		sqlStr = sqlStr & " and usermail like '%"&html2db(frectusermail)&"%'" + vbcrlf
		end if
		
		if frectusername <> "" then
		sqlStr = sqlStr & " and username like '%"&html2db(frectusername)&"%'" + vbcrlf
		end if	

		'response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount<1 then exit sub

		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,username,usermail,regdate,isusing" + vbcrlf
		sqlStr = sqlStr & " from db_user.dbo.tbl_mailzine_notmember with (nolock)" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf	

		if frectisusing <> "" then
		sqlStr = sqlStr & " and isusing='"&frectisusing&"'" + vbcrlf
		end if
		if frectusermail <> "" then
		sqlStr = sqlStr & " and usermail like '%"&html2db(frectusermail)&"%'" + vbcrlf
		end if
	
		
		if frectusername <> "" then
		sqlStr = sqlStr & " and username like '%"&html2db(frectusername)&"%'" + vbcrlf
		end if	
		
		sqlStr = sqlStr & " order by idx Desc" + vbcrlf

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new CMailzineListSubItem
				
				FItemList(i).Fidx = rsget("idx")
				FItemList(i).fusername = db2html(rsget("username"))
			    FItemList(i).fusermail = db2html(rsget("usermail"))
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")	
							
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

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