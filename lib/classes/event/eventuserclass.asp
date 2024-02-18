<%
'###########################################################
' Description :  이벤트 응모자  클래스
' History : 2007.09.06 한용민 생성
'###########################################################

Class Ceventuser
	public fuserid
	public fjuminno
	public fusername
	public fusermail
	public fuserphone
	public fusercell
	public fzipcode
	public faddress1
	public fuseraddr
	public fLevel
	public fevtcom_txt
	public fWcnt
	public fWdate
	public finvaliduserid
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

end class

class Ceventuserlist
	public flist
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

	public frectseachbox
	public frectgubun
	public frectinvaliduseryn

	public function frectseach()
		if frectseachbox <> "" then
			frectseach = "& frectseachbox &"
		else
			frectseach = 0
		end if
	end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function
	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  flist(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//admin/eventseach/event_user_list.asp
	public sub Feventuserlist3
		dim sql , i, sqlsearch
	
		if frectinvaliduseryn="Y" then
			sqlsearch = sqlsearch & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			sqlsearch = sqlsearch & " and iu.idx is null"
		end if
	
		sql = "select top " & FPagesize*FCurrPage
		sql = sql & " u.userid,u.juminno,u.username,u.userphone,u.usercell,u.zipcode"
		sql = sql & " ,replace(u.usermail,'\0x5F','_')as usermail"
		sql = sql & " ,replace(u.zipaddr,'\0x5F','_') as address1"
		sql = sql & " ,replace(u.useraddr,'\0x5F','_') as useraddr"
		sql = sql & " ,L.userlevel as Level"
		sql = sql & " ,c.comment, iu.invaliduserid"
		sql = sql & " from db_contents.dbo.tbl_one_comment c"
		sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
		sql = sql & " 	on c.userid= u.userid"
		sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
		sql = sql & " 	on u.userid=L.userid"
		sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
		sql = sql & " 	on c.userid=iu.invaliduserid"
		sql = sql & " 	and iu.isusing='Y'"
		sql = sql & " 	and iu.gubun='"&frectgubun&"'"
		sql = sql & " where c.isusing='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
		sql = sql & " order by u.username asc"
	
		'response.write sql & "<br>"
		rsget.open sql,dbget,1
	
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
		if not rsget.eof then
			do until rsget.eof
				set flist(i) = new Ceventuser
	
				flist(i).finvaliduserid = rsget("invaliduserid")
				flist(i).fuserid = rsget("userid")
				flist(i).fjuminno = rsget("juminno")
				flist(i).fusername = rsget("username")
				flist(i).fusermail = rsget("usermail")
				flist(i).fuserphone = rsget("userphone")
				flist(i).fusercell = rsget("usercell")
				flist(i).fzipcode = rsget("zipcode")
				flist(i).faddress1 = rsget("address1")
				flist(i).fuseraddr = rsget("useraddr")
				flist(i).fLevel = rsget("Level")
				flist(i).fevtcom_txt = rsget("comment")
	
				rsget.movenext
				i = i+1
				loop
			end if
		rsget.close
	end sub

	'//admin/eventseach/event_user_list.asp
	public sub Feventuserlist5
		dim sql , i, sqlsearch
	
		if frectinvaliduseryn="Y" then
			sqlsearch = sqlsearch & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			sqlsearch = sqlsearch & " and iu.idx is null"
		end if
	
		sql = "select top " & FPagesize*FCurrPage
		sql = sql & " u.userid,u.juminno,u.username,u.userphone,u.usercell,u.zipcode"
		sql = sql & " ,replace(u.usermail,'\0x5F','_')as usermail"
		sql = sql & " ,replace(u.zipaddr,'\0x5F','_') as address1"
		sql = sql & " ,replace(u.useraddr,'\0x5F','_') as useraddr"
		sql = sql & " ,L.userlevel as Level"
		sql = sql & " ,c.evtcom_txt, iu.invaliduserid"
		sql = sql & " from db_event.dbo.tbl_event_comment c"
		sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
		sql = sql & " 	on c.userid= u.userid"
		sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
		sql = sql & " 	on u.userid=L.userid"
		sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
		sql = sql & " 	on c.userid=iu.invaliduserid"
		sql = sql & " 	and iu.isusing='Y'"
		sql = sql & " 	and iu.gubun='"&frectgubun&"'"	
		sql = sql & " where c.evtcom_using='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
		sql = sql & " order by u.username asc"
	
		'response.write sql & "<br>"
		rsget.open sql,dbget,1
	
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount
			
		redim flist(FTotalCount)
		i = 0
	
		if not rsget.eof then
			do until rsget.eof
				set flist(i) = new Ceventuser
	
				flist(i).finvaliduserid = rsget("invaliduserid")
				flist(i).fuserid = rsget("userid")
				flist(i).fjuminno = rsget("juminno")
				flist(i).fusername = rsget("username")
				flist(i).fusermail = rsget("usermail")
				flist(i).fuserphone = rsget("userphone")
				flist(i).fusercell = rsget("usercell")
				flist(i).fzipcode = rsget("zipcode")
				flist(i).faddress1 = rsget("address1")
				flist(i).fuseraddr = rsget("useraddr")
				flist(i).fLevel = rsget("Level")
				flist(i).fevtcom_txt = rsget("evtcom_txt")
	
				rsget.movenext
				i = i+1
				loop
			end if
		rsget.close
	end sub

	'//admin/eventseach/event_user_list.asp
	public sub Feventuserlist7
		dim sql , i, sqlsearch
	
		if frectinvaliduseryn="Y" then
			sqlsearch = sqlsearch & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			sqlsearch = sqlsearch & " and iu.idx is null"
		end if
		
		sql = "select top " & FPagesize*FCurrPage
		sql = sql & " u.userid,u.juminno,u.username,u.userphone,u.usercell,u.zipcode"
		sql = sql & " ,replace(u.usermail,'\0x5F','_')as usermail"
		sql = sql & " ,replace(u.zipaddr,'\0x5F','_') as address1"
		sql = sql & " ,replace(u.useraddr,'\0x5F','_') as useraddr"
		sql = sql & " ,L.userlevel as Level"
		sql = sql & " ,c.comment, iu.invaliduserid"
		sql = sql & " from db_contents.dbo.tbl_weekly_codi_comment c"
		sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
		sql = sql & " 	on c.userid= u.userid"
		sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
		sql = sql & " 	on u.userid=L.userid"
		sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
		sql = sql & " 	on c.userid=iu.invaliduserid"
		sql = sql & " 	and iu.isusing='Y'"
		sql = sql & " 	and iu.gubun='"&frectgubun&"'"	
		sql = sql & " where c.isusing='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
		sql = sql & " order by u.username asc"
	
		'response.write sql & "<Br>"
		rsget.open sql,dbget,1
	
		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount
	
		redim flist(FTotalCount)
		i = 0
	
		if not rsget.eof then
			do until rsget.eof
				set flist(i) = new Ceventuser
	
				flist(i).finvaliduserid = rsget("invaliduserid")
				flist(i).fuserid = rsget("userid")
				flist(i).fjuminno = rsget("juminno")
				flist(i).fusername = rsget("username")
				flist(i).fusermail = rsget("usermail")
				flist(i).fuserphone = rsget("userphone")
				flist(i).fusercell = rsget("usercell")
				flist(i).fzipcode = rsget("zipcode")
				flist(i).faddress1 = rsget("address1")
				flist(i).fuseraddr = rsget("useraddr")
				flist(i).fLevel = rsget("Level")
				flist(i).fevtcom_txt = rsget("comment")
	
				rsget.movenext
				i = i+1
				loop
			end if
		rsget.close
	end sub

	'//admin/eventseach/event_user_list.asp
	public sub Feventuserlist99
		dim sql , i, sqlsearch
	
			if frectinvaliduseryn="Y" then
				sqlsearch = sqlsearch & " and iu.idx is not null"
			elseif frectinvaliduseryn="N" then
				sqlsearch = sqlsearch & " and iu.idx is null"
			end if
	
			sql = "Select count(*), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")"
			sql = sql & " from db_culture_station.dbo.tbl_culturestation_event_comment c"
			sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
			sql = sql & " 	on c.userid= u.userid"
			sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
			sql = sql & " 	on u.userid=L.userid"
			sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
			sql = sql & " 	on c.userid=iu.invaliduserid"
			sql = sql & " 	and iu.isusing='Y'"
			sql = sql & " 	and iu.gubun='"&frectgubun&"'"
			sql = sql &	" where c.isusing='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
		
			'response.write sql & "<Br>"		
			rsget.Open SQL,dbget,1
				FTotalCount = rsget(0)
				FtotalPage = rsget(1)
			rsget.Close
			
			if FTotalCount < 1 then exit sub
	
			sql = "select top " & CStr(FPageSize*FCurrPage)
			sql = sql & " u.userid,u.juminno,u.username,u.userphone,u.usercell,u.zipcode"
			sql = sql & " ,replace(u.usermail,'\0x5F','_')as usermail"
			sql = sql & " ,replace(u.zipaddr,'\0x5F','_') as address1"
			sql = sql & " ,replace(u.useraddr,'\0x5F','_') as useraddr"
			sql = sql & " ,L.userlevel as Level"
			sql = sql & " ,c.comment, iu.invaliduserid"
			sql = sql & " ,(select count(*) from [db_event].[dbo].[tbl_event_prize] as z where z.evt_winner = c.userid) as Wcnt"
			sql = sql & " ,(select top 1 evt_regdate from [db_event].[dbo].[tbl_event_prize] as zz where zz.evt_winner = c.userid) as Wdate"
			sql = sql & " from db_culture_station.dbo.tbl_culturestation_event_comment c"
			sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
			sql = sql & " 	on c.userid= u.userid"
			sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
			sql = sql & " 	on u.userid=L.userid"
			sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
			sql = sql & " 	on c.userid=iu.invaliduserid"
			sql = sql & " 	and iu.isusing='Y'"
			sql = sql & " 	and iu.gubun='"&frectgubun&"'"
			sql = sql & " where c.isusing='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
			sql = sql & " order by u.username asc"
	
			'response.write sql & "<Br>"
			rsget.pagesize = FPageSize
			rsget.Open SQL,dbget,1
			FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
			if FResultCount<1 then FResultCount=0
			redim preserve flist(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof
				set flist(i) = new Ceventuser
	
				flist(i).finvaliduserid = rsget("invaliduserid")
				flist(i).fuserid = rsget("userid")
				flist(i).fjuminno = rsget("juminno")
				flist(i).fusername = rsget("username")
				flist(i).fusermail = rsget("usermail")
				flist(i).fuserphone = rsget("userphone")
				flist(i).fusercell = rsget("usercell")
				flist(i).fzipcode = rsget("zipcode")
				flist(i).faddress1 = rsget("address1")
				flist(i).fuseraddr = rsget("useraddr")
				flist(i).fLevel = rsget("Level")
				flist(i).fevtcom_txt = db2html(rsget("comment"))
				flist(i).fWcnt = rsget("Wcnt")
				flist(i).fWdate = left(rsget("Wdate"),10)
	
				rsget.movenext
				i = i+1
				loop
			end if
		rsget.close
	end sub

	''엑셀 다운용
	public sub Feventuserlist9				'컬쳐스테이션
		dim sql , i, sqlsearch
	
		if frectinvaliduseryn="Y" then
			sqlsearch = sqlsearch & " and iu.idx is not null"
		elseif frectinvaliduseryn="N" then
			sqlsearch = sqlsearch & " and iu.idx is null"
		end if
	
		sql = "select top " & CStr(FPageSize*FCurrPage)
		sql = sql & " u.userid,u.juminno,u.username,u.userphone,u.usercell,u.zipcode"
		sql = sql & " ,replace(u.usermail,'\0x5F','_')as usermail"
		sql = sql & " ,replace(u.zipaddr,'\0x5F','_') as address1"
		sql = sql & " ,replace(u.useraddr,'\0x5F','_') as useraddr"
		sql = sql & " ,L.userlevel as Level"
		sql = sql & " ,c.comment, iu.invaliduserid"
		sql = sql & " ,(select count(*) from [db_event].[dbo].[tbl_event_prize] as z where z.evt_winner = c.userid) as Wcnt"
		sql = sql & " ,(select top 1 evt_regdate from [db_event].[dbo].[tbl_event_prize] as zz where zz.evt_winner = c.userid order by evt_regdate desc) as Wdate "
		sql = sql & " from db_culture_station.dbo.tbl_culturestation_event_comment c"
		sql = sql & " JOIN [db_user].dbo.tbl_user_n u"
		sql = sql & " 	on c.userid= u.userid"
		sql = sql & " JOIN [db_user].dbo.tbl_logindata L"
		sql = sql & " 	on u.userid=L.userid"
		sql = sql & " left join db_user.dbo.tbl_invalid_user iu"
		sql = sql & " 	on c.userid=iu.invaliduserid"
		sql = sql & " 	and iu.isusing='Y'"
		sql = sql & " 	and iu.gubun='"&frectgubun&"'"	
		sql = sql & " where c.isusing='Y' and c.evt_code='"& frectseachbox &"' " & sqlsearch
		sql = sql & " order by u.username asc"
	
		'response.write sql & "<Br>"
		rsget.open sql,dbget,1
	
		FTotalCount = rsget.recordcount
		redim flist(FTotalCount)
		i = 0
	
		if not rsget.eof then
			do until rsget.eof
				set flist(i) = new Ceventuser
	
				flist(i).finvaliduserid = rsget("invaliduserid")
				flist(i).fuserid = rsget("userid")
				flist(i).fjuminno = rsget("juminno")
				flist(i).fusername = rsget("username")
				flist(i).fusermail = rsget("usermail")
				flist(i).fuserphone = rsget("userphone")
				flist(i).fusercell = rsget("usercell")
				flist(i).fzipcode = rsget("zipcode")
				flist(i).faddress1 = rsget("address1")
				flist(i).fuseraddr = rsget("useraddr")
				flist(i).fLevel = rsget("Level")
				flist(i).fevtcom_txt = db2html(rsget("comment"))
				flist(i).fWcnt = rsget("Wcnt")
				flist(i).fWdate = left(rsget("Wdate"),10)
				rsget.movenext
				i = i+1
				loop
			end if
		rsget.close
	end sub
end class

Sub DraweventGubun(eventbox, usinguserid)
	dim userquery, tem_str

	response.write "<select name='" & eventbox & "'>"		
	response.write "<option value=''"							
	response.write ">선택</option>"								
	
	response.write "<option value='3'"
		if usinguserid = "3" then								
			response.write "selected"
		end if
	response.write ">한줄낙서</option>"
	
	response.write "<option value='5'"
		if usinguserid = "5" then								
			response.write "selected"
		end if
	response.write ">문화이벤트</option>"
	
	response.write "<option value='7'"
		if usinguserid = "7" then								
			response.write "selected"
		end if
	response.write ">위클리코디네이터</option>"

	response.write "<option value='9'"
		if usinguserid = "9" then								
			response.write "selected"
		end if
	response.write ">CultureStation</option>"

	response.write "</select>"
End Sub
%>