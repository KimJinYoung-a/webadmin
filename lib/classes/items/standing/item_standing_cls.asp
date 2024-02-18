<%
'###########################################################
' Description : 정기구독 상품 클래스
' History : 2016.06.16 한용민 생성
'###########################################################

Class Citemstanding_oneitem
	public fitemid
	public fitemoption
	public foptionname
	public freserveItemName
	public freserveDlvDate
	public freserveItemGubun
	public freserveItemID
	public freserveItemOption
	public freserveidx
	public fuidx
	public forgitemid
	public forgitemoption
	public forderserial
	public fuserid
	public fitemno
	public fsendstatus
	public fsenddate
	public fusername
	public fzipcode
	public freqzipaddr
	public fuseraddr
	public fuserphone
	public fusercell
	public fisusing
	public fregdate
	public fregadminid
	public flastupdate
	public flastadminid
	public fjukyogubun
	public fstandingusercount
	public fstartreserveidx
	public fendreserveidx
	public FMakerid
	public fitemname
	public fitemoptionname
	public freqname_u
	public freqzipcode_u
	public freqzipaddr_u
	public freqaddress_u
	public freqphone_u
	public freqhp_u

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Citemstanding
	Public FItemList()
	public foneitem
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FPageCount

	public fstandingarr

	public frectitemgubun
	public frectitemid
	public FRectreserveitemid
	public frectitemoption
	public frectreserveidx
	public frectorderserial
	public frectuserid
	public frectusername
	public frectisusing
	public frectuidx
	public FRectsendstatus
	public FRectjukyogubun
	public fdbdatamart

	public sub fitemstanding_user_one()
		dim SqlStr, sqlsearch, i

		if frectuidx="" then exit sub

		if frectreserveidx <> "" then
			sqlsearch = sqlsearch & " and su.reserveidx="& frectreserveidx &""
		end if
		if frectuidx <> "" then
			sqlsearch = sqlsearch & " and su.uidx="& frectuidx &""
		end if

		sqlStr = " select top 1" & vbcrlf
		sqlStr = sqlStr & " su.uidx, su.orgitemid, su.orgitemoption, su.reserveidx, su.jukyogubun, su.orderserial, su.userid, su.itemno"
		sqlStr = sqlStr & " , su.sendstatus, su.senddate, isnull(u.username,su.username) as username, isnull(u.zipcode,su.zipcode) as zipcode"
		sqlStr = sqlStr & " , isnull(u.zipaddr,su.reqzipaddr) as reqzipaddr, isnull(u.useraddr,su.useraddr) as useraddr"
		sqlStr = sqlStr & " , isnull(u.userphone,su.userphone) as userphone, isnull(u.usercell,su.usercell) as usercell"
		sqlStr = sqlStr & " , su.isusing, su.regdate, su.regadminid, su.lastupdate, su.lastadminid"
		sqlStr = sqlStr & " , u.username as reqname_u" & vbcrlf
		sqlStr = sqlStr & " , u.zipcode as reqzipcode_u, u.zipaddr as reqzipaddr_u" & vbcrlf
		sqlStr = sqlStr & " , u.useraddr as reqaddress_u" & vbcrlf
		sqlStr = sqlStr & " , u.userphone as reqphone_u, u.usercell as reqhp_u" & vbcrlf
		sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_user] su"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on su.userid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by uidx desc"

		'Response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new Citemstanding_oneitem
				FOneItem.fjukyogubun = rsget("jukyogubun")
				FOneItem.fuidx = rsget("uidx")
				FOneItem.forgitemid = rsget("orgitemid")
				FOneItem.forgitemoption = rsget("orgitemoption")
				FOneItem.freserveidx = rsget("reserveidx")
				FOneItem.forderserial = rsget("orderserial")
				FOneItem.fuserid = rsget("userid")
				FOneItem.fitemno = rsget("itemno")
				FOneItem.fsendstatus = rsget("sendstatus")
				FOneItem.fsenddate = rsget("senddate")
				FOneItem.fusername = db2html(rsget("username"))
				FOneItem.fzipcode = rsget("zipcode")
				FOneItem.freqzipaddr = db2html(rsget("reqzipaddr"))
				FOneItem.fuseraddr = db2html(rsget("useraddr"))
				FOneItem.fuserphone = rsget("userphone")
				FOneItem.fusercell = rsget("usercell")
				FOneItem.fisusing = rsget("isusing")
				FOneItem.fregdate = rsget("regdate")
				FOneItem.fregadminid = rsget("regadminid")
				FOneItem.flastupdate = rsget("lastupdate")
				FOneItem.flastadminid = rsget("lastadminid")
				FOneItem.freqname_u = db2html(rsget("reqname_u"))
				FOneItem.freqzipcode_u = rsget("reqzipcode_u")
				FOneItem.freqzipaddr_u = db2html(rsget("reqzipaddr_u"))
				FOneItem.freqaddress_u = db2html(rsget("reqaddress_u"))
				FOneItem.freqphone_u = rsget("reqphone_u")
				FOneItem.freqhp_u = rsget("reqhp_u")
		end if
		rsget.close
	end sub

	' 이펑션 수정시 fitemstanding_user_getrows 이펑션도 똑같이 수정해야함
	public function fitemstanding_user()
		dim sqlStr, i, sqlsearch

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and su.orgitemid="& frectitemid &""
		end if
		if FRectreserveitemid <> "" then
			sqlsearch = sqlsearch & " and so.reserveitemid="& FRectreserveitemid &""
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and su.orgitemoption='"& frectitemoption &"'"
		end if
		if frectreserveidx <> "" then
			sqlsearch = sqlsearch & " and su.reserveidx='"& frectreserveidx &"'"
		end if
		if frectorderserial <> "" then
			sqlsearch = sqlsearch & " and su.orderserial='"& frectorderserial &"'"
		end if
		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and m.userid='"& frectuserid &"'"
		end if
		if frectusername <> "" then
			sqlsearch = sqlsearch & " and (m.reqname='"& trim(frectusername) &"' or u.username='"& trim(frectusername) &"' or su.username='"& trim(frectusername) &"')" & vbcrlf
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and su.isusing='"& frectisusing &"'"
		end if
		if FRectsendstatus = "05" then
			sqlsearch = sqlsearch & " and su.sendstatus in (0,5)"
		elseif FRectsendstatus = "37" then
			sqlsearch = sqlsearch & " and su.sendstatus in (3,7)"
		elseif FRectsendstatus <> "" then
			sqlsearch = sqlsearch & " and su.sendstatus="& FRectsendstatus &""
		end if
		if FRectjukyogubun <> "" then
			sqlsearch = sqlsearch & " and su.jukyogubun='"& FRectjukyogubun &"'"
		end if

		sqlStr = " select count(*) as cnt"
		sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_user] su"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item_standing_order so" & vbcrlf
		sqlStr = sqlStr & " 	on su.orgitemid = so.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and su.orgitemoption = so.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " 	and su.reserveidx = so.reserveidx" & vbcrlf
		sqlStr = sqlStr & " left join db_order.dbo.tbl_order_master m with (nolock)"
		sqlStr = sqlStr & " 	on su.orderserial = m.orderserial"
		sqlStr = sqlStr & " 	and m.cancelyn='N'"
		sqlStr = sqlStr & " 	and m.ipkumdiv=7"
		sqlStr = sqlStr & " 	and m.jumundiv not in (6,9)"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on su.userid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit function

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " su.uidx, su.orgitemid, su.orgitemoption, su.reserveidx, su.jukyogubun, isnull(m.orderserial,su.orderserial) as orderserial"
		sqlStr = sqlStr & " , isnull(m.userid,su.userid) as userid"
		sqlStr = sqlStr & " , isnull(isnull((select sum(itemno)"
		sqlStr = sqlStr & " 	from db_order.dbo.tbl_order_detail"
		sqlStr = sqlStr & " 	where su.orderserial=orderserial"
		sqlStr = sqlStr & " 	and cancelyn='A'"
		sqlStr = sqlStr & " 	and so.reserveItemID=ItemID and so.reserveitemoption=itemoption"
		sqlStr = sqlStr & " 	),su.itemno),0) as itemno"
		sqlStr = sqlStr & " , su.sendstatus, su.senddate, isnull(isnull(m.reqname,u.username),su.username) as reqname" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqzipcode,u.zipcode),su.zipcode) as reqzipcode, isnull(isnull(m.reqzipaddr,u.zipaddr),su.reqzipaddr) as reqzipaddr" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqaddress,u.useraddr),su.useraddr) as reqaddress" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqphone,u.userphone),su.userphone) as reqphone, isnull(isnull(m.reqhp,u.usercell),su.usercell) as reqhp" & vbcrlf
		sqlStr = sqlStr & " , su.isusing, su.regdate, su.regadminid, su.lastupdate, su.lastadminid"
		sqlStr = sqlStr & " , so.reserveItemID, so.reserveItemoption, replace(replace(replace( so.reserveItemname ,char(9),''),char(10),''),char(13),'') as reserveItemname"
		sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_user] su"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item_standing_order so" & vbcrlf
		sqlStr = sqlStr & " 	on su.orgitemid = so.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and su.orgitemoption = so.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " 	and su.reserveidx = so.reserveidx" & vbcrlf
		sqlStr = sqlStr & " left join db_order.dbo.tbl_order_master m with (nolock)"
		sqlStr = sqlStr & " 	on su.orderserial = m.orderserial"
		sqlStr = sqlStr & " 	and m.cancelyn='N'"
		sqlStr = sqlStr & " 	and m.ipkumdiv=7"
		sqlStr = sqlStr & " 	and m.jumundiv not in (6,9)"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on su.userid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by su.uidx desc"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Citemstanding_oneitem
					FItemList(i).fjukyogubun = rsget("jukyogubun")
					FItemList(i).fuidx = rsget("uidx")
					FItemList(i).forgitemid = rsget("orgitemid")
					FItemList(i).forgitemoption = rsget("orgitemoption")
					FItemList(i).freserveidx = rsget("reserveidx")
					FItemList(i).forderserial = rsget("orderserial")
					FItemList(i).fuserid = rsget("userid")
					FItemList(i).fitemno = rsget("itemno")
					FItemList(i).fsendstatus = rsget("sendstatus")
					FItemList(i).fsenddate = rsget("senddate")
					FItemList(i).fusername = db2html(rsget("reqname"))
					FItemList(i).fzipcode = rsget("reqzipcode")
					FItemList(i).freqzipaddr = db2html(rsget("reqzipaddr"))
					FItemList(i).fuseraddr = db2html(rsget("reqaddress"))
					FItemList(i).fuserphone = rsget("reqphone")
					FItemList(i).fusercell = rsget("reqhp")
					FItemList(i).fisusing = rsget("isusing")
					FItemList(i).fregdate = rsget("regdate")
					FItemList(i).fregadminid = rsget("regadminid")
					FItemList(i).flastupdate = rsget("lastupdate")
					FItemList(i).flastadminid = rsget("lastadminid")
					FItemList(i).freserveItemID = rsget("reserveItemID")
					FItemList(i).freserveItemoption = rsget("reserveItemoption")
					FItemList(i).freserveItemname = db2html(rsget("reserveItemname"))
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end function

	' 이펑션 수정시 fitemstanding_user 이펑션도 똑같이 수정해야함
	public function fitemstanding_user_getrows()
		dim sqlStr, i, sqlsearch

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and su.orgitemid="& frectitemid &""
		end if
		if FRectreserveitemid <> "" then
			sqlsearch = sqlsearch & " and so.reserveitemid="& FRectreserveitemid &""
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and su.orgitemoption='"& frectitemoption &"'"
		end if
		if frectreserveidx <> "" then
			sqlsearch = sqlsearch & " and su.reserveidx='"& frectreserveidx &"'"
		end if
		if frectorderserial <> "" then
			sqlsearch = sqlsearch & " and su.orderserial='"& frectorderserial &"'"
		end if
		if frectuserid <> "" then
			sqlsearch = sqlsearch & " and m.userid='"& frectuserid &"'"
		end if
		if frectusername <> "" then
			sqlsearch = sqlsearch & " and (m.reqname='"& trim(frectusername) &"' or u.username='"& trim(frectusername) &"' or su.username='"& trim(frectusername) &"')" & vbcrlf
		end if
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and su.isusing='"& frectisusing &"'"
		end if
		if FRectsendstatus = "05" then
			sqlsearch = sqlsearch & " and su.sendstatus in (0,5)"
		elseif FRectsendstatus = "37" then
			sqlsearch = sqlsearch & " and su.sendstatus in (3,7)"
		elseif FRectsendstatus <> "" then
			sqlsearch = sqlsearch & " and su.sendstatus="& FRectsendstatus &""
		end if
		if FRectjukyogubun <> "" then
			sqlsearch = sqlsearch & " and su.jukyogubun='"& FRectjukyogubun &"'"
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " su.uidx, su.orgitemid, su.orgitemoption, su.reserveidx, su.jukyogubun, isnull(m.orderserial,su.orderserial) as orderserial"
		sqlStr = sqlStr & " , isnull(m.userid,su.userid) as userid"
		sqlStr = sqlStr & " , isnull(isnull((select sum(itemno)"
		sqlStr = sqlStr & " 	from db_order.dbo.tbl_order_detail"
		sqlStr = sqlStr & " 	where su.orderserial=orderserial"
		sqlStr = sqlStr & " 	and cancelyn='A'"
		sqlStr = sqlStr & " 	and so.reserveItemID=ItemID and so.reserveitemoption=itemoption"
		sqlStr = sqlStr & " 	),su.itemno),0) as itemno"
		sqlStr = sqlStr & " , su.sendstatus, su.senddate, isnull(isnull(m.reqname,u.username),su.username) as reqname" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqzipcode,u.zipcode),su.zipcode) as reqzipcode, isnull(isnull(m.reqzipaddr,u.zipaddr),su.reqzipaddr) as reqzipaddr" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqaddress,u.useraddr),su.useraddr) as reqaddress" & vbcrlf
		sqlStr = sqlStr & " , isnull(isnull(m.reqphone,u.userphone),su.userphone) as reqphone, isnull(isnull(m.reqhp,u.usercell),su.usercell) as reqhp" & vbcrlf
		sqlStr = sqlStr & " , su.isusing, su.regdate, su.regadminid, su.lastupdate, su.lastadminid"
		sqlStr = sqlStr & " , so.reserveItemID, so.reserveItemoption, replace(replace(replace( so.reserveItemname ,char(9),''),char(10),''),char(13),'') as reserveItemname"
		sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_user] su"
		sqlStr = sqlStr & " join db_item.dbo.tbl_item_standing_order so" & vbcrlf
		sqlStr = sqlStr & " 	on su.orgitemid = so.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and su.orgitemoption = so.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " 	and su.reserveidx = so.reserveidx" & vbcrlf
		sqlStr = sqlStr & " left join db_order.dbo.tbl_order_master m with (nolock)"
		sqlStr = sqlStr & " 	on su.orderserial = m.orderserial"
		sqlStr = sqlStr & " 	and m.cancelyn='N'"
		sqlStr = sqlStr & " 	and m.ipkumdiv=7"
		sqlStr = sqlStr & " 	and m.jumundiv not in (6,9)"
		sqlStr = sqlStr & " left join db_user.dbo.tbl_user_n u with (readuncommitted)" & vbcrlf
		sqlStr = sqlStr & " 	on su.userid = u.userid" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by su.uidx desc"

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount=rsget.recordcount
		ftotalcount=rsget.recordcount

		if not rsget.EOF then
			fstandingarr = rsget.getrows()
		end if
		rsget.Close
	end function

	public sub fitemstanding_one()
		dim SqlStr, sqlsearch, i

		if frectitemid="" or frectitemoption="" then exit sub

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and so.orgitemid="& frectitemid &""
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and so.orgitemoption='"& frectitemoption &"'"
		end if
		if frectreserveidx <> "" then
			sqlsearch = sqlsearch & " and so.reserveidx='"& frectsendkey &"'"
		end if

		sqlStr = " select top 1" & vbcrlf
		sqlStr = sqlStr & " so.orgitemid, replace(replace(replace( so.reserveItemname ,char(9),''),char(10),''),char(13),'') as reserveItemname, so.reserveDlvDate, so.reserveItemGubun, so.reserveItemID" & vbcrlf
		sqlStr = sqlStr & " , so.reserveItemOption, so.reserveidx, so.regdate, so.regadminid, so.lastupdate, so.lastadminid" & vbcrlf
		sqlStr = sqlStr & " ,o.itemid, o.itemoption, o.isusing, o.optionname" & vbcrlf
		sqlStr = sqlStr & " from db_item.[dbo].[tbl_item_standing_order] so" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_item_option o" & vbcrlf
		sqlStr = sqlStr & " 	on so.orgitemid=o.itemid" & vbcrlf
		sqlStr = sqlStr & " 	and so.orgitemoption=o.itemoption" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'Response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new Citemstanding_oneitem	
				FOneItem.forgitemid      = rsget("orgitemid")
				FOneItem.fitemid      = rsget("itemid")
				FOneItem.fitemoption      = rsget("itemoption")
				FOneItem.fisusing      = rsget("isusing")
				FOneItem.foptionname      = db2html(rsget("optionname"))
				FOneItem.freserveItemName      = db2html(rsget("reserveItemName"))
				FOneItem.freserveDlvDate      = rsget("reserveDlvDate")
				FOneItem.freserveItemGubun      = rsget("reserveItemGubun")
				FOneItem.freserveItemID      = rsget("reserveItemID")
				FOneItem.freserveItemOption      = rsget("reserveItemOption")
				FOneItem.freserveidx      = rsget("reserveidx")
				FOneItem.fregdate      = rsget("regdate")
				FOneItem.fregadminid      = rsget("regadminid")
				FOneItem.flastupdate      = rsget("lastupdate")
				FOneItem.flastadminid      = rsget("lastadminid")
		end if
		rsget.close
	end sub

	public sub fitemstanding_item()
		dim SqlStr, sqlsearch, i

		if frectitemid="" or frectitemoption="" then exit sub

		if frectitemid <> "" then
			sqlsearch = sqlsearch & " and o.itemid="& frectitemid &""
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and o.itemoption='"& frectitemoption &"'"
		end if

		sqlStr = " select top 1" & vbcrlf
		sqlStr = sqlStr & " i.itemid, o.itemoption, i.makerid, i.itemname, o.optionname, o.isusing, si.startreserveidx, si.endreserveidx" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_item i with (nolock)" & vbcrlf
		sqlStr = sqlStr & " join db_item.dbo.tbl_item_option o with (nolock)" & vbcrlf
		sqlStr = sqlStr & " 	on i.itemid=o.itemid" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_standing_item si" & vbcrlf
		sqlStr = sqlStr & " 	on o.itemid=si.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and o.itemoption=si.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'Response.write sqlStr &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		
		if not rsget.EOF then
			set FOneItem = new Citemstanding_oneitem	
				FOneItem.fitemid      = rsget("itemid")
				FOneItem.fitemoption      = rsget("itemoption")
				FOneItem.fstartreserveidx      = rsget("startreserveidx")
				FOneItem.fendreserveidx      = rsget("endreserveidx")
				FOneItem.FMakerid      = rsget("makerid")
				FOneItem.fitemname      = db2html(rsget("itemname"))
				FOneItem.fitemoptionname      = db2html(rsget("optionname"))
				FOneItem.fisusing      = rsget("isusing")
		end if
		rsget.close
	end sub

	public Function fitemstanding_option()
		dim sqlStr, i, sqlsearch

		if frectitemid="" or frectitemoption="" then exit Function

		if frectitemid<>"" then
			sqlsearch = sqlsearch & " and o.itemid = "& frectitemid &"" & vbcrlf
		end if
		if frectitemoption <> "" then
			sqlsearch = sqlsearch & " and o.itemoption='"& frectitemoption &"'"
		end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " o.itemid, o.itemoption, o.isusing, o.optionname" & vbcrlf
		sqlStr = sqlStr & " ,so.orgitemid, replace(replace(replace( so.reserveItemname ,char(9),''),char(10),''),char(13),'') as reserveItemname, so.reserveDlvDate, so.reserveItemGubun, so.reserveItemID" & vbcrlf
		sqlStr = sqlStr & " , so.reserveItemOption, so.reserveidx" & vbcrlf
		sqlStr = sqlStr & " , isnull((case" & vbcrlf
		sqlStr = sqlStr & " 	when si.startreserveidx=so.reserveidx then" & vbcrlf	' 첫회차일경우 주문디비 긁는다.
		sqlStr = sqlStr & " 		0" & vbcrlf
		' sqlStr = sqlStr & " 		(select count(d.orderserial)" & vbcrlf
		' sqlStr = sqlStr & " 		from db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
		' sqlStr = sqlStr & " 		where d.cancelyn<>'Y'" & vbcrlf
		' sqlStr = sqlStr & " 		and d.itemid=so.orgitemid" & vbcrlf
		' sqlStr = sqlStr & " 		and d.itemoption=so.orgitemoption)" & vbcrlf
		sqlStr = sqlStr & " 	else" & vbcrlf
		sqlStr = sqlStr & " 		(select count(su.orderserial)" & vbcrlf
		sqlStr = sqlStr & " 		from db_item.[dbo].[tbl_item_standing_user] su" & vbcrlf
		sqlStr = sqlStr & " 		where su.isusing='Y'" & vbcrlf
		sqlStr = sqlStr & " 		and o.itemid=su.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 		and o.itemoption=su.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " 		and so.reserveidx=su.reserveidx)" & vbcrlf
		sqlStr = sqlStr & " 	end),0) as standingusercount" & vbcrlf
		sqlStr = sqlStr & " from db_item.dbo.tbl_item_option o with (nolock)" & vbcrlf
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_standing_item si" & vbcrlf
		sqlStr = sqlStr & " 	on o.itemid=si.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and o.itemoption=si.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " left join db_item.[dbo].[tbl_item_standing_order] so" & vbcrlf
		sqlStr = sqlStr & " 	on o.itemid = so.orgitemid" & vbcrlf
		sqlStr = sqlStr & " 	and o.itemoption = so.orgitemoption" & vbcrlf
		sqlStr = sqlStr & " where 1=1 " & sqlsearch
		sqlStr = sqlStr & " order by so.reserveidx asc" & vbcrlf

		'response.write sqlStr & "<Br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount
		ftotalcount = rsget.RecordCount

		i=0		
		if  not rsget.EOF  then
			redim preserve FItemList(FResultCount)

			do until rsget.eof
				set FItemList(i) = new Citemstanding_oneitem
					FItemList(i).forgitemid      = rsget("orgitemid")
					FItemList(i).fitemid      = rsget("itemid")
					FItemList(i).fitemoption      = rsget("itemoption")
					FItemList(i).fisusing      = rsget("isusing")
					FItemList(i).foptionname      = db2html(rsget("optionname"))
					FItemList(i).freserveItemName      = db2html(rsget("reserveItemName"))
					FItemList(i).freserveDlvDate      = rsget("reserveDlvDate")
					FItemList(i).freserveItemGubun      = rsget("reserveItemGubun")
					FItemList(i).freserveItemID      = rsget("reserveItemID")
					FItemList(i).freserveItemOption      = rsget("reserveItemOption")
					FItemList(i).freserveidx      = rsget("reserveidx")
					FItemList(i).fstandingusercount      = rsget("standingusercount")
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

		IF application("Svr_Info")<>"Dev" THEN
			fdbdatamart="dbdatamart."
		end if
	End Sub
	Private Sub Class_Terminate()
    End Sub
End Class

' 회차 모두 가져오기
function drawSelectBoxsendkey(selectBoxName, selectedId, itemid, itemoption, chplg)
	dim tmp_str, query1, i

	if itemid="" or itemoption="" then exit function
	%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
	<%
		query1 = " select reserveidx from db_item.[dbo].[tbl_item_standing_order]"
		query1 = query1 & " where 1=1 "

		if itemid <> "" then
			query1 = query1 & " and orgitemid="& itemid &""
		end if
		if itemoption <> "" then
			query1 = query1 & " and orgitemoption='"& itemoption &"'"
		end if

		query1 = query1 & " group by reserveidx"
   		query1 = query1 & " order by reserveidx asc"

		'response.write query1 & "<br>"
		rsget.Open query1,dbget,1

		if not rsget.EOF  then
		rsget.Movefirst

		do until rsget.EOF
		if Lcase(selectedId) = Lcase(rsget("reserveidx")) then
			tmp_str = " selected"
		end if
		response.write("<option value='"&rsget("reserveidx")&"' "&tmp_str&">"&rsget("reserveidx")&"</option>")
		tmp_str = ""
		rsget.MoveNext
		loop
		end if
		rsget.close
	response.write("</select>")
	'response.write query1 &"<Br>"
end function

function getsendkey(itemgubun, itemid, itemoption)
	dim query1, tmpsendkey, i

	if itemid="" or itemoption="" then exit function

	query1 = " select max(sendkey) as sendkey from db_item.[dbo].[tbl_item_standing_order]"
	query1 = query1 & " where 1=1 "

	if itemid <> "" then
		query1 = query1 & " and orgitemid="& itemid &""
	end if
	if itemoption <> "" then
		query1 = query1 & " and orgitemoption='"& itemoption &"'"
	end if

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		tmpsendkey = rsget("sendkey")
	end if
	rsget.close

	getsendkey=tmpsendkey
end function

function getsendstatuscnt(sendstatus, orgitemid, orgitemoption, sendkey, isusing, orderserial, jukyogubun, usercell)
	dim query1, tmpcnt, i

	query1 = " select count(uidx) as cnt"
	query1 = query1 & " from db_item.[dbo].[tbl_item_standing_user] su"
	query1 = query1 & " where 1=1 "

	if sendstatus = "05" then
		query1 = query1 & " and su.sendstatus in (0,5)"
	end if
	if orgitemid <> "" then
		query1 = query1 & " and su.orgitemid="& orgitemid &""
	end if
	if orgitemoption <> "" then
		query1 = query1 & " and su.orgitemoption='"& orgitemoption &"'"
	end if
	if sendkey <> "" then
		query1 = query1 & " and su.sendkey="& sendkey &""
	end if
	if isusing <> "" then
		query1 = query1 & " and su.isusing='"& isusing &"'"
	end if
	if jukyogubun <> "" then
		query1 = query1 & " and su.jukyogubun='"& jukyogubun &"'"
	end if
	if jukyogubun="ORDER" then
		if orderserial <> "" then
			query1 = query1 & " and su.orderserial='"& orderserial &"'"
		end if
	else
		if usercell <> "" then
			query1 = query1 & " and su.usercell='"& usercell &"'"
		end if
	end if

	'response.write query1 & "<br>"
	rsget.Open query1,dbget,1
	if not rsget.EOF  then
		tmpcnt = rsget("cnt")
	end if
	rsget.close

	getsendstatuscnt=tmpcnt
end function

function getsendstatusname(vsendstatus)
	dim tmpsendstatus

	if vsendstatus="" or isnull(vsendstatus) then exit function

	if vsendstatus="0" then
		tmpsendstatus="발송대기"
	elseif vsendstatus="3" then
		tmpsendstatus="발송완료"
	elseif vsendstatus="5" then
		tmpsendstatus="재발송<br>대기"
	elseif vsendstatus="7" then
		tmpsendstatus="재발송<br>완료"
	end if

	getsendstatusname=tmpsendstatus
end function

function drawSelectBoxsendstatus(selectBoxName, selectedId, chplg)
	dim i

%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <%if selectedId="" then response.write " selected"%>>선택하세요</option>
		<option value='0' <%if selectedId="0" then response.write " selected"%>>발송대기</option>
		<option value='3' <%if selectedId="3" then response.write " selected"%>>발송완료</option>
		<option value='5' <%if selectedId="5" then response.write " selected"%>>재발송대기</option>
		<option value='7' <%if selectedId="7" then response.write " selected"%>>재발송완료</option>
		<option value='05' <%if selectedId="05" then response.write " selected"%>>발송대기,재발송대기</option>
		<option value='37' <%if selectedId="37" then response.write " selected"%>>발송완료,재발송완료</option>
	</select>
<%
end function

function getjukyoname(vjukyo)
	dim tmpjukyo

	if vjukyo="" or isnull(vjukyo) then exit function

	if vjukyo="ORDER" then
		tmpjukyo="주문"
	elseif vjukyo="GIFT" then
		tmpjukyo="서비스발송"
	else
		tmpjukyo=vjukyo
	end if

	getjukyoname=tmpjukyo
end function

function drawSelectBoxjukyo(selectBoxName, selectedId, chplg)
	dim i
%>
	<select class="select" name="<%=selectBoxName%>" <%=chplg%>>
		<option value='' <% if selectedId="" then response.write " selected" %>>선택하세요</option>
		<option value='ORDER' <% if selectedId="ORDER" then response.write " selected" %>>주문</option>
		<option value='EVENT' <% if selectedId="EVENT" then response.write " selected" %>>이벤트발송</option>
	</select>
<%
end function
%>