<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim iid,userid,chkcnt
iid	=request("iid")


sql = " select count(userid) as cnt from [db_contents].[dbo].[tbl_diary_event_evaluate] " &_
			" where userid =(select top 1 userid from [db_contents].[dbo].[tbl_diary_event_evaluate] where idx=" & iid & ")"

rsget.open sql,dbget,1

if not rsget.eof then
	chkcnt= rsget("cnt")
end if
rsget.close

if chkcnt<1 then
sql = " insert into [db_contents].[dbo].[tbl_diary_event_evaluate] " &_
			" (itemid ,itemname ,contents, userid ,point ,point1 ,point2 ,point3 ,point4  " &_
			" ,regdate ,linkimg1 ,linkimg2, smallimg)  " &_
			" SELECT top 1 u.itemid, i.itemname, u.contents ,u.userid  " &_
			" , IsNULL(u.tpoint,0) as point,  IsNULL(u.point1,0) as point_fun  " &_
			" , IsNULL(u.point2,0) as point_dgn, IsNULL(u.point3,0) as point_prc  " &_
			" , IsNULL(u.point4,0) as point_stf , convert(varchar(10),u.regdate,21) as regdate  " &_
			" , IsNULL(u.file1,'') as linkimg1, IsNULL(u.file2,'') as linkimg2, i.smallimage  " &_
			" FROM [db_board].[10x10].tbl_user_goodusing u  " &_
			" JOIN [db_item].[10x10].tbl_item i  " &_
			" ON i.itemid=u.itemid  " &_
			" WHERE u.id=" & iid

rsget.open sql,dbget,1

else
response.write "<script language='javascript' type='text/javascript'>alert('이미 등록되었습니다.');</script>"
end if

dbget.close()	:	response.End

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
