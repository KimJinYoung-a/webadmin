<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->

<%
dim mode, lecidx
mode	=	RequestCheckvar(request("mode"),16)
lecidx	=	RequestCheckvar(request("lecidx"),10)

dim sqlStr
dim imagecontent, imagecontent1, imagecontent2, imagecontent3, imagecontent4, imagecontent5

if mode="lecitem" then
''idx,cate_large,lec_date,lec_title,lecturer_id,lecturer_name
''lec_cost,buying_cost,mileage,mat_cost,matinclude_yn,limit_count,
''min_count,limit_sold,lec_startday1,lec_endday1,lec_startday2,lec_endday2,
''lec_startday3,lec_endday3,lec_startday4,lec_endday4,lec_startday5,lec_endday5,
''reg_startday,reg_endday,lec_count,lec_time,lec_period,lec_space,lec_outline,
''lec_contents,lec_etccontents,isusing,reg_yn,disp_yn,keyword,lec_mapimg,
''basicimg,icon1,listimg,icon2,smallimg,mainimg,storyimg,
''addimg1,addimg2,addimg3,addimg4,addimg5,
''addcontents1,addcontents2,addcontents3,addcontents4,addcontents5,regdate,


'''강좌정보이전

	sqlStr = "insert into [db_academy].[dbo].tbl_lec_item("
	sqlStr = sqlStr + " idx, cate_large, lec_date, lec_title, lecturer_id, lecturer_name"
	sqlStr = sqlStr + " , lec_cost, buying_cost, mileage,mat_cost, matinclude_yn, limit_count"
	sqlStr = sqlStr + " , min_count, limit_sold "
	sqlStr = sqlStr + " , lec_startday1, lec_endday1, lec_startday2, lec_endday2"
	sqlStr = sqlStr + " , lec_startday3, lec_endday3, lec_startday4, lec_endday4, lec_startday5, lec_endday5"
	sqlStr = sqlStr + " , reg_startday, reg_endday, lec_count, lec_time,lec_period, lec_space, lec_outline"
	sqlStr = sqlStr + " , lec_contents, lec_etccontents, isusing, reg_yn, disp_yn, keyword"
	sqlStr = sqlStr + " , basicimg , icon1,listimg,icon2,smallimg,mainimg,storyimg"
	sqlStr = sqlStr + " , regdate)"

	sqlStr = sqlStr + " select  idx, '10', mastercode, lectitle, lecturerid, lecturer"
	sqlStr = sqlStr + " ,lecsum, IsNULL(i.buycash,0), IsNULL(i.mileage,0),matsum, matinclude, le.properperson"
	sqlStr = sqlStr + " ,minperson, IsNULL(i.limitsold,0)"
	sqlStr = sqlStr + " ,lecdate01,lecdate01_end,lecdate02,lecdate02_end"
	sqlStr = sqlStr + " ,lecdate03,lecdate03_end,lecdate04,lecdate04_end,lecdate05,lecdate05_end"
	sqlStr = sqlStr + " ,reservestart, reserveend, leccount, lectime, lecperiod, lecspace, leccontents"
	sqlStr = sqlStr + " ,leccurry, lecetc, le.isusing, (case when regfinish='N' then 'Y' else 'N' end ), le.isusing, i.keywords"
	sqlStr = sqlStr + " ,i.basicimage, i.icon1image, i.listimage, i.icon2image, i.smallimage, i.mainimage, Left(i.storyimage,17)"
	sqlStr = sqlStr + " ,IsNULL(i.oregdate,getdate())"
	sqlStr = sqlStr + " from [db_contents].[dbo].tbl_lecture_item le"
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on le.linkitemid=i.itemid"
	sqlStr = sqlStr + " where le.idx=" + CStr(lecidx)

	rsget.Open sqlStr, dbget, 1

''이미지정보이전
	sqlStr = " update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set addimg1=T.addimage"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select L.idx,Left(i.addimage,17) as addimage"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)
	sqlStr = sqlStr + "	and Left(i.addimage,17)='A'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.idx"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set addimg2=T.addimage"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select L.idx,Right(Left(i.addimage,35),17) as addimage"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)
	sqlStr = sqlStr + "	and Left(Right(Left(i.addimage,35),17),1)='A'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.idx"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set addimg3=T.addimage"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select L.idx,Right(Left(i.addimage,53),17) as addimage"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)
	sqlStr = sqlStr + "	and Left(Right(Left(i.addimage,53),17),1)='A'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.idx"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set addimg4=T.addimage"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select L.idx,Right(Left(i.addimage,71),17) as addimage"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)
	sqlStr = sqlStr + "	and Left(Right(Left(i.addimage,71),17),1)='A'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.idx"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set addimg5=T.addimage"
	sqlStr = sqlStr + " from ("
	sqlStr = sqlStr + " select L.idx,Right(Left(i.addimage,89),17) as addimage"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)
	sqlStr = sqlStr + "	and Left(Right(Left(i.addimage,89),17),1)='A'"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where [db_academy].[dbo].tbl_lec_item.idx=T.idx"
	rsget.Open sqlStr, dbget, 1

''이미지설명
	sqlStr = "select L.idx, i.imagecontent"
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + "	[db_item].[dbo].tbl_item i,"
	sqlStr = sqlStr + "	[db_contents].[dbo].tbl_lecture_item L"
	sqlStr = sqlStr + "	where i.itemid=L.linkitemid"
	sqlStr = sqlStr + "	and L.idx=" + CStr(lecidx)

	rsget.Open sqlStr, dbget, 1
	if not rsget.Eof then
		imagecontent = db2html(rsget("imagecontent"))
	end if
	rsget.Close

	if (not IsNULL(imagecontent)) and (imagecontent<>"") then
		imagecontent = split(imagecontent,"|")
		if UBound(imagecontent)>0 then
			imagecontent1 = imagecontent(0)
		end if

		if UBound(imagecontent)>0 then
			imagecontent2 = imagecontent(1)
		end if

		if UBound(imagecontent)>0 then
			imagecontent3 = imagecontent(2)
		end if

		if UBound(imagecontent)>0 then
			imagecontent4 = imagecontent(3)
		end if

		if UBound(imagecontent)>0 then
			imagecontent5 = imagecontent(4)
		end if

		sqlStr = " update [db_academy].[dbo].tbl_lec_item"
		sqlStr = sqlStr + " set addcontents1='" + html2db(imagecontent1) + "'"
		sqlStr = sqlStr + " ,addcontents2='" + html2db(imagecontent2) + "'"
		sqlStr = sqlStr + " ,addcontents3='" + html2db(imagecontent3) + "'"
		sqlStr = sqlStr + " ,addcontents4='" + html2db(imagecontent4) + "'"
		sqlStr = sqlStr + " ,addcontents5='" + html2db(imagecontent5) + "'"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)

		rsget.Open sqlStr, dbget, 1

		''강좌일정
		''lecdate01,lecdate01_end, lecdate02,lecdate02_end
		''lecdate03,lecdate03_end, lecdate04,lecdate04_end
		''lecdate05,lecdate05_end

		sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
		sqlStr = sqlStr + " select idx,lecdate01,lecdate01_end"
		sqlStr = sqlStr + "  from [db_contents].[dbo].tbl_lecture_item i"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_schedule s "
		sqlStr = sqlStr + " on i.idx=s.lec_idx and lecdate01=startdate and lecdate01_end=enddate"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)
		sqlStr = sqlStr + " and lecdate01<>'1900-01-01'"
		sqlStr = sqlStr + " and s.lec_idx is null"

		rsget.Open sqlStr, dbget, 1


		sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
		sqlStr = sqlStr + " select idx,lecdate02,lecdate02_end"
		sqlStr = sqlStr + "  from [db_contents].[dbo].tbl_lecture_item i"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_schedule s "
		sqlStr = sqlStr + " on i.idx=s.lec_idx and lecdate02=startdate and lecdate02_end=enddate"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)
		sqlStr = sqlStr + " and lecdate02<>'1900-01-01'"
		sqlStr = sqlStr + " and s.lec_idx is null"

		rsget.Open sqlStr, dbget, 1


		sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
		sqlStr = sqlStr + " select idx,lecdate03,lecdate03_end"
		sqlStr = sqlStr + "  from [db_contents].[dbo].tbl_lecture_item i"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_schedule s "
		sqlStr = sqlStr + " on i.idx=s.lec_idx and lecdate03=startdate and lecdate03_end=enddate"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)
		sqlStr = sqlStr + " and lecdate03<>'1900-01-01'"
		sqlStr = sqlStr + " and s.lec_idx is null"

		rsget.Open sqlStr, dbget, 1


		sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
		sqlStr = sqlStr + " select idx,lecdate04,lecdate04_end"
		sqlStr = sqlStr + "  from [db_contents].[dbo].tbl_lecture_item i"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_schedule s "
		sqlStr = sqlStr + " on i.idx=s.lec_idx and lecdate04=startdate and lecdate04_end=enddate"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)
		sqlStr = sqlStr + " and lecdate04<>'1900-01-01'"
		sqlStr = sqlStr + " and s.lec_idx is null"

		rsget.Open sqlStr, dbget, 1

		sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
		sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
		sqlStr = sqlStr + " select idx,lecdate05,lecdate05_end"
		sqlStr = sqlStr + "  from [db_contents].[dbo].tbl_lecture_item i"
		sqlStr = sqlStr + " left join [db_academy].[dbo].tbl_lec_schedule s "
		sqlStr = sqlStr + " on i.idx=s.lec_idx and lecdate05=startdate and lecdate05_end=enddate"
		sqlStr = sqlStr + " where idx=" + CStr(lecidx)
		sqlStr = sqlStr + " and lecdate05<>'1900-01-01'"
		sqlStr = sqlStr + " and s.lec_idx is null"

		rsget.Open sqlStr, dbget, 1
	end if

	response.write "<script>alert('ok')</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End
elseif mode="lecitemdel" then
	sqlStr = "delete from [db_academy].[dbo].tbl_lec_item where idx=" + CStr(lecidx)
	rsget.Open sqlStr, dbget, 1

	response.write "<script>alert('ok')</script>"
	response.write "<script>window.close()</script>"
	dbget.close()	:	response.End

	sqlStr = "update [db_academy].[dbo].tbl_lec_item"
	sqlStr = sqlStr + " set lec_date=T.mastercode"
	sqlStr = sqlStr + " ,lec_title=T.lectitle"
	sqlStr = sqlStr + " ,lecturer_id=T.lecturerid"
	sqlStr = sqlStr + " ,lecturer_name=T.lecturer"
	sqlStr = sqlStr + " ,lec_cost=T.lecsum"
	sqlStr = sqlStr + " ,buying_cost=IsNULL(T.buycash,0)"
	sqlStr = sqlStr + " ,mileage=IsNULL(T.mileage,0)"
	sqlStr = sqlStr + " ,mat_cost=matsum"
	sqlStr = sqlStr + " ,matinclude_yn=matinclude"
	sqlStr = sqlStr + " ,limit_count=T.properperson"
	sqlStr = sqlStr + " ,min_count=minperson"
	sqlStr = sqlStr + " ,limit_sold=IsNULL(T.limitsold,0)"
	sqlStr = sqlStr + " ,lec_startday1=T.lecdate01"
	sqlStr = sqlStr + " ,lec_endday1=T.lecdate01_end"
	sqlStr = sqlStr + " ,lec_startday2=T.lecdate02"
	sqlStr = sqlStr + " ,lec_endday2=T.lecdate02_end"
	sqlStr = sqlStr + " ,lec_startday3=T.lecdate03"
	sqlStr = sqlStr + " ,lec_endday3=T.lecdate03_end"
	sqlStr = sqlStr + " ,lec_startday4=T.lecdate04"
	sqlStr = sqlStr + " ,lec_endday4=T.lecdate04_end"
	sqlStr = sqlStr + " ,lec_startday5=T.lecdate05"
	sqlStr = sqlStr + " ,lec_endday5=T.lecdate05_end"
	sqlStr = sqlStr + " ,reg_startday=T.reservestart"
	sqlStr = sqlStr + " ,reg_endday=T.reserveend"
	sqlStr = sqlStr + " ,lec_count=T.leccount"
	sqlStr = sqlStr + " ,lec_time=T.lectime"
	sqlStr = sqlStr + " ,lec_period=T.lecperiod"
	sqlStr = sqlStr + " ,lec_space=T.lecspace"
	sqlStr = sqlStr + " ,lec_outline=T.leccontents"

	sqlStr = sqlStr + " where le.idx=" + CStr(lecidx)



	sqlStr = sqlStr + " , , , , ,, , "
	sqlStr = sqlStr + " , lec_contents, lec_etccontents, isusing, reg_yn, disp_yn, keyword"
	sqlStr = sqlStr + " , basicimg , icon1,listimg,icon2,smallimg,mainimg,storyimg"
	sqlStr = sqlStr + " , regdate)"

	sqlStr = sqlStr + " ,, , , , , , "
	sqlStr = sqlStr + " ,leccurry, lecetc, le.isusing, (case when regfinish='N' then 'Y' else 'N' end ), le.isusing, i.keywords"
	sqlStr = sqlStr + " ,i.basicimage, i.icon1image, i.listimage, i.icon2image, i.smallimage, i.mainimage, Left(i.storyimage,17)"
	sqlStr = sqlStr + " ,IsNULL(i.oregdate,getdate())"
	sqlStr = sqlStr + " from [db_contents].[dbo].tbl_lecture_item le"
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on le.linkitemid=i.itemid"
	sqlStr = sqlStr + " where le.idx=" + CStr(lecidx)

	''rsget.Open sqlStr, dbget, 1
elseif  mode="lecschedule" then


	sqlStr= "delete from [db_academy].[dbo].tbl_lec_schedule"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
	sqlStr = sqlStr + " select idx,lec_startdate1,lec_enddate1"
	sqlStr = sqlStr + " where lec_startdate1<>'1900-01-01'"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
	sqlStr = sqlStr + " select idx,lec_startdate2,lec_enddate2"
	sqlStr = sqlStr + " where lec_startdate2<>'1900-01-01'"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
	sqlStr = sqlStr + " select idx,lec_startdate3,lec_enddate3"
	sqlStr = sqlStr + " where lec_startdate3<>'1900-01-01'"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
	sqlStr = sqlStr + " select idx,lec_startdate4,lec_enddate4"
	sqlStr = sqlStr + " where lec_startdate4<>'1900-01-01'"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " insert into [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " (lec_idx,startdate,enddate)"
	sqlStr = sqlStr + " select idx,lec_startdate5,lec_enddate5"
	sqlStr = sqlStr + " where lec_startdate5<>'1900-01-01'"
	rsget.Open sqlStr, dbget, 1

	''검토
	sqlStr = " select * from ("
	sqlStr = sqlStr + " select lec_idx,startdate,enddate,count(lec_idx) as cnt"
	sqlStr = sqlStr + "  from  [db_academy].[dbo].tbl_lec_schedule"
	sqlStr = sqlStr + " group by  lec_idx,startdate,enddate"
	sqlStr = sqlStr + " ) T"
	sqlStr = sqlStr + " where T.cnt>1"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->