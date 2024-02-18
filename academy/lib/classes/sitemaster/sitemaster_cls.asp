<%
'###########################################################
' Description : 코너관리
' History : 2009.09.15 한용민 생성
'###########################################################

Class cposcode_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	
	public fimagepath_etc
	public flinkpath
	public fevt_code
	public fregdate
	public fimagecount
	public fimage_order
	public fitemid
	public fgubun
	public frelation_itemcode
	public frelation_itemtitle
	public frelation_itemtitle2
	public frelation_itemcontents
	
	public fleftimagecolor
	public frightimagecolor
	public fsdate
	public fedate

	'xml 등록자,이미지 추적위해 추가 2017-03-16 유태욱
	public fxmluserid
	public fxmlregdate
	public fxmlimage
	public fdesigner
			    
end class

class cposcode_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	
	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectGubun
	public FRectIdx
	public frecttoplimit
	public FListIsUsing
	public FListGubun
	public FRectSearchSDate
	public FRectSearchEDate

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//academy/main/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_main_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where b.gubun='" & FRectGubun & "' " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf		
			end if	

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if

	        if FRectSearchSDate<>"" Then
	        	sqlStr = sqlStr & "  AND sdate >= '" & FRectSearchSDate & "'" & vbcrlf
	        end If
	        
			if FRectSearchEDate<>"" Then
	        	sqlStr = sqlStr & "  AND edate <= '" & FRectSearchEDate & "'" & vbcrlf
	        end If

	        if FRectvaliddate<>"" then
	        	sqlStr = sqlStr & "  AND (edate >= left(getdate(),10) ) " & vbcrlf
	        end if

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.regdate,a.poscode,a.isusing,a.image_order, a.leftimagecolor, a.relation_itemcontents, a.imagepath_etc, a.sdate, a.edate, a.xmluserid, a.xmlimage, a.xmlregdate " & vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_main_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where b.gubun='" & FRectGubun & "' " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf		
			end if	
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	

	        if FRectSearchSDate<>"" Then
	        	sqlStr = sqlStr & "  AND sdate >= '" & FRectSearchSDate & "'" & vbcrlf
	        end If
	        
			if FRectSearchEDate<>"" Then
	        	sqlStr = sqlStr & "  AND edate <= '" & FRectSearchEDate & "'" & vbcrlf
	        end If

	        if FRectvaliddate<>"" then
	        	sqlStr = sqlStr & "  AND (edate >= left(getdate(),10) ) " & vbcrlf
	        end if

		sqlStr = sqlStr & " order by a.idx Desc" + vbcrlf

'		response.write sqlStr &"<br>"
'		response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

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
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsACADEMYget("poscode")
				FItemList(i).fposname = db2html(rsACADEMYget("posname"))
				FItemList(i).fimagetype = rsACADEMYget("imagetype")
				FItemList(i).fimagewidth = rsACADEMYget("imagewidth")
				FItemList(i).fimageheight = rsACADEMYget("imageheight")
				FItemList(i).fisusing = rsACADEMYget("isusing")
				FItemList(i).fidx = rsACADEMYget("idx")
				FItemList(i).fimagepath = rsACADEMYget("imagepath")
				
				FItemList(i).fimagepath_etc = rsACADEMYget("imagepath_etc")
				
				FItemList(i).flinkpath = rsACADEMYget("linkpath")
				FItemList(i).fregdate = rsACADEMYget("regdate")		
				FItemList(i).fimagecount = rsACADEMYget("imagecount")
				FItemList(i).fimage_order = rsACADEMYget("image_order")
				FItemList(i).fleftimagecolor = rsACADEMYget("leftimagecolor")
				FItemList(i).frelation_itemcontents =  db2html(rsACADEMYget("relation_itemcontents"))

				FItemList(i).fsdate = rsACADEMYget("sdate")
				FItemList(i).fedate = rsACADEMYget("edate")
				
				FItemList(i).fxmluserid = rsACADEMYget("xmluserid")
				FItemList(i).fxmlimage = rsACADEMYget("xmlimage")
				FItemList(i).fxmlregdate = rsACADEMYget("xmlregdate")
				
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

'//academy/main/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.regdate,b.poscode,b.isusing,b.image_order" & vbcrlf		
		sqlStr = sqlStr & " ,b.relation_itemcode,b.relation_itemtitle,b.relation_itemtitle2,b.relation_itemcontents	" & vbcrlf			
		sqlStr = sqlStr & " ,b.leftimagecolor,b.rightimagecolor, b.imagepath_etc	" & vbcrlf			
		sqlStr = sqlStr & " ,b.sdate, b.edate, b.designer	" & vbcrlf			
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_academy.dbo.tbl_main_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        
        if Not rsACADEMYget.Eof then
    
			FOneItem.fposcode = rsACADEMYget("poscode")
			FOneItem.fposname = db2html(rsACADEMYget("posname"))
			FOneItem.fimagetype = rsACADEMYget("imagetype")
			FOneItem.fimagewidth = rsACADEMYget("imagewidth")
			FOneItem.fimageheight = rsACADEMYget("imageheight")
			FOneItem.fisusing = rsACADEMYget("isusing")
			FOneItem.fidx = rsACADEMYget("idx")
			FOneItem.fimagepath = db2html(rsACADEMYget("imagepath"))
			
			FOneItem.fimagepath_etc = db2html(rsACADEMYget("imagepath_etc"))
			
			FOneItem.flinkpath = db2html(rsACADEMYget("linkpath"))
			FOneItem.fregdate = rsACADEMYget("regdate")
			FOneItem.fimagecount = rsACADEMYget("imagecount") 
			FOneItem.fimage_order = rsACADEMYget("image_order")
			FOneItem.frelation_itemcode = rsACADEMYget("relation_itemcode")
			FOneItem.frelation_itemtitle =  db2html(rsACADEMYget("relation_itemtitle"))
			FOneItem.frelation_itemtitle2 =  db2html(rsACADEMYget("relation_itemtitle2"))
			FOneItem.frelation_itemcontents =  db2html(rsACADEMYget("relation_itemcontents"))

			FOneItem.fleftimagecolor = rsACADEMYget("leftimagecolor")
			FOneItem.frightimagecolor = rsACADEMYget("rightimagecolor")
			FOneItem.fsdate = rsACADEMYget("sdate")
			FOneItem.fedate = rsACADEMYget("edate")
			FOneItem.fdesigner = rsACADEMYget("designer")
        end if
        rsACADEMYget.Close
    end Sub
	
	'////academy/main/imagemake_poscode.asp
    public Sub fposcode_oneitem()		
        dim SqlStr
        SqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount, gubun" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)
         
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new cposcode_oneitem
        if Not rsACADEMYget.Eof then
            
            FOneItem.fposcode = rsACADEMYget("poscode")
            FOneItem.fposname = db2html(rsACADEMYget("posname"))
            FOneItem.fimagetype	= rsACADEMYget("imagetype")
            FOneItem.fimagewidth = rsACADEMYget("imagewidth")
            FOneItem.fimageheight = rsACADEMYget("imageheight")
            FOneItem.fisusing = rsACADEMYget("isusing")
            FOneItem.fimagecount = rsACADEMYget("imagecount")
            FOneItem.fgubun = rsACADEMYget("gubun")
                       
        end if
        rsACADEMYget.close
    end Sub

	'///academy/main/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,addSql, i

		if FListIsUsing="Y" or FListIsUsing="N" then
			addSql = " and isusing='" & FListIsUsing & "'"
		end if
		if FListGubun<>"" then
			addSql = addSql & " and gubun='" & FListGubun & "'"
		end if

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" & addSql + vbcrlf

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount, gubun" + vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_main_poscode" + vbcrlf			
		sqlStr = sqlStr & " where 1=1" & addSql + vbcrlf

		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

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
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new cposcode_oneitem
				
				FItemList(i).fposcode = rsACADEMYget("poscode")
				FItemList(i).fposname = db2html(rsACADEMYget("posname"))
				FItemList(i).fimagetype = rsACADEMYget("imagetype")
				FItemList(i).fimagewidth = rsACADEMYget("imagewidth")
				FItemList(i).fimageheight = rsACADEMYget("imageheight")
				FItemList(i).fisusing = rsACADEMYget("isusing")
				FItemList(i).fimagecount = rsACADEMYget("imagecount")
				FItemList(i).fgubun = rsACADEMYget("gubun")
														
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub
		
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

'//적용구분 
function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag,gubun)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_academy.dbo.tbl_main_poscode where isusing='Y' and gubun = '" & gubun & "' order by poscode desc"
   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("poscode")&"' "&tmp_str&">" + db2html(rsACADEMYget("posname")) + "</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
end function

function DrawGroupGubunCombo(selectBoxName,selectedId,changeFlag)
    dim bufStr, tmp_str
    bufStr = "<select class='select' name='gubun' " + changeFlag + ">" + VbCrlf

    if selectedId="index" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='index' " + tmp_str + " >index" + VbCrlf
    if selectedId="index_diy" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='index_diy' " + tmp_str + " >index DIY샾" + VbCrlf
    if selectedId="index_good" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='index_good' " + tmp_str + " >index 좋은강사" + VbCrlf
    if selectedId="index_event" then tmp_str = "selected" else tmp_str = ""
        bufStr = bufStr + " <option value='index_event' " + tmp_str + " >index 이벤트" + VbCrlf        
    bufStr = bufStr + " </select>" + VbCrlf
    
	response.write bufStr
end function

public Sub SelectLecturerId(byval lecturer_id)
	dim sqlStr,i
	sqlStr = "select  c.userid,p.company_name,c.socname, c.socname_kor"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
	sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
	sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
	sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
	sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf

	rsget.open sqlStr,dbget,1

	if not rsget.eof then
			response.write "<select name='temp_lec_id' onchange='javascript:FnLecturerApp(this.value);'>"
			response.write "<option value=''>선택</option>"
		for i=0 to rsget.recordcount-1
			if lecturer_id=db2html(rsget("userid")) then
			response.write "<option value='" & db2html(rsget("userid")) & "," & db2html(rsget("company_name")) & "," & rsget("socname") & "," & left(rsget("socname_kor"),10) & "' selected>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			else
			response.write "<option value='" & db2html(rsget("userid")) & "," & db2html(rsget("company_name")) & "," & rsget("socname") & "," & left(rsget("socname_kor"),10) & "'>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
			end if
		rsget.movenext
		next
			response.write "</select>"
	end if
	rsget.close

end sub
%>