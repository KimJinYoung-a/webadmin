<%
'###########################################################
' Description : 컬러트랜드 관리
' Hieditor : 2012.03.29 한용민 생성
'###########################################################

class ccolortrend_item
	public fctcode
	public fcolorcode
	public fisusing
	public fstate
	public fmainimage
	public fmainimagelink
	public fmainimagelinknew
	public ftextimage
	public fstartdate
	public flastupdate
	public fregdate
	public flastadminid
	public fcolorName
	public fColorIcon
	public fstatename
	public Fidx
	public Fitemid
	public forderno
	public Fitemname
	public FImageSmall
	public Fsellyn
	public Flimityn
	public Flimitno
	public Flimitsold
	public fsortNo
	public fthisweek

	Public Fviewno
	Public Fpartwdid
	Public Fpartmdid
	Public Flistimg
	Public FNmainimg
	Public Fcolortitle

	Public FpartwdName
	Public FpartMdName
	
	public function IsSoldOut()
		IsSoldOut = (FSellyn="N") or (FSellyn="S") or ((FLimityn="Y") and (FLimitno-FLimitsold<1))
	end function
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

class ccolortrend_list
	Public FItemList()
	public foneitem
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	public FPageCount
	Public FCurrPage
	Public FPageSize
	public FTotalPage
	public frectctcode
	public frectcolorcode
	public frectstate
	public frectisusing
	public frectitemid
	public frectitemname

	Public frectviewno
	Public frectpartwdid
	Public frectpartmdid
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0

	End Sub
	Private Sub Class_Terminate()
	End Sub

	'/admin/itemmaster/colortrend_edit.asp
	Public Sub getcolortrend_one()
		dim sqlstr,i , sqlsearch
		
		if frectctcode <> "" then
			sqlsearch = sqlsearch & " and t.ctcode = "&frectctcode&""
		end if
		
		sqlstr = "select top 1"
		sqlStr = sqlStr & " t.ctcode ,t.colorCode ,t.isusing ,t.state ,t.mainimage ,t.mainimagelink ,t.mainimagelinknew ,t.textimage"
		sqlStr = sqlStr & " ,t.startdate ,t.lastupdate ,t.regdate ,t.lastadminid"		
		sqlStr = sqlStr & " ,t.viewno ,t.partwdid ,t.partmdid ,t.listimg , t.Nmainimg, t.colortitle"	
		sqlStr = sqlStr & " from db_item.dbo.tbl_colortrend t"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by t.ctcode desc"

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new ccolortrend_item
        
        if Not rsget.Eof then

			foneitem.fctcode = rsget("ctcode")
			foneitem.fcolorCode = rsget("colorCode")
			foneitem.fisusing = rsget("isusing")
			foneitem.fstate = rsget("state")
			foneitem.fstate = rsget("state")
			foneitem.fmainimage = rsget("mainimage")
			foneitem.fmainimagelink = db2html(rsget("mainimagelink"))
			foneitem.fmainimagelinknew = db2html(rsget("mainimagelinknew"))
			foneitem.ftextimage = rsget("textimage")
			foneitem.fstartdate = rsget("startdate")
			foneitem.flastupdate = rsget("lastupdate")
			foneitem.fregdate = rsget("regdate")
			foneitem.flastadminid = rsget("lastadminid")
			foneitem.Fviewno = rsget("viewno")
			foneitem.Fpartwdid = rsget("partwdid")
			foneitem.Fpartmdid = rsget("partmdid")
			foneitem.Flistimg = rsget("listimg")
			foneitem.FNmainimg = rsget("Nmainimg")
			foneitem.Fcolortitle = rsget("colortitle")

        end if
        rsget.Close
    end Sub
    
	'/admin/itemmaster/colortrend.asp
	public function getcolortrend()
        dim sqlStr, sqlsearch, i

		If frectviewno <> "" Then
			If frectviewno > 0 then
				sqlsearch = sqlsearch & " and t.viewno='"& frectviewno &"'"
			End If 
		End If 

		If frectpartwdid <> "" Then
			sqlsearch = sqlsearch & " and t.partwdid='"&frectpartwdid&"'"
		End If 

		If  frectpartmdid <> "" Then
			sqlsearch = sqlsearch & " and t.partmdid='"&frectpartmdid&"'"
		End If 

		if frectcolorcode <> "" then
			sqlsearch = sqlsearch & " and t.colorcode='"&frectcolorcode&"'"
		end if	
		if frectctcode <> "" then
			sqlsearch = sqlsearch & " and t.ctcode='"&frectctcode&"'"
		end if		
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and t.isusing='" + FRectIsUsing + "'"
        end if
		
		If frectstate <> "" THEN
			IF frectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and t.state = 7 and getdate() < t.startdate"
			ELSEIF frectstate = 7 THEN	'오픈
				sqlsearch  = sqlsearch & " and t.state = 7 and getdate() >= t.startdate"
			ELSE
				sqlsearch  = sqlsearch & " and  t.state = "&frectstate & ""
			END IF
		End If
        
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_item.dbo.tbl_colortrend t"
		sqlStr = sqlStr & " join [db_item].[dbo].tbl_colorChips c"
		sqlStr = sqlStr & " 	on t.colorcode = c.colorcode"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
					
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " t.ctcode ,t.colorCode ,t.isusing ,t.state ,t.mainimage ,t.mainimagelink ,t.textimage"
		sqlStr = sqlStr & " ,t.startdate ,t.lastupdate ,t.regdate ,t.lastadminid"		
		sqlStr = sqlStr & " ,c.colorName ,c.ColorIcon ,c.sortNo"
		sqlStr = sqlStr & " ,(case when t.state = 7 and getdate() < t.startdate then '6'" + vbcrlf
		sqlStr = sqlStr & " 	when t.state = 7 and getdate() >= t.startdate then '7'" + vbcrlf
		sqlStr = sqlStr & " 	else t.state end) as statename" + vbcrlf
		sqlstr = sqlstr & " ,(select top 1 ctcode"
		sqlstr = sqlstr & " 	from db_item.dbo.tbl_colortrend"
		sqlstr = sqlstr & " 	where state = 7"
		sqlstr = sqlstr & " 	and startdate <= getdate()"
		sqlstr = sqlstr & " 	and isusing='Y'"
		sqlstr = sqlstr & " 	order by ctcode desc"
		sqlstr = sqlstr & " ) as thisweek"
		sqlStr = sqlStr & " ,(SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = t.partMDid and statediv = 'Y') as partMDname "
		sqlStr = sqlStr & " ,(SELECT top 1 p.username from db_partner.dbo.tbl_user_tenbyten as p WHERE p.userid = t.partWDid and statediv = 'Y') as partWDname "
		sqlStr = sqlStr & " , t.viewno , t.colortitle"
		sqlStr = sqlStr & " from db_item.dbo.tbl_colortrend t"
		sqlStr = sqlStr & " join [db_item].[dbo].tbl_colorChips c"
		sqlStr = sqlStr & " 	on t.colorcode = c.colorcode"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by t.ctcode desc"
		
		'response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
        
        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new ccolortrend_item

				FItemList(i).fctcode = rsget("ctcode")
				FItemList(i).fcolorCode = rsget("colorCode")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fstate = rsget("state")
				FItemList(i).fstatename = rsget("statename")
				FItemList(i).fmainimage = rsget("mainimage")
				FItemList(i).fmainimagelink = db2html(rsget("mainimagelink"))
				FItemList(i).ftextimage = rsget("textimage")
				FItemList(i).fstartdate = rsget("startdate")
				FItemList(i).flastupdate = rsget("lastupdate")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastadminid = rsget("lastadminid")
				FItemList(i).fcolorName = db2html(rsget("colorName"))
				FItemList(i).FcolorIcon	= "http://fiximage.10x10.co.kr/web2012/colortrend/ico_color_"&Format00(2,rsget("colorCode"))&".gif"
				FItemList(i).fsortNo = rsget("sortNo")
				FItemList(i).fthisweek = rsget("thisweek")
				FItemList(i).FpartmdName = rsget("partMDname")
				FItemList(i).FpartwdName = rsget("partWDname")
				FItemList(i).Fviewno = rsget("viewno")
				FItemList(i).Fcolortitle = rsget("colortitle")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
	
	'//admin/itemmaster/colortrend_item.asp
	public Function getcolortrend_item()
		dim sqlStr ,sqlsearch , i

		if FRectcolorcode <> "" then
			sqlsearch = sqlsearch + " and ti.colorcode = " + FRectcolorcode + "" + vbcrlf
		end if
		if FRectIsUsing <> "" then
			sqlsearch = sqlsearch + " and ti.isusing = '" + FRectIsUsing + "'" + vbcrlf
		end if
		if FRectitemid <> "" then
			sqlsearch = sqlsearch + " and ti.itemid = " + FRectitemid + "" + vbcrlf
		end if
		if FRectitemname <> "" then
			sqlsearch = sqlsearch + " and i.itemname like '%"+ FRectitemname + "%'" + vbcrlf
		end if
				
		'총수 접수
		sqlStr = "select count(*), CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ")"
		sqlStr = sqlStr + " from db_item.dbo.tbl_colortrend_item ti"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_colorChips c"
		sqlStr = sqlStr + " 	on ti.colorcode = c.colorcode"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on ti.itemid = i.itemid"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		
		'response.write sqlStr &"<BR>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		if FTotalCount < 1 then exit Function
		
		'내용 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " ti.idx ,ti.colorCode ,ti.itemid ,ti.orderno ,ti.isusing ,c.colorIcon" + vbcrlf
		sqlStr = sqlStr + " ,i.itemname, i.smallimage ,i.sellyn ,i.limityn ,i.limitno ,i.limitsold" + vbcrlf
		sqlStr = sqlStr + " from db_item.dbo.tbl_colortrend_item ti"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_colorChips c"
		sqlStr = sqlStr + " 	on ti.colorcode = c.colorcode"
		sqlStr = sqlStr + " join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " 	on ti.itemid = i.itemid"
		sqlStr = sqlStr + " where 1=1 " & sqlsearch
		sqlStr = sqlStr + " order by ti.orderno asc, ti.idx desc"

		'response.write sqlStr &"<BR>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new ccolortrend_item

				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).fcolorCode			= rsget("colorCode")
				FItemList(i).fitemid			= rsget("itemid")
				FItemList(i).forderno			= rsget("orderno")
				FItemList(i).fisusing			= rsget("isusing")
				FItemList(i).FitemName		= db2html(rsget("itemname"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
				FItemList(i).fcolorIcon		= webImgUrl & "/color/colorchip/" & rsget("colorIcon")				
				FItemList(i).FSellyn		= rsget("sellyn")
				FItemList(i).Flimityn		= rsget("limityn")
				FItemList(i).Flimitno		= rsget("limitno")
				FItemList(i).Flimitsold		= rsget("limitsold")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
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
end class

'//상태값 공통함수 select박스
function Drawcolortrendstate(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option>
		<!--<option value="6" <% if selectedId="6" then response.write "selected" %>>오픈예정</option>-->
		<option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option>
		<!--<option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option>-->
	</select>
<%
end function

'//상태값 공통함수
function getcolortrendstate(v)
	if v = "0" then
		getcolortrendstate = "등록대기"
	elseif v = "3" then
		getcolortrendstate = "이미지등록요청"
	elseif v = "5" then
		getcolortrendstate = "오픈요청"
	elseif v = "6" then
		getcolortrendstate = "오픈예정"
	elseif v = "7" then
		getcolortrendstate = "오픈"
	'elseif v = "9" then
	'	getcolortrendstate = "종료"
	end if						
end Function

'/담당MD 리스트가져오기 (팀장 미만,직원 이상)
Sub sbGetpartid(ByVal selName, ByVal sIDValue, ByVal sScript,part_sn)
	Dim strSql, arrList, intLoop
	
	if part_sn = "" then exit sub
	
	strSql = " SELECT userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "   
	strSql = strSql & " WHERE part_sn IN("&part_sn&") and  posit_sn>='4' and  posit_sn<='12' and isUsing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & " and userid <> '' order by posit_sn, empno"
	
	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
	arrList = rsget.getRows()	
	End IF
	rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
	<%   
	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
	<option value="<%=arrList(0,intLoop)%>" <% if arrList(0,intLoop) = sIDValue then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
	<%   		
		Next
	End IF
	%>
	</select>
<%	
End Sub
%>