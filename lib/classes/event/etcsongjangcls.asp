<%
'####################################################
' Description : 기타송장
' History : 서동석 생성
'			2022.10.06 한용민 수정(쿼리튜닝, 오류수정)
'####################################################

'' 테이블 변경 [db_contents].[dbo].tbl_etc_songjang -> [db_sitemaster].[dbo].tbl_etc_songjang
'' 2008.04.30 이벤트상품 테이블 조인, 이벤트코드 추가,  사은품명 추가, 검색에 배송추가

Class CEventsBeasongItem
	public Fid
	public Fgubuncd
	public Fgubunname
	public Fuserid
	public Fusername
	public Freqname
	public Freqphone
	public Freqhp
	public Freqzipcode
	public Freqaddress1
	public Freqaddress2
	public Freqetc
	public Fregdate
	public Fsongjangno
	public Fsongjangdiv
	public Fsenddate
	public Fissended
	public Finputdate
	public FPrizeTitle

    public FisUpchebeasong
    public Fdelivermakerid
    public Fevtprize_code
    public FreqDeliverDate
    public Fdeleteyn
    public Fjungsan
    public FjungsanYN
    public Fdivname
    public Fevtprize_giftkindcode
    public Fgiftkind_name
    public Fevtcode
    public Fevtgroupcode
    public Fgetcode
    public Fgift_code
    public Fgift_itemid

    public FetcKey
    public FetcBaljuNo

	public Fevtprize_enddate

    Public Function getEventKind()
		If Fevtcode = "1" AND Fevtcode <> "" Then
			getEventKind = "디자인핑거스"
			Fgetcode = "<a href='"&wwwUrl&"/designfingers/designfingers.asp?fingerid="&Fevtgroupcode&"' target='_blank'>"&Fevtgroupcode&"</a>"
		ElseIf Fgubuncd ="90" Then
			getEventKind = "반품"
'		ElseIf Fgubuncd ="96" Then
'			getEventKind = "고객"
'		ElseIf Fgubuncd ="97" Then
'			getEventKind = "29cm용"
		ElseIf Fgubuncd ="98" Then
			getEventKind = "판촉"
		ElseIf Fgubuncd ="99" Then
			getEventKind = "기타"
		ElseIf Fgubuncd ="80" Then
		    getEventKind = "CS출고"
		ElseIf Fgubuncd ="70" Then
		    getEventKind = "매장출고"
		ElseIf Fevtcode ="4" AND Fevtcode <> "" Then
			getEventKind = "컬쳐스테이션"
			Fgetcode = "<a href='"&wwwUrl&"/culturestation/culturestation_event.asp?evt_code="&Fevtgroupcode&"' target='_blank'>"&Fevtgroupcode&"</a>"
		Else
			getEventKind = "이벤트"
			Fgetcode = "<a href='"&wwwUrl&"/event/eventmain.asp?eventid="&Fevtcode&"' target='_blank'>"&Fevtcode&"</a>"
		End If
	End Function

    public function getPrizeTitle()
        if IsNULL(Fevtprize_giftkindcode) or (Fevtprize_giftkindcode=0) then
            getPrizeTitle = FPrizeTitle
        else
            getPrizeTitle = Fgiftkind_name
        end if
    end function

	public function IsInputData()
		if IsNULL(Finputdate) or (Finputdate="") then
			IsInputData = false
		else
			IsInputData = true
		end if
	end function

	public function IsSended()
		if (Fissended="Y") then
			IsSended = true
		else
			IsSended = false
		end if
	end function

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CEventsBeasong
	public FOneItem
	public FItemList()

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

	public FRectUserID
	public FRectId
    public FRectOnlySongjangNotInput
    public FRectSearchType
    public FRectSearchValue
    public FRectDeleteyn
    public FRectGubuncd
    public FRectIsupchebeasong
    public FRectDeliverMakerid
    public FRectDeliverAreaInputedOnly
	public FRectinputdatetype
    public FRectOnlyMisend
    public FRectJungsanYN
	public FRectIsFinish
	public FRectIsInput
    public FRectDateGubun
    public FRectStartdate
    public FRectEndDate

	Private Sub Class_Initialize()
        redim preserve FItemList(0)
        FCurrPage =1

		FPageSize = 10
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Sub GetOneWinnerItem()
		dim sqlStr , i

		if FRectId="" then exit Sub

		sqlStr = "select top 1 w.*, k.giftkind_name,k.itemid as gift_itemid,g.gift_code from [db_sitemaster].[dbo].tbl_etc_songjang as w "
		sqlStr = sqlStr + "     left join db_event.dbo.tbl_giftkind k on w.evtprize_giftkindcode=k.giftkind_code "
		sqlStr = sqlStr + "		left join db_event.dbo.tbl_event_prize p on w.evtprize_code = p.evtprize_code" + vbCrlf
		sqlStr = sqlStr + "     left join db_event.dbo.tbl_gift g on p.evt_code=g.evt_code"
		sqlStr = sqlStr + " where  id=" + CStr(FRectId) + ""

        if FRectDeleteyn<>"" then
            sqlStr = sqlStr + " and deleteyn='" + FRectDeleteyn + "'"
        end if

        if (FRectDeliverMakerid<>"") then
            sqlStr = sqlStr + " and delivermakerid='" + FRectDeliverMakerid + "'"
        end if
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.RecordCount

		if Not rsget.Eof then
			set FOneItem = new CEventsBeasongItem

			FOneItem.Fid           = rsget("id")
			FOneItem.Fgubuncd      = rsget("gubuncd")
			FOneItem.Fgubunname    = db2html(rsget("gubunname"))
			FOneItem.Fuserid       = rsget("userid")
			FOneItem.Fusername     = db2html(rsget("username"))
			FOneItem.Freqname      = db2html(rsget("reqname"))
			FOneItem.Freqphone     = rsget("reqphone")
			FOneItem.Freqhp        = rsget("reqhp")
			FOneItem.Freqzipcode   = rsget("reqzipcode")
			FOneItem.Freqaddress1  = db2html(rsget("reqaddress1"))
			FOneItem.Freqaddress2  = db2html(rsget("reqaddress2"))
			FOneItem.Freqetc       = db2html(rsget("reqetc"))
			FOneItem.Fregdate      = rsget("regdate")
			FOneItem.Fsongjangno   = rsget("songjangno")
			FOneItem.Fsongjangdiv  = rsget("songjangdiv")
			FOneItem.Fsenddate     = rsget("senddate")
			FOneItem.Fissended     = rsget("issended")
			FOneItem.Finputdate	   = rsget("inputdate")
			FOneItem.FPrizeTitle	= db2html(rsget("prizetitle"))
			FOneItem.Fevtprize_code = rsget("evtprize_code")

			FOneItem.FisUpchebeasong    = rsget("isupchebeasong")
			FOneItem.FDeliverMakerid    = rsget("delivermakerid")
			FOneItem.FreqDeliverDate    = rsget("reqdeliverdate")
			FOneItem.Fdeleteyn      	= rsget("deleteyn")
			FOneItem.Fevtprize_giftkindcode	= rsget("evtprize_giftkindcode")
			FOneItem.Fgiftkind_name = db2html(rsget("giftkind_name"))
			FOneItem.FjungsanYN    	= rsget("jungsanYN")
			FOneItem.Fjungsan      	= rsget("jungsan")
			FOneItem.Fgift_code     = rsget("gift_code")
			FOneItem.Fgift_itemid     = rsget("gift_itemid")
		end if
		rsget.close

	end Sub

    public Sub getSVCSongjangList(yyyymmdd)
        dim sqlStr
        sqlStr = "select top " + CStr(FPageSize*FCurrpage) + "" & vbCrlf
		sqlStr = sqlStr + " w.etcBaljuNo, w.songjangno, w.etcKey " + vbCrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_etc_songjang w" + vbCrlf
		sqlStr = sqlStr + " where w.deleteyn='N'" + vbCrlf
		sqlStr = sqlStr + " and w.regdate>='"&yyyymmdd&"'"
		sqlStr = sqlStr + " and w.regdate<'"&DateAdd("d",1,yyyymmdd)&"'"
		sqlStr = sqlStr + " and w.gubuncd='91'"
		sqlStr = sqlStr + " order by w.id "

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
    			set FItemList(i) = new CEventsBeasongItem
    			FItemList(i).FetcBaljuNo    = rsget("etcBaljuNo")
    			FItemList(i).Fsongjangno    = rsget("songjangno")
    			FItemList(i).FetcKey        = rsget("etcKey")
    			i=i+1
    			rsget.MoveNext
    		loop
		end if
		rsget.close

    end Sub

	'//admin/etcsongjang/eventsongjang.asp
    public sub getEventBeasongInfoList()
		dim sqlStr,i,sqlStr2, maxid, sqlsearch

		if FRectinputdatetype<>"" then
			if FRectinputdatetype="3MBEFORE" then
				sqlsearch = sqlsearch + " and not( isnull(w.inputdate,'')='' and w.regdate < dateadd(m,-3,getdate()) )"
			end if
		end if

		if (FRectDeliverMakerid<>"") then
		    sqlsearch = sqlsearch + " and datediff(d,reqdeliverdate,getdate())>=-2 " + vbCrlf
		end if

		'sqlsearch = sqlsearch + " and ((inputdate is Not Null) or ((inputdate is Null) and (datediff(d,regdate,getdate())<31)))"

		if (FRectDeliverAreaInputedOnly<>"") then
            sqlsearch = sqlsearch + " and w.inputdate is Not NULL" + vbCrlf
        end if

		if (FRectDeliverMakerid<>"") then
            sqlsearch = sqlsearch + " and w.delivermakerid='" + FRectDeliverMakerid + "'"
        end if

		if FRectDeleteyn<>"" then
            sqlsearch = sqlsearch + " and w.deleteyn='" + FRectDeleteyn + "'"
        end if

        if (FRectJungsanYN<>"") then
            if FRectJungsanYN="Y" then
                sqlsearch = sqlsearch + " and (w.jungsanyn='Y' or jungsan<>0)"
            else
                sqlsearch = sqlsearch + " and isNULL(w.jungsanyn,'N')='" + FRectJungsanYN + "'"
            end if
        end if

        If FRectGubuncd <> "" then
        	Select Case FRectGubuncd
        		Case "90","99","98","70"
        		'Case "90","99","98","96","97"
        			sqlsearch = sqlsearch + " and w.gubuncd='"+FRectGubuncd+"'"
        		Case "ev"
        			sqlsearch = sqlsearch + " and w.gubuncd = '01' and p.evt_code not in ('1','4') "
        		Case "1","4"
        			sqlsearch = sqlsearch + " and w.gubuncd = '01' and p.evt_code = '"+FRectGubuncd+"' "
        	End Select
        End If

		if FRectOnlySongjangNotInput <> "" then
			sqlsearch = sqlsearch + " and ((w.songjangno is NULL) or (w.songjangno=''))"
		end if

        if (FRectOnlyMisend<>"") then
            sqlsearch = sqlsearch + " and w.senddate is NULL"
            sqlsearch = sqlsearch + " and w.regdate>'2008-01-01'"
        end if

        Select Case FRectIsFinish
        	Case "Y"
				sqlsearch = sqlsearch + " and w.senddate is Not NULL"
				sqlsearch = sqlsearch + " and w.regdate>'2008-01-01'"
			Case "N"
				sqlsearch = sqlsearch + " and w.senddate is NULL"
				sqlsearch = sqlsearch + " and w.regdate>'2008-01-01'"
        end Select

        Select Case FRectIsInput
        	Case "Y"
				sqlsearch = sqlsearch + " and w.inputdate is Not Null"
			Case "N"
				sqlsearch = sqlsearch + " and w.inputdate is Null"
        end Select

        if FRectIsupchebeasong="Y" then
            sqlsearch = sqlsearch + " and w.isupchebeasong='Y'"
        elseif FRectIsupchebeasong="N" then
            sqlsearch = sqlsearch + " and ((w.isupchebeasong='N') or (w.isupchebeasong is NULL))"
        end if

		if FRectSearchValue<>"" then
			if FRectSearchType="gubun" then
				sqlsearch = sqlsearch + " and w.gubunname like '%" + FRectSearchValue + "%'"
			elseif FRectSearchType="userid" then
				sqlsearch = sqlsearch + " and w.userid='" + FRectSearchValue + "'"
			elseif FRectSearchType="id" then
				sqlsearch = sqlsearch + " and w.id=" + FRectSearchValue + ""
			elseif FRectSearchType="username" then
				sqlsearch = sqlsearch + " and w.username like '" + FRectSearchValue + "%'"
			elseif FRectSearchType="reqname" then
				sqlsearch = sqlsearch + " and w.reqname like '" + FRectSearchValue + "%'"
			elseif 	FRectSearchType="eCode" then
				sqlsearch = sqlsearch + " and ( p.evt_code = " + FRectSearchValue + " or  p.evtgroup_code = " + FRectSearchValue + " )"
			elseif 	FRectSearchType="dlvMkrid" then
				sqlsearch = sqlsearch + " and w.delivermakerid='" + FRectSearchValue + "'"
			elseif 	FRectSearchType="songjangno" then
				sqlsearch = sqlsearch + " and w.songjangno='" + FRectSearchValue + "'"
				sqlsearch = sqlsearch + " and w.id >= '" & maxid & "'"
			end if
		end if

        if (FRectStartdate <> "") and (FRectEndDate <> "") then
            select case FRectDateGubun
                case "reqDeliverDate"
                    sqlsearch = sqlsearch + " and w.reqdeliverdate >= '" & FRectStartdate & "' "
                    sqlsearch = sqlsearch + " and w.reqdeliverdate < '" & FRectEndDate & "' "
                case "senddate"
                    sqlsearch = sqlsearch + " and w.senddate >= '" & FRectStartdate & "' "
                    sqlsearch = sqlsearch + " and w.senddate < '" & FRectEndDate & "' "
                case else
                    '
            end select
        end if

		sqlStr = "select count(*) as cnt " + vbCrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_etc_songjang w with (nolock)" + vbCrlf
		sqlStr = sqlStr + "		left join db_event.dbo.tbl_event_prize p with (nolock) on w.evtprize_code = p.evtprize_code" + vbCrlf
		sqlStr = sqlStr + " where w.id<>0 " & sqlsearch

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrpage) + "" & vbCrlf
		sqlStr = sqlStr + " w.id, w.userid, w.username, w.reqname" + vbCrlf
		sqlStr = sqlStr + " ,w.songjangno, w.regdate, w.senddate, w.inputdate, w.gubunname, w.prizetitle, w.issended" + vbCrlf
		sqlStr = sqlStr + " ,w.evtprize_code, w.isupchebeasong, w.delivermakerid, w.reqdeliverdate, w.songjangdiv, w.deleteyn" + vbCrlf
		sqlStr = sqlStr + " ,w.evtprize_giftkindcode, k.giftkind_name,p.evt_code, p.evtgroup_code, w.jungsan, w.jungsanYN, w.gubuncd, p.evtprize_enddate " + vbCrlf
		sqlStr = sqlStr + " ,(select top 1 divname from db_order.[dbo].tbl_songjang_div as v where w.songjangdiv = v.divcd) as divname " + vbCrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_etc_songjang w with (nolock)" + vbCrlf
		sqlStr = sqlStr + "		left join db_event.dbo.tbl_event_prize p with (nolock) on w.evtprize_code = p.evtprize_code" + vbCrlf
		sqlStr = sqlStr + "     left join db_event.dbo.tbl_giftkind k with (nolock) on w.evtprize_giftkindcode=k.giftkind_code" + vbCrlf
		sqlStr = sqlStr + " where w.id<>0 " & sqlsearch
		sqlStr = sqlStr + " order by w.id desc" + vbCrlf

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

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
    			set FItemList(i) = new CEventsBeasongItem
    			FItemList(i).Fid  			= rsget("id")
    			FItemList(i).Fuserid      	= rsget("userid")
    			FItemList(i).FuserName   	= rsget("username")
    			FItemList(i).FreqName   	= rsget("reqname")
    			FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
    			FItemList(i).Fsongjangno    = rsget("songjangno")
    			FItemList(i).Fsenddate      = rsget("senddate")
    			FItemList(i).Fregdate       = rsget("regdate")
    			FItemList(i).Finputdate     = rsget("inputdate")
    			FItemList(i).Fgubunname		= db2html(rsget("gubunname"))
    			FItemList(i).Fprizetitle	= db2html(rsget("prizetitle"))
    			FItemList(i).Fissended      = rsget("issended")

    			FItemList(i).Fisupchebeasong= rsget("isupchebeasong")
    			FItemList(i).Fdelivermakerid= rsget("delivermakerid")
    			FItemList(i).Fevtprize_code = rsget("evtprize_code")

    			FItemList(i).FreqDeliverDate    = rsget("reqdeliverdate")
    			FItemList(i).Fdeleteyn      = rsget("deleteyn")

    			FItemList(i).Fevtprize_giftkindcode  = rsget("evtprize_giftkindcode")
    			FItemList(i).Fgiftkind_name = db2html(rsget("giftkind_name"))
    			FItemList(i).Fevtcode	= rsget("evt_code")
    			FItemList(i).Fevtgroupcode	= rsget("evtgroup_code")
    			FItemList(i).Fjungsan	= rsget("jungsan")
    			FItemList(i).FjungsanYN	= rsget("jungsanYN")
    			FItemList(i).Fdivname	= rsget("divname")
    			FItemList(i).Fgubuncd	= rsget("gubuncd")

				FItemList(i).Fevtprize_enddate	= rsget("evtprize_enddate")


    			i=i+1
    			rsget.MoveNext
    		loop
		end if
		rsget.close
	end sub

	public Sub GetWinnerList()
		dim sqlStr , i

		sqlStr = "select count(id) as cnt from [db_sitemaster].[dbo].tbl_etc_songjang"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"

        if FRectDeleteyn<>"" then
            sqlStr = sqlStr + " and w.deleteyn='" + FRectDeleteyn + "'"
        end if

		''2주안에 입력안하면 보이지 않게 함..
		'sqlStr = sqlStr + " and ((inputdate is Not Null) or ((inputdate is Null) and (datediff(d,regdate,getdate())<31)))"
		if FRectId<>"" then
			sqlStr = sqlStr + " and id=" + FRectId + ""
		end if

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		sqlStr = "select * from [db_sitemaster].[dbo].tbl_etc_songjang"
		sqlStr = sqlStr + " where userid='" + FRectUserID + "'"
		if FRectDeleteyn<>"" then
            sqlStr = sqlStr + " and w.deleteyn='" + FRectDeleteyn + "'"
        end if

		'sqlStr = sqlStr + " and ((inputdate is Not Null) or ((inputdate is Null) and (datediff(d,regdate,getdate())<31)))"
		if FRectId<>"" then
			sqlStr = sqlStr + " and id=" + FRectId + ""
		end if
		sqlStr = sqlStr + " order by id desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CEventsBeasongItem

				FItemList(i).Fid           = rsget("id")
				FItemList(i).Fgubuncd      = rsget("gubuncd")
				FItemList(i).Fgubunname    = db2html(rsget("gubunname"))
				FItemList(i).Fuserid       = rsget("userid")
				FItemList(i).Fusername     = db2html(rsget("username"))
				FItemList(i).Freqname      = db2html(rsget("reqname"))
				FItemList(i).Freqphone     = rsget("reqphone")
				FItemList(i).Freqhp        = rsget("reqhp")
				FItemList(i).Freqzipcode   = rsget("reqzipcode")
				FItemList(i).Freqaddress1  = db2html(rsget("reqaddress1"))
				FItemList(i).Freqaddress2  = db2html(rsget("reqaddress2"))
				FItemList(i).Freqetc       = db2html(rsget("reqetc"))
				FItemList(i).Fregdate      = rsget("regdate")
				FItemList(i).Fsongjangno   = rsget("songjangno")
				FItemList(i).Fsongjangdiv  = rsget("songjangdiv")
				FItemList(i).Fsenddate     = rsget("senddate")
				FItemList(i).Fissended     = rsget("issended")
				FItemList(i).Finputdate	   = rsget("inputdate")
				FItemList(i).FPrizeTitle	= db2html(rsget("prizetitle"))

				FItemList(i).FreqDeliverDate    = rsget("reqdeliverdate")
				FItemList(i).Fdeleteyn      = rsget("deleteyn")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end Sub

	' /admin/etcsongjang/popsongjangmaker_event.asp
    public sub getEventSongJangList(byval idarr)
		dim sqlStr, i

		if idarr="" or isnull(idarr) then exit sub

		sqlStr = "select top 1000" & vbCrlf
		sqlStr = sqlStr + " w.id, w.userid, w.username, w.reqname, w.reqphone, w.reqhp, w.reqzipcode, w.reqaddress1," + vbCrlf
		sqlStr = sqlStr + " w.reqaddress2, w.songjangno, w.reqetc, w.regdate, w.senddate, w.gubunname, w.prizetitle" + vbCrlf
		sqlStr = sqlStr + " ,w.evtprize_giftkindcode, k.giftkind_name" + vbCrlf
		sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_etc_songjang w with (nolock)"
		sqlStr = sqlStr & " left join db_event.dbo.tbl_giftkind k with (nolock)"
		sqlStr = sqlStr & "		on w.evtprize_giftkindcode=k.giftkind_code"
		sqlStr = sqlStr + " where w.id in (" + Cstr(idarr) + ")"
		if FRectDeleteyn<>"" then
            sqlStr = sqlStr + " and w.deleteyn='" + FRectDeleteyn + "'"
        end if

        if FRectDeliverMakerid<>"" then
            sqlStr = sqlStr + " and w.delivermakerid='" + FRectDeliverMakerid + "'"
        end if

		sqlStr = sqlStr + " order by w.id desc" + vbCrlf

		'response.write sqlStr & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.RecordCount
		FResultCount = rsget.RecordCount
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF then
		do until rsget.EOF
			set FItemList(i) = new CEventsBeasongItem
			FItemList(i).Fid  			= rsget("id")
			FItemList(i).Fuserid      	= rsget("userid")
			FItemList(i).FuserName   	= db2html(rsget("username"))
			FItemList(i).FreqName   	= db2html(rsget("reqname"))
			FItemList(i).Freqphone   	= rsget("reqphone")
			FItemList(i).Freqhp   		= rsget("reqhp")
			FItemList(i).Freqzipcode   	= rsget("reqzipcode")
			FItemList(i).Freqaddress1   = db2html(rsget("reqaddress1"))
			FItemList(i).Freqaddress2   = db2html(rsget("reqaddress2"))
			FItemList(i).Freqetc   		= db2html(rsget("reqetc"))
			FItemList(i).Fsongjangno    = rsget("songjangno")
			FItemList(i).Fsenddate      = rsget("senddate")
			FItemList(i).Fregdate       = rsget("regdate")
			FItemList(i).Fgubunname	    = db2html(rsget("gubunname"))
			FItemList(i).Fprizetitle	= db2html(rsget("prizetitle"))
			FItemList(i).Fevtprize_giftkindcode  = rsget("evtprize_giftkindcode")
    		FItemList(i).Fgiftkind_name = db2html(rsget("giftkind_name"))

			i=i+1
			rsget.MoveNext
		loop
		end if
		rsget.close
	end sub

    public sub getMomoBeasongInfoList()
		dim sqlStr,i, vSubQuery

		vSubQuery = ""
		if FRectSearchType = "userid" then
			vSubQuery = vSubQuery & " and O.userid = '" + FRectSearchValue + "'"
		elseif FRectSearchType = "id" then
			vSubQuery = vSubQuery & " and O.orderid = '" + FRectSearchValue + "'"
		elseif FRectSearchType = "username" then
			vSubQuery = vSubQuery & " and O.username like '" + FRectSearchValue + "%'"
		elseif FRectSearchType = "reqname" then
			vSubQuery = vSubQuery & " and O.reqname like '" + FRectSearchValue + "%'"
		end if

		sqlStr = "select count(*) as cnt " + vbCrlf
		sqlStr = sqlStr + " from [db_momo].[dbo].tbl_momo_order AS O " + vbCrlf
		sqlStr = sqlStr + " where O.outputdate is Not Null " + vbCrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.close

		sqlStr = "select top " + CStr(FPageSize*FCurrpage) + "" & vbCrlf
		sqlStr = sqlStr + " 	O.orderid, O.userid, O.username, O.reqname, O.songjangno, O.outputdate, O.orderdate, O.itemname " + vbCrlf
		sqlStr = sqlStr + " from [db_momo].[dbo].tbl_momo_order AS O " + vbCrlf
		sqlStr = sqlStr + " where O.outputdate is Not Null " + vbCrlf
		sqlStr = sqlStr + " order by O.orderid desc" + vbCrlf

		'response.write	sqlStr
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
    			set FItemList(i) = new CEventsBeasongItem
    			FItemList(i).Fid  			= rsget("orderid")
    			FItemList(i).Fuserid      	= rsget("userid")
    			FItemList(i).FuserName   	= rsget("username")
    			FItemList(i).FreqName   	= rsget("reqname")
    			FItemList(i).Fsongjangno    = rsget("songjangno")
    			FItemList(i).Fsenddate      = rsget("outputdate")
    			FItemList(i).Fregdate       = rsget("orderdate")
    			FItemList(i).Fprizetitle	= db2html(rsget("itemname"))

    			i=i+1
    			rsget.MoveNext
    		loop
		end if
		rsget.close
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
%>
