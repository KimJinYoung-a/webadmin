<%
Class CLecUserItem
    public Flecturer_id
    public Flecturer_name
    public Fen_name
    public Flec_yn
    public Fdiy_yn
    public Flec_margin
    public Fmat_margin
    public Fdiy_margin
    public Fdiy_dlv_gubun
    public Fregdate
    
    public FDefaultFreebeasongLimit
    public FDefaultDeliveryPay
    
    
    public FTen_socname_kor
    public FTen_socname
    public FTen_defaultmargine
    public FTen_defaultFreebeasongLimit
    public FTen_defaultdeliverPay
    public FTen_defaultdeliverytype
    
    function getTenDlvStr()
        if (FTen_defaultdeliverytype=-1) then
            getTenDlvStr = ""
        elseif (FTen_defaultFreebeasongLimit=0 and FTen_defaultdeliverPay=0 and FTen_defaultdeliverytype=0) then
            getTenDlvStr = "" ''"업체무료배송"
        else
            if (FTen_defaultdeliverytype=9) then
                getTenDlvStr= FTen_defaultFreebeasongLimit&"원 이상 무료배송 미만 배송비"&FTen_defaultdeliverPay
            elseif (FTen_defaultdeliverytype=7) then
                getTenDlvStr="업체착불배송"
            end if
        end if
        
    end function

    
    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class


Class CLecUser
    public FOneItem
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	
	public FRectLecturerID
	
	public sub getTenLecUserInfo
	    Dim sqlStr
	    
	    sqlStr = " select c.userid,"
	    sqlStr = sqlStr & " c.socname_kor as Ten_socname_kor,c.socname as Ten_socname,"
	    sqlStr = sqlStr & " c.defaultmargine as Ten_defaultmargine,"
	    sqlStr = sqlStr & " c.defaultFreebeasongLimit as Ten_defaultFreebeasongLimit,"
	    sqlStr = sqlStr & " c.defaultdeliverPay as Ten_defaultdeliverPay,"
	    sqlStr = sqlStr & " c.defaultdeliverytype as Ten_defaultdeliverytype from [TENDB].db_user.dbo.tbl_user_c c"
	    sqlStr = sqlStr & " where c.userid='" & FRectLecturerID & "'"
	    
	    rsacademyget.CursorLocation = adUseClient
        rsacademyget.Open sqlStr,dbacademyget,adOpenForwardOnly, adLockReadOnly

	    if  not rsacademyget.EOF  then
	        set FOneItem = new CLecUserItem
	        
	        FOneItem.Flecturer_id   = rsacademyget("userid")
            
            FOneItem.FTen_socname_kor              =  db2HTML(rsacademyget("Ten_socname_kor"))
            FOneItem.FTen_socname                   =  db2HTML(rsacademyget("Ten_socname"))
            FOneItem.FTen_defaultmargine            = rsacademyget("Ten_defaultmargine")
            FOneItem.FTen_defaultFreebeasongLimit   = rsacademyget("Ten_defaultFreebeasongLimit")
            FOneItem.FTen_defaultdeliverPay         = rsacademyget("Ten_defaultdeliverPay")
            FOneItem.FTen_defaultdeliverytype       = rsacademyget("Ten_defaultdeliverytype")

        else
            set FOneItem = new CLecUserItem
            
            FOneItem.Flecturer_id   = FRectLecturerID
            
            FOneItem.FTen_socname_kor              =  ""
            FOneItem.FTen_socname                   =  ""
            FOneItem.FTen_defaultmargine            = 0
            FOneItem.FTen_defaultFreebeasongLimit   = 0
            FOneItem.FTen_defaultdeliverPay         = 0
            FOneItem.FTen_defaultdeliverytype       = -1

	    end if
	    rsacademyget.close
	    
    end sub

	public Sub getOneLecUserInfo
	    Dim sqlStr
	    
	    sqlStr = " select l.*"
        sqlStr = sqlStr & " from db_academy.dbo.tbl_lec_user l"
	    sqlStr = sqlStr & " where l.lecturer_id='" & FRectLecturerID & "'"
	    
	    rsacademyget.CursorLocation = adUseClient
        rsacademyget.Open sqlStr,dbacademyget,adOpenForwardOnly, adLockReadOnly

	    if  not rsacademyget.EOF  then
	        set FOneItem = new CLecUserItem
	        
	        FOneItem.Flecturer_id   = rsacademyget("lecturer_id")
            FOneItem.Flecturer_name = db2HTML(rsacademyget("lecturer_name"))
            FOneItem.Fen_name       = db2HTML(rsacademyget("en_name"))
            FOneItem.Flec_yn        = rsacademyget("lec_yn")
            FOneItem.Fdiy_yn        = rsacademyget("diy_yn")
            FOneItem.Flec_margin    = rsacademyget("lec_margin")
            FOneItem.Fmat_margin    = rsacademyget("mat_margin")
            FOneItem.Fdiy_margin    = rsacademyget("diy_margin")
            FOneItem.Fdiy_dlv_gubun = rsacademyget("diy_dlv_gubun")
            FOneItem.Fregdate       = rsacademyget("regdate")
            
            FOneItem.FDefaultFreebeasongLimit = rsacademyget("DefaultFreebeasongLimit")
            FOneItem.FDefaultDeliveryPay      = rsacademyget("DefaultDeliveryPay")
            
        else
            set FOneItem = new CLecUserItem
            
            FOneItem.Flecturer_id   = FRectLecturerID
            FOneItem.Flecturer_name = ""
            FOneItem.Flec_yn        = ""
            FOneItem.Fdiy_yn        = ""
            FOneItem.Flec_margin    = 0
            FOneItem.Fmat_margin     = 0
            FOneItem.Fdiy_margin    = 0
            FOneItem.Fdiy_dlv_gubun = 0
            FOneItem.Fregdate       = NULL
            
            FOneItem.FDefaultFreebeasongLimit = 0
            FOneItem.FDefaultDeliveryPay      = 0
            

	    end if
	    rsacademyget.close
    end Sub
    
    Private Sub Class_Initialize()
		'redim preserve FItemList(0)
		redim  FItemList(0)
		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
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