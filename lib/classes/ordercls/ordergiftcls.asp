<%
'###########################################################################
'	2008월 01월 24일 한용민 수정(추가)
'###########################################################################

''' 출고지시시 사은품 작성. Table : [db_order].[dbo].tbl_order_gift_balju
''' 관련 Procedure [db_order].[dbo].ten_order_Gift_Maker : 출고지시번호로 사은품 목록 생성.


Class COrderGiftItem
    public Forderserial		'주문번호
    public Fevt_code		'이벤트코드
    public Fgift_code		'사은품코드
    public Fisupchebeasong		'배송구분
    public Fbaljuid				'출고지시id
    public Fevt_name			'이벤트명
    public Fevt_startdate		'이벤트시작일
    public Fevt_enddate			'이벤트끝난일
    public Fgift_scope			'사은품조건
    public Fgift_type			'사은품조건
    public Fgift_range1			'사은품조건
    public Fgift_range2			'사은품조건
    public Fgift_itemname		'사은품명
    public Fgift_img			'사은품이미지
    public Fevtgroup_code		'이벤트그룹코드
    public fbaljudate    		'출고지시일
    public fgift_code_count 	'이벤트코드그룹총갯수  
    
    public function GetEventConditionStr()
        dim reStr
        reStr = ""
        if (Fgift_scope="1") then
            reStr = reStr + "전체 구매 고객 "
        elseif (Fgift_scope="2") then
            reStr = reStr + "이벤트등록상품 "
        elseif (Fgift_scope="3") then
            reStr = reStr + "선택브랜드상품 "
        end if
        
        if (Fgift_type="1") then
            reStr = reStr + "모든 구매자"
        elseif (Fgift_type="2") then
            if (Fgift_range2>900000) then
                reStr = reStr + CStr(Fgift_range1) + " 원 이상 "
            else 
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 원 "
            end if
        elseif (Fgift_type="3") then
            if (Fgift_range2>900000) then
                reStr = reStr + CStr(Fgift_range1) + " 개 이상 "
            else
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " 개 "
            end if
        end if
        
        GetEventConditionStr = reStr
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class COrderGift
    public FItemList()
	public FOneItem

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
    
    public FRectBaljuid				'입력 받아올 출고지시id
    public FRectIsUpcheBeasong		'입력 받아올 배송구분
    public FRecteventid				'입력 받아올 이벤트id
    public FRectStartdate			'입력 받아올 이벤트 시작일
    public FRectEndDate      		'입력 받아올 이벤트 마지막일
    public frectdateview
    public frectdateview1
    public frectdate_display      
    
    public Sub GetOrderGiftList()
        dim sqlStr,i
        sqlStr = "select count(orderserial) as cnt "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift_balju"
        sqlStr = sqlStr + " where 1=1"
        
        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and baljuid=" + CStr(FRectBaljuid) + ""
        end if
        
        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and IsUpcheBeasong='" + FRectIsUpcheBeasong + "'"
        end if
        
        rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.close
		
		
		sqlStr = "select top " + CStr(FPageSize * FCurrPage) 
		sqlStr = sqlStr + " * from [db_order].[dbo].tbl_order_gift_balju "
		sqlStr = sqlStr + " where 1=1"
        
		if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and baljuid=" + CStr(FRectBaljuid) + ""
        end if
        
        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and IsUpcheBeasong='" + FRectIsUpcheBeasong + "'"
        end if
        
		sqlStr = sqlStr + " order by baljuid, evt_code, gift_code, orderserial"
				
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem
				FItemList(i).Forderserial    = rsget("orderserial")
                FItemList(i).Fevt_code       = rsget("evt_code")
                FItemList(i).Fgift_code      = rsget("gift_code")
                FItemList(i).Fisupchebeasong = rsget("isupchebeasong")
                FItemList(i).Fbaljuid        = rsget("baljuid")
                FItemList(i).Fevt_name       = db2html(rsget("evt_name"))
                FItemList(i).Fevt_startdate  = rsget("evt_startdate")
                FItemList(i).Fevt_enddate    = rsget("evt_enddate")
                FItemList(i).Fgift_scope     = rsget("gift_scope")
                FItemList(i).Fgift_type      = rsget("gift_type")
                FItemList(i).Fgift_range1    = rsget("gift_range1")
                FItemList(i).Fgift_range2    = rsget("gift_range2")
                FItemList(i).Fgift_itemname  = db2html(rsget("gift_itemname"))
                FItemList(i).Fgift_img       = rsget("gift_img")
                FItemList(i).Fevtgroup_code  = rsget("evtgroup_code")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close
		
    end Sub
    
public Sub GeteventOrderGiftcount()			'이벤트(사은품) 출고지시리스트 페이지 ( 그룹:합계 )
        dim sqlStr,i
        sqlStr = "select count(orderserial) as cnt "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift_balju"
        sqlStr = sqlStr + " where 1=1"
         	
        rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.close
										
		sqlStr = "select top "& FPageSize * FCurrPage &""
			if frectdate_display <> "on" then				'날짜표시가 x 일경우
				if frectdateview1 = "no" then				'출고지시일 기준
					sqlStr = sqlStr & " convert(varchar(10),b.baljudate,121) as baljudate,"
				elseif frectdateview1 = "yes" then			'주문일 기준
					sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"
				else
					sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"		
				end if		
			end if	
		sqlStr = sqlStr & " count(a.gift_code)as gift_code_count,"
			if FRectBaljuid <> "" then					'출고지시코드 검색
				sqlStr = sqlStr & " a.baljuid,"		
			end if
			if FRecteventid <> "" then					'이벤트코드 검색
				sqlStr = sqlStr & " a.evt_code,"		
			end if			
		sqlStr = sqlStr & " a.evt_code,a.evt_name,a.gift_code,a.gift_itemname,a.isupchebeasong,a.gift_type,"
		sqlStr = sqlStr & " a.gift_range1,a.gift_range2,a.gift_scope"		
		sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_gift_balju a" 
		sqlStr = sqlStr & " join db_order.[dbo].tbl_baljumaster b"
		sqlStr = sqlStr & " on a.baljuid = b.id" 		
		sqlStr = sqlStr & " join [db_order].[dbo].tbl_order_master c"
		sqlStr = sqlStr & " on a.orderserial=c.orderserial"		 
		sqlStr = sqlStr & " where 1=1"

        if FRectBaljuid = "" and FRecteventid = "" and FRectIsUpcheBeasong="" then		'출고지시id와 이벤트id와 배송구분이 전체일경우(폼로드시 뿌려지는 페이지없음) 
        	sqlStr = sqlStr + " and a.evt_code='0'"
        end if
		
        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and a.baljuid='" + FRectBaljuid + "'"
        end if		    
        
        if FRecteventid <> "" then
            sqlStr = sqlStr + " and a.evt_code='" + FRecteventid + "'"
        end if 
        
        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and a.IsUpcheBeasong='" + FRectIsUpcheBeasong + "'"
        end if
        if frectdateview = "no" then
	         if frectdateview1 = "no" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and b.baljudate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if             
	        end if         
	        if frectdateview1 = "yes" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and c.regdate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if             
	        end if  
   		end if       
		sqlStr = sqlStr & " group by"
			if frectdate_display <> "on" then		
				if frectdateview1 = "no" then
					sqlStr = sqlStr & " convert(varchar(10),b.baljudate,121),"
				elseif frectdateview1 = "yes" then
				sqlStr = sqlStr & " convert(varchar(10),c.regdate,121),"		
				else
				sqlStr = sqlStr & " convert(varchar(10),c.regdate,121),"
				end if	
			end if	
		sqlStr = sqlStr & " a.evt_code,a.gift_code,a.evt_name,a.gift_scope,"
			if FRectBaljuid <> "" then
				sqlStr = sqlStr & " a.baljuid,"		
			end if		
			if FRecteventid <> "" then
			sqlStr = sqlStr & " a.evt_code,"		
			end if	
		sqlStr = sqlStr & " a.gift_itemname,a.isupchebeasong,a.gift_type,a.gift_range1,a.gift_range2"
		sqlStr = sqlStr & " order by"
		sqlStr = sqlStr & " a.gift_code"
			if frectdate_display <> "on" then		
				if frectdateview1 = "no" then
				sqlStr = sqlStr & " ,convert(varchar(10),b.baljudate,121)"
				elseif frectdateview1 = "yes" then
				sqlStr = sqlStr & " ,convert(varchar(10),c.regdate,121)"		
				else
				sqlStr = sqlStr & " ,convert(varchar(10),c.regdate,121)"
				end if
			end if	
			if FRecteventid <> "" then
				sqlStr = sqlStr & " ,a.evt_code desc"		
			end if				
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		'response.write sqlStr&"<br>"			'오류시 뿌려본다.
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem
				
					if FRectBaljuid <> "" then				
	            	    FItemList(i).Fbaljuid      = rsget("baljuid")
	        	    end if
                FItemList(i).Fgift_code      = rsget("gift_code")
                FItemList(i).Fevt_code      = rsget("evt_code")
                FItemList(i).Fisupchebeasong = rsget("isupchebeasong")
                FItemList(i).Fevt_name       = db2html(rsget("evt_name"))
                FItemList(i).Fgift_type     = rsget("gift_type")
                FItemList(i).Fgift_itemname  = db2html(rsget("gift_itemname"))
					if frectdate_display <> "on" then                
             		   FItemList(i).Fbaljudate  = rsget("baljudate")					
					end if 
				FItemList(i).fgift_code_count  = rsget("gift_code_count")
				FItemList(i).Fgift_range1    = rsget("gift_range1")
				FItemList(i).Fgift_range2    = rsget("gift_range2")
				FItemList(i).fgift_scope    = rsget("gift_scope")
					if FRecteventid <> "" then               
						FItemList(i).Fevt_code       = rsget("evt_code")
	          		end if				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close		
end Sub  

public Sub GeteventOrderGiftList()			'이벤트(사은품) 출고지시리스트 페이지 ( 그룹:내역 )
        dim sqlStr,i
        sqlStr = "select count(orderserial) as cnt "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift_balju"
        sqlStr = sqlStr + " where 1=1"
        
        rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.close	
		
		sqlStr = "select top " + CStr(FPageSize * FCurrPage)
			if frectdateview1 = "no" then
			sqlStr = sqlStr & " convert(varchar(10),b.baljudate,121) as baljudate,"
			elseif frectdateview1 = "yes" then
			sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"
			else
			sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"		
			end if			
		sqlStr = sqlStr & " a.*" 
		sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_gift_balju a"
		sqlStr = sqlStr & " join db_order.[dbo].tbl_baljumaster b"
		sqlStr = sqlStr & " on a.baljuid = b.id"		
		sqlStr = sqlStr & " join [db_order].[dbo].tbl_order_master c"
		sqlStr = sqlStr & " on a.orderserial=c.orderserial"		 		
		sqlStr = sqlStr & " where 1=1"
        
        if FRectBaljuid = "" and FRecteventid = "" and FRectIsUpcheBeasong="" then
        	sqlStr = sqlStr + " and a.evt_code='0'"
        end if	

        if FRectBaljuid<>"" then
            sqlStr = sqlStr + " and a.baljuid='" + FRectBaljuid + "'"
        end if            		
    
        if FRecteventid <> "" then
            sqlStr = sqlStr + " and a.evt_code='" + FRecteventid + "'"
        end if 
        
        if FRectIsUpcheBeasong<>"" then
            sqlStr = sqlStr + " and a.IsUpcheBeasong='" + FRectIsUpcheBeasong + "'"
        end if
        
        if frectdateview = "no" then
	         if frectdateview1 = "no" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and b.baljudate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if             
	        end if         
	        if frectdateview1 = "yes" then
		        if FRectStartdate<>"" then
		            sqlStr = sqlStr + " and c.regdate between '" & FRectStartdate & "' and  '" & FRectEndDate & "'"
		        end if             
	        end if  
   		end if                    
        
		sqlStr = sqlStr & " order by b.baljudate,a.baljuid ,a.evt_code, a.gift_code, a.orderserial desc"		
		
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1
		'response.write sqlStr&"<br>"
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		if  not rsget.EOF  then
		    i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new COrderGiftItem
				FItemList(i).Forderserial    = rsget("orderserial")
                FItemList(i).Fevt_code       = rsget("evt_code")
                FItemList(i).Fgift_code      = rsget("gift_code")
                FItemList(i).Fisupchebeasong = rsget("isupchebeasong")
                FItemList(i).Fbaljuid        = rsget("baljuid")
                FItemList(i).Fevt_name       = db2html(rsget("evt_name"))
                FItemList(i).Fevt_startdate  = rsget("evt_startdate")
                FItemList(i).Fevt_enddate    = rsget("evt_enddate")
                FItemList(i).Fgift_scope     = rsget("gift_scope")
                FItemList(i).Fgift_itemname  = db2html(rsget("gift_itemname"))
                FItemList(i).Fbaljudate  = rsget("baljudate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.close		
end Sub  
       
    Private Sub Class_Initialize()
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