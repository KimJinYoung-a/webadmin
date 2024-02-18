<%
'###########################################################################
'	2008�� 01�� 24�� �ѿ�� ����(�߰�)
'###########################################################################

''' ������ý� ����ǰ �ۼ�. Table : [db_order].[dbo].tbl_order_gift_balju
''' ���� Procedure [db_order].[dbo].ten_order_Gift_Maker : ������ù�ȣ�� ����ǰ ��� ����.


Class COrderGiftItem
    public Forderserial		'�ֹ���ȣ
    public Fevt_code		'�̺�Ʈ�ڵ�
    public Fgift_code		'����ǰ�ڵ�
    public Fisupchebeasong		'��۱���
    public Fbaljuid				'�������id
    public Fevt_name			'�̺�Ʈ��
    public Fevt_startdate		'�̺�Ʈ������
    public Fevt_enddate			'�̺�Ʈ������
    public Fgift_scope			'����ǰ����
    public Fgift_type			'����ǰ����
    public Fgift_range1			'����ǰ����
    public Fgift_range2			'����ǰ����
    public Fgift_itemname		'����ǰ��
    public Fgift_img			'����ǰ�̹���
    public Fevtgroup_code		'�̺�Ʈ�׷��ڵ�
    public fbaljudate    		'���������
    public fgift_code_count 	'�̺�Ʈ�ڵ�׷��Ѱ���  
    
    public function GetEventConditionStr()
        dim reStr
        reStr = ""
        if (Fgift_scope="1") then
            reStr = reStr + "��ü ���� �� "
        elseif (Fgift_scope="2") then
            reStr = reStr + "�̺�Ʈ��ϻ�ǰ "
        elseif (Fgift_scope="3") then
            reStr = reStr + "���ú귣���ǰ "
        end if
        
        if (Fgift_type="1") then
            reStr = reStr + "��� ������"
        elseif (Fgift_type="2") then
            if (Fgift_range2>900000) then
                reStr = reStr + CStr(Fgift_range1) + " �� �̻� "
            else 
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " �� "
            end if
        elseif (Fgift_type="3") then
            if (Fgift_range2>900000) then
                reStr = reStr + CStr(Fgift_range1) + " �� �̻� "
            else
                reStr = reStr + CStr(Fgift_range1) + "~" + CStr(Fgift_range2) + " �� "
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
    
    public FRectBaljuid				'�Է� �޾ƿ� �������id
    public FRectIsUpcheBeasong		'�Է� �޾ƿ� ��۱���
    public FRecteventid				'�Է� �޾ƿ� �̺�Ʈid
    public FRectStartdate			'�Է� �޾ƿ� �̺�Ʈ ������
    public FRectEndDate      		'�Է� �޾ƿ� �̺�Ʈ ��������
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
    
public Sub GeteventOrderGiftcount()			'�̺�Ʈ(����ǰ) ������ø���Ʈ ������ ( �׷�:�հ� )
        dim sqlStr,i
        sqlStr = "select count(orderserial) as cnt "
        sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_gift_balju"
        sqlStr = sqlStr + " where 1=1"
         	
        rsget.Open sqlStr, dbget, 1
		    FTotalCount = rsget("cnt")
		rsget.close
										
		sqlStr = "select top "& FPageSize * FCurrPage &""
			if frectdate_display <> "on" then				'��¥ǥ�ð� x �ϰ��
				if frectdateview1 = "no" then				'��������� ����
					sqlStr = sqlStr & " convert(varchar(10),b.baljudate,121) as baljudate,"
				elseif frectdateview1 = "yes" then			'�ֹ��� ����
					sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"
				else
					sqlStr = sqlStr & " convert(varchar(10),c.regdate,121) as baljudate,"		
				end if		
			end if	
		sqlStr = sqlStr & " count(a.gift_code)as gift_code_count,"
			if FRectBaljuid <> "" then					'��������ڵ� �˻�
				sqlStr = sqlStr & " a.baljuid,"		
			end if
			if FRecteventid <> "" then					'�̺�Ʈ�ڵ� �˻�
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

        if FRectBaljuid = "" and FRecteventid = "" and FRectIsUpcheBeasong="" then		'�������id�� �̺�Ʈid�� ��۱����� ��ü�ϰ��(���ε�� �ѷ����� ����������) 
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
		'response.write sqlStr&"<br>"			'������ �ѷ�����.
		
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

public Sub GeteventOrderGiftList()			'�̺�Ʈ(����ǰ) ������ø���Ʈ ������ ( �׷�:���� )
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