<%
'######################################
' 부서별 손익 Class
'######################################
''안분 가능 사업부 표시.
function IsDIVAllowBIZ(bizcd)
    IsDIVAllowBIZ = true
    Exit function
    
    IsDIVAllowBIZ = FALSE
    if (bizcd="0000000101") _
        or (bizcd="0000000201") _
        or (bizcd="0000000301") _
        or (bizcd="0000000401") _
        
        THEN
            IsDIVAllowBIZ = true
    end if
end function

Sub DrawYYYYMMsimpleBox(byval compnm,compval,chgact)
	dim buf,i, j, v

	buf = "<select class='select' name='"&compnm&"' "&chgact&">"
    for i=2011 to Year(now)
        for j=1 to 12
            v = CStr(i)+"-"+Format00(2,j)
    		if (v=CStr(compval)) then
    			buf = buf + "<option value='" + v +"' selected>" + v + "</option>"
    		else
        	buf = buf + "<option value=" + v + " >" +v + "</option>"
            end if
        next
	next
    buf = buf + "</select>"
    response.write buf
end Sub

Class CBizProfitDivMaster
    public FdivMastKey
    public FYYYYMM
    public FpBIZSECTION_CD
    public FpCUST_CD
    public FpACC_CD
    
    public FpBIZSECTION_NM
    public FpCUST_NM
    public FpACC_NM
    public FpACC_USE_CD
    
    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
End Class

Class CBizProfitDivDetail
    public FdivMastKey
    public FBIZSECTION_CD
    public FdivPro
    public FBIZSECTION_NM
    
    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
End Class

Class CBizProfitListItem
    public FBIZSECTION_CD
    public FBIZSECTION_NM
    public FACC_GRP_CD   
    public FACC_GRP_NM   
    public FACC_CD_UP    
    public FACC_CD_UPNM  
    public FACC_CD       
    public FACC_USE_CD   
    public FACC_NM       
    public FdebitSum     
    public FcreditSum    
    
    public FSLDATE
    public Fcust_cd
    public Fcust_NM
    public FBIZ_NO
    public FACC_CD_RMK        

    public FSLTR_SAN_STS            '''0대기 1승인
    public FINTERNAL_TRANS          '''Y내부거래
    
    public ForgBIZSECTION_CD
	public ForgBIZSECTION_NM
	public FdivPro
	public FdivType
	public FdivKey
	public FdivCnt
	
	public FSLTRKEY
	public FSLTRKEY_SEQ
    
    public FOrgDEBIT  
    public FOrgCREDIT 


    public function getDivTypeName()
        if IsNULL(FdivType) then Exit function
        
        if (FdivType=0) then
            getDivTypeName = "결제요청서"
        elseif (FdivType=1) then
            getDivTypeName = "수기안분"
        end if
    end function
    
    public Function IsDivAssigned()
        IsDivAssigned = Not IsNULL(FdivKey)
    end function

    public function IsINTERNALTRANS()
        IsINTERNALTRANS = FINTERNAL_TRANS="Y"
    end function

    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
End Class

Class CBizProfitSumItem
    public FBIZSECTION_CD
    public FBIZSECTION_NM
	public FACC_GRP_CD
	public FACC_GRP_NM
	public FACC_CD_UP
	public FACC_CD_UPNM
	public FACC_CD
	public FACC_USE_CD
	public FACC_NM
	public FdebitSum  
    public FcreditSum 
    public FjpCNT     

    public FOrgBIZSECTION_CD  ''안분이전 사업부문
    public FOrgBIZSECTION_NM  ''안분이전
    public FDIVAssigned       ''안분인지여부
    
    public Function IsDivAssigned()             ''안분된 내역인가..
        IsDivAssigned = (FDIVAssigned="Y")
    end function

    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
End Class

CLASS CBizProfitDivCrossTabItem
    public FBIZSECTION_CD
    public FBIZSECTION_NM
    public FdebitSum  
    public FcreditSum 
    public FCNT     
    
    Private Sub Class_Initialize()
    
    End Sub

	Private Sub Class_Terminate()

	End Sub
	
End Class

Class CBizProfit
    public FItemList()
	public FOneItem
	
    public FPageSize
    public FCurrPage
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	
    
    public FRectBizSecCD
    public FRectStdt
    public FRectEddt
    public FRectSANSTS
    public FRectAccUseCd
    public FRectINTRANS
    public FRectdivAssign
    public FRectdivdpType
    public FRectCustCD
    public FRECTYYYYMM
    public FRectdivMastKey
    
    public FRectSLTRKEY
    public FRectSLTRKEY_SEQ
    
    public Sub getBizProfitDivCrossTabList()
        Dim sqlStr, ArrList, i
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitAssigned_CrossTabList_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitAssigned_CrossTabList"
        END IF
        
        sqlStr =sqlStr+"('"&FRectStdt&"','"&FRectEddt&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectSANSTS&"','"&FRectINTRANS&"','"&FRectCUSTCD&"')"

		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitDivCrossTabItem
		        FItemList(i).FBIZSECTION_CD = ArrList(0,i)
                FItemList(i).FBIZSECTION_NM = ArrList(1,i)
                FItemList(i).FdebitSum      = ArrList(2,i) 
                FItemList(i).FcreditSum     = ArrList(3,i) 
                FItemList(i).FCnt        = ArrList(4,i)

		    Next
		END IF
    end Sub

    public Sub getBizProfitDivCrossList()
        Dim sqlStr, ArrList, i
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitAssigned_CrossList_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitAssigned_CrossList"
        END IF
        
        sqlStr =sqlStr+"('"&FRectStdt&"','"&FRectEddt&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectSANSTS&"','"&FRectINTRANS&"','"&FRectCUSTCD&"')"

		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitListItem
		        FItemList(i).ForgBIZSECTION_CD = ArrList(0,i) 
                FItemList(i).ForgBIZSECTION_NM = ArrList(1,i) 
		        
                FItemList(i).FACC_GRP_CD    = ArrList(2,i) 
                FItemList(i).FACC_GRP_NM    = ArrList(3,i) 
                FItemList(i).FACC_CD_UP     = ArrList(4,i) 
                FItemList(i).FACC_CD_UPNM   = ArrList(5,i)
                FItemList(i).FACC_CD        = ArrList(6,i) 
                FItemList(i).FACC_USE_CD    = ArrList(7,i) 
                FItemList(i).FACC_NM        = ArrList(8,i) 
                FItemList(i).FdebitSum      = ArrList(9,i) 
                FItemList(i).FcreditSum     = ArrList(10,i)
                 
                FItemList(i).FSLDATE    = ArrList(11,i) 
                FItemList(i).Fcust_cd    = ArrList(12,i) 
                FItemList(i).Fcust_NM    = ArrList(13,i) 
                FItemList(i).FBIZ_NO     = ArrList(14,i) 
                FItemList(i).FACC_CD_RMK = ArrList(15,i) 
                FItemList(i).FSLTR_SAN_STS  = ArrList(16,i) 
                FItemList(i).FINTERNAL_TRANS= ArrList(17,i) 
                
                FItemList(i).FBIZSECTION_CD = ArrList(18,i)
                FItemList(i).FBIZSECTION_NM = ArrList(19,i)
                FItemList(i).FdivPro           = ArrList(20,i) 
                FItemList(i).FdivType          = ArrList(21,i) 
                FItemList(i).FdivKey           = ArrList(22,i)
                FItemList(i).FSLTRKEY          = ArrList(23,i)
                FItemList(i).FSLTRKEY_SEQ      = ArrList(24,i)
                
	            FItemList(i).FOrgDEBIT      = ArrList(25,i)
	            FItemList(i).FOrgCREDIT     = ArrList(26,i)

		    Next
		END IF
    end Sub

    public Sub getBizProfitDivedList()
        Dim sqlStr, ArrList, i
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList_DIVed_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList_DIVed"
        END IF
        
        sqlStr =sqlStr+"('"&FRECTSLTRKEY&"','"&FRECTSLTRKEY_SEQ&"')"

		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitListItem
		        FItemList(i).FBIZSECTION_CD = ArrList(0,i)
                FItemList(i).FBIZSECTION_NM = ArrList(1,i)
                FItemList(i).FACC_GRP_CD    = ArrList(2,i) 
                FItemList(i).FACC_GRP_NM    = ArrList(3,i) 
                FItemList(i).FACC_CD_UP     = ArrList(4,i) 
                FItemList(i).FACC_CD_UPNM   = ArrList(5,i)
                FItemList(i).FACC_CD        = ArrList(6,i) 
                FItemList(i).FACC_USE_CD    = ArrList(7,i) 
                FItemList(i).FACC_NM        = ArrList(8,i) 
                FItemList(i).FdebitSum      = ArrList(9,i) 
                FItemList(i).FcreditSum     = ArrList(10,i) 
                FItemList(i).FSLDATE    = ArrList(11,i) 
                FItemList(i).Fcust_cd    = ArrList(12,i) 
                FItemList(i).Fcust_NM    = ArrList(13,i) 
                FItemList(i).FBIZ_NO     = ArrList(14,i) 
                FItemList(i).FACC_CD_RMK = ArrList(15,i) 
                FItemList(i).FSLTR_SAN_STS  = ArrList(16,i) 
                FItemList(i).FINTERNAL_TRANS= ArrList(17,i) 
                
                FItemList(i).ForgBIZSECTION_CD = ArrList(18,i) 
                FItemList(i).ForgBIZSECTION_NM = ArrList(19,i) 
                FItemList(i).FdivPro           = ArrList(20,i) 
                FItemList(i).FdivType          = ArrList(21,i) 
                FItemList(i).FdivKey           = ArrList(22,i)
                ''FItemList(i).FSLTRKEY          = ArrList(24,i)
                ''FItemList(i).FSLTRKEY_SEQ      = ArrList(25,i)
                ''FItemList(i).FdivCnt           = ArrList(26,i)

		    Next
		end if
		
    End Sub

    public Sub getBizProfitDivDetail()
        Dim sqlStr, ArrList, i
        IF (FRectdivMastKey="") then FRectdivMastKey=0
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivDetailList_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivDetailList"
        END IF

        sqlStr =sqlStr+"("&FRectdivMastKey&")"	 
        
        dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitDivDetail
		        FItemList(i).FdivMastKey = ArrList(0,i)
                FItemList(i).FBIZSECTION_CD    = ArrList(1,i) 
                FItemList(i).FdivPro    = ArrList(2,i) 
                FItemList(i).FBIZSECTION_NM     = ArrList(3,i) 
		    Next
		end if
    end Sub

    public Sub getOneBizProfitDivMaster()
        Dim sqlStr, ArrList
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterONe_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterONe"
        END IF
        sqlStr =sqlStr+"("&FRectdivMastKey&")"
        
        dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    
	        set FOneItem = new CBizProfitDivMaster
	        FOneItem.FdivMastKey = ArrList(0,0)
            FOneItem.FYYYYMM = ArrList(1,0)
            FOneItem.FpBIZSECTION_CD    = ArrList(2,0) 
            FOneItem.FpCUST_CD    = ArrList(3,0) 
            FOneItem.FpACC_CD     = ArrList(4,0) 
            FOneItem.FpACC_USE_CD = ArrList(5,0) 
            FOneItem.FpBIZSECTION_NM     = ArrList(6,0) 
            FOneItem.FpCUST_NM     = ArrList(7,0) 
            FOneItem.FpACC_NM     = ArrList(8,0) 
		end if
		
    end Sub
    
    public Sub getOneBizProfitDivMasterBySearch()
        Dim sqlStr, ArrList
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterONEBySearch_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterONEBySearch"
        END IF
        sqlStr =sqlStr+"('"&FRECTYYYYMM&"','"&FRectBizSecCD&"','"&FRectCustCD&"','"&FRectAccUseCd&"')"

        dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    
	        set FOneItem = new CBizProfitDivMaster
	        FOneItem.FdivMastKey = ArrList(0,0)
            FOneItem.FYYYYMM = ArrList(1,0)
            FOneItem.FpBIZSECTION_CD    = ArrList(2,0) 
            FOneItem.FpCUST_CD    = ArrList(3,0) 
            FOneItem.FpACC_CD     = ArrList(4,0) 
            FOneItem.FpACC_USE_CD = ArrList(5,0) 
            FOneItem.FpBIZSECTION_NM     = ArrList(6,0) 
            FOneItem.FpCUST_NM     = ArrList(7,0) 
            FOneItem.FpACC_NM     = ArrList(8,0) 
		end if
		
    end Sub
    
    

    public Sub getBizProfitDivMasterList()
        Dim sqlStr, ArrList, i
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterList_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterList"
        END IF
        
        sqlStr =sqlStr+"('"&FRECTYYYYMM&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectCustCD&"')"	 
        dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitDivMaster
		        FItemList(i).FdivMastKey = ArrList(0,i)
                FItemList(i).FYYYYMM = ArrList(1,i)
                FItemList(i).FpBIZSECTION_CD    = ArrList(2,i) 
                FItemList(i).FpCUST_CD    = ArrList(3,i) 
                FItemList(i).FpACC_CD     = ArrList(4,i) 
                FItemList(i).FpACC_USE_CD = ArrList(5,i) 
                FItemList(i).FpBIZSECTION_NM     = ArrList(6,i) 
                FItemList(i).FpCUST_NM     = ArrList(7,i) 
                FItemList(i).FpACC_NM     = ArrList(8,i) 
		    Next
		end if
    end Sub

    public Sub getBizProfitList()
        Dim sqlStr, ArrList, i
        
        IF (FRectdivdpType="") then FRectdivdpType="0"
        
        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitListCNT_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitListCNT"
        END IF
        
        IF (FRectdivAssign="Y") then
            IF (application("Svr_Info")="Dev") THEN
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitListCNT_DIVASSIGN_TEST"
            ELSE
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitListCNT_DIVASSIGN"
            END IF
        END IF
        
        sqlStr =sqlStr+"('"&FRectStdt&"','"&FRectEddt&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectSANSTS&"','"&FRectINTRANS&"',"&FRectdivdpType&",'"&FRectCUSTCD&"')"	 

		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			FTotalCount = dbiTms_rsget("cnt")
		END IF
		dbiTms_rsget.close
		
		FtotalPage =  CLng(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
    		
		IF (FTotalCount>0) THEN
		    
            
            IF (application("Svr_Info")="Dev") THEN
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList_TEST"
            ELSE
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList"
            END IF
            
            IF (FRectdivAssign="Y") then
                IF (application("Svr_Info")="Dev") THEN
                    sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList_DIVASSIGN_TEST"
                ELSE
                    sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitList_DIVASSIGN"
                END IF
            END IF
            
            sqlStr =sqlStr+"("&FPageSize&","&FCurrPage&",'"&FRectStdt&"','"&FRectEddt&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectSANSTS&"','"&FRectINTRANS&"',"&FRectdivdpType&",'"&FRectCUSTCD&"')"
    
    		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
    			ArrList = dbiTms_rsget.getRows()
    		END IF
    		dbiTms_rsget.close
    		
    		
    		If IsArray(ArrList) then
    		    FResultCount = UBound(ArrList,2)+1
    		    redim preserve FItemList(FResultCount)
    		    For i=0 to FResultCount-1
    		        set FItemList(i) = new CBizProfitListItem
    		        FItemList(i).FBIZSECTION_CD = ArrList(1,i)
                    FItemList(i).FBIZSECTION_NM = ArrList(2,i)
                    FItemList(i).FACC_GRP_CD    = ArrList(3,i) 
                    FItemList(i).FACC_GRP_NM    = ArrList(4,i) 
                    FItemList(i).FACC_CD_UP     = ArrList(5,i) 
                    FItemList(i).FACC_CD_UPNM   = ArrList(6,i)
                    FItemList(i).FACC_CD        = ArrList(7,i) 
                    FItemList(i).FACC_USE_CD    = ArrList(8,i) 
                    FItemList(i).FACC_NM        = ArrList(9,i) 
                    FItemList(i).FdebitSum      = ArrList(10,i) 
                    FItemList(i).FcreditSum     = ArrList(11,i) 
                    FItemList(i).FSLDATE    = ArrList(12,i) 
                    FItemList(i).Fcust_cd    = ArrList(13,i) 
                    FItemList(i).Fcust_NM    = ArrList(14,i) 
                    FItemList(i).FBIZ_NO     = ArrList(15,i) 
                    FItemList(i).FACC_CD_RMK = ArrList(16,i) 
                    FItemList(i).FSLTR_SAN_STS  = ArrList(17,i) 
                    FItemList(i).FINTERNAL_TRANS= ArrList(18,i) 
                    
                    FItemList(i).ForgBIZSECTION_CD = ArrList(19,i) 
                    FItemList(i).ForgBIZSECTION_NM = ArrList(20,i) 
                    FItemList(i).FdivPro           = ArrList(21,i) 
                    FItemList(i).FdivType          = ArrList(22,i) 
                    FItemList(i).FdivKey           = ArrList(23,i)
                    FItemList(i).FSLTRKEY          = ArrList(24,i)
                    FItemList(i).FSLTRKEY_SEQ      = ArrList(25,i)
                    FItemList(i).FdivCnt           = ArrList(26,i)

    		    Next
    		end if
    	END IF
    end Sub

    public Sub getBizProfitSum()
        Dim sqlStr, ArrList, i

        IF (application("Svr_Info")="Dev") THEN
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfit_TEST"
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfit"
        END IF
        
        IF (FRectdivAssign="Y") then
            IF (application("Svr_Info")="Dev") THEN
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfit_DIVASSIGN_TEST"
            ELSE
                sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfit_DIVASSIGN"
            END IF
        END IF
        
        IF (FRectdivdpType="") then FRectdivdpType="0"
        
        ''IF (FRectBizSecCD="") then FRectBizSecCD="NULL"
            
        sqlStr =sqlStr+"('"&FRectStdt&"','"&FRectEddt&"','"&FRectBizSecCD&"','"&FRectAccUseCd&"','"&FRectSANSTS&"','"&FRectINTRANS&"',"&FRectdivdpType&")"	 
        
		dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
			ArrList = dbiTms_rsget.getRows()
		END IF
		dbiTms_rsget.close
		
		If IsArray(ArrList) then
		    FResultCount = UBound(ArrList,2)+1
		    redim preserve FItemList(FResultCount)
		    For i=0 to FResultCount-1
		        set FItemList(i) = new CBizProfitSumItem
		        FItemList(i).FBIZSECTION_CD = ArrList(0,i)
                FItemList(i).FBIZSECTION_NM = ArrList(1,i)
                FItemList(i).FACC_GRP_CD    = ArrList(2,i) 
                FItemList(i).FACC_GRP_NM    = ArrList(3,i) 
                FItemList(i).FACC_CD_UP     = ArrList(4,i) 
                FItemList(i).FACC_CD_UPNM   = ArrList(5,i)
                FItemList(i).FACC_CD        = ArrList(6,i) 
                FItemList(i).FACC_USE_CD    = ArrList(7,i) 
                FItemList(i).FACC_NM        = ArrList(8,i) 
                FItemList(i).FdebitSum      = ArrList(9,i) 
                FItemList(i).FcreditSum     = ArrList(10,i) 
                FItemList(i).FjpCNT         = ArrList(11,i) 
                
                FItemList(i).FOrgBIZSECTION_CD  = ArrList(12,i) 
                FItemList(i).FOrgBIZSECTION_NM  = ArrList(13,i) 
                FItemList(i).FDIVAssigned       = ArrList(14,i) 

		    Next
		end if
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub
	
    public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class
%>