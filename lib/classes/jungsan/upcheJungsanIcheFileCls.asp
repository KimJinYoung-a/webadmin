<%
Class CIcheFileMasterItem
    public FipFileNo
    public FipFileName
    public FipFileRegdate
    public FipFileState
    public FipFileGbn
    public FIcheDate
    public FpayreqIdx
    
    function isWonChonFile()
        isWonChonFile = (FipFileGbn="WN")
    end function
    
    function getBizSectionCD()
        if (FipFileGbn="ON") then getBizSectionCD="0000000101"
        if (FipFileGbn="OF") then getBizSectionCD="0000000201"
    end function

End Class

Class CIcheFileWithEseroList
    public FipFileNo
    public FipFileState
    public FipFileGbn
    public FIcheDate
    public FpayreqIdx
    
    public FipFileDetailIdx
    public FtargetGbn
    public FtargetIdx
    public FipFileDetailState
    public FrefIpFileDetailIdx
    
    public FtaxKey
    public FAppDate
    public FSellCorpNo
    public Ftax_cust_cd
    public FTax_arap_cd
    public FmatchSeq
    public FmatchType
    public FmatchState
    public FbizSecCd
    public FerpLinkType
    public FerpLinkKey
    
    public FTotSum
    public Fsuplysum
    public FtaxSum
    public FdtlName
    public FSellCorpName
    
    function getIpFileDetailStateName()
        if IsNULL(FipFileDetailState) then Exit function
        
        if (FipFileDetailState=0) then
            getIpFileDetailStateName = "작성중"
        elseif (FipFileDetailState=7) then
            getIpFileDetailStateName = "입금완료"
        elseif (FipFileDetailState=8) then
            getIpFileDetailStateName = "ERP서류전송"
        else
            getIpFileDetailStateName = Cstr(FipFileDetailState)
        end if
    end function

End Class

Class CupcheJungsanIcheFile
    public FSDate
    public FEDate
    
    public FItemList()
    public FOneItem
    public FTotCnt
    public FPageSize
    public FCurrPage
    public FResultCount
    public FSPageNo
    public FEPageNo 
    
    public FRectipFileNo
    public FRectisMappingYn
    public FRectErpSendState
    public FRectDetailState
    
'    public FsearchText 
'    public FtaxsellType
'    public FtaxModiType
'    public FtaxType
'    public FMappingTypeYn
'    public FMappingType
'    public FTaxKey
'    public FRectCorpNo
'    public FErpSendType
    
    Function getOneIcheFileMaster
        Dim sqlStr,i
        sqlStr = "select * from db_jungsan.dbo.tbl_jungsan_ipkumFile_Master where ipFileNo="&FRectIpFileNo
        
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            set FOneItem = new CIcheFileMasterItem
            FOneItem.FipFileNo     = rsget("ipFileNo")
            FOneItem.FipFileName   = rsget("ipFileName")
            FOneItem.FipFileRegdate= rsget("ipFileRegdate")
            FOneItem.FipFileState  = rsget("ipFileState")
            FOneItem.FipFileGbn    = rsget("ipFileGbn")
            FOneItem.FIcheDate     = rsget("IcheDate")
            FOneItem.FpayreqIdx    = rsget("payreqIdx")
        end if
        rsget.Close
    end Function
    
    Function fnGetIcheFileMappingList()
        Dim strSql, ArrList, i
        strSql ="[db_partner].[dbo].[sp_Ten_Esero_getIcheFileMappingListCnt]("&FRectipFileNo&",'"&FRectisMappingYn&"','"&FRectErpSendState&"','"&FRectDetailState&"')"	 
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
	
		IF FTotCnt > 0 THEN
    		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
    		FEPageNo = FPageSize*FCurrPage		
    		
    	    strSql ="[db_partner].[dbo].sp_Ten_Esero_getIcheFileMappingList("&FRectipFileNo&",'"&FRectisMappingYn&"','"&FRectErpSendState&"','"&FRectDetailState&"',"&FsPageNO&","&FePageNO&")"	 
    		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    		IF Not (rsget.EOF OR rsget.BOF) THEN
    			ArrList = rsget.getRows()
    		END IF
    		rsget.close
    		
    		If IsArray(ArrList) then
    		    FResultCount = UBound(ArrList,2)+1
    		    redim preserve FItemList(FResultCount)
    		    For i=0 to FResultCount-1
    		        set FItemList(i) = new CIcheFileWithEseroList
    		        FItemList(i).FipFileNo          = ArrList(0,i) ''ipFileNo
                    FItemList(i).FipFileState       = ArrList(1,i) ''ipFileState
                    FItemList(i).FipFileGbn         = ArrList(2,i) ''ipFileGbn
                    FItemList(i).FIcheDate          = ArrList(3,i) ''IcheDate
                    FItemList(i).FpayreqIdx         = ArrList(4,i) ''payreqIdx 
                                       
                    FItemList(i).FipFileDetailIdx   = ArrList(5,i) ''ipFileDetailIdx
                    FItemList(i).FtargetGbn         = ArrList(6,i) ''targetGbn
                    FItemList(i).FtargetIdx         = ArrList(7,i) ''targetIdx
                    FItemList(i).FipFileDetailState = ArrList(8,i) ''ipFileDetailState
                    FItemList(i).FrefIpFileDetailIdx= ArrList(9,i) ''refIpFileDetailIdx
                                       
                    FItemList(i).FtaxKey            = ArrList(10,i) ''taxKey,
                    FItemList(i).FAppDate           = ArrList(11,i) '' AppDate, 
                    FItemList(i).FSellCorpNo        = ArrList(12,i) ''SellCorpNo, 
                    FItemList(i).Ftax_cust_cd       = ArrList(13,i) ''tax_cust_cd,
                    FItemList(i).FTax_arap_cd       = ArrList(14,i) '' Tax_arap_cd
                    FItemList(i).FmatchSeq          = ArrList(15,i) ''matchSeq, 
                    FItemList(i).FmatchType         = ArrList(16,i) ''matchType, 
                    FItemList(i).FmatchState        = ArrList(17,i) ''matchState,
                    FItemList(i).FbizSecCd          = ArrList(18,i) '' bizSecCd,
                    FItemList(i).FerpLinkType       = ArrList(19,i) '' erpLinkType, 
                    FItemList(i).FerpLinkKey        = ArrList(20,i) ''erpLinkKey
                    
                    FItemList(i).FTotSum            = ArrList(21,i) ''TotSum
                    FItemList(i).Fsuplysum          = ArrList(22,i) ''suplysum
                    FItemList(i).FtaxSum            = ArrList(23,i) ''taxSum
                    FItemList(i).FdtlName           = ArrList(24,i) ''dtlName
                    FItemList(i).FSellCorpName      = ArrList(25,i) ''SellCorpName

    		    Next
    		end if
        END IF
    end function
End Class
%>