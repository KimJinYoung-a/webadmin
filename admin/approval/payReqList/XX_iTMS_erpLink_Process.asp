<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 ERP연동
' History : 2011.12.16 eastone  생성 erpLink_Process.asp
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->  
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/lib/classes/approval/payreqListCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->  

<%
DIM isTESTMODE : isTESTMODE = TRUE
Dim LTp : LTp = reQuestCheckvar(request("LTp"),10)  ''타입
Dim chk : chk = request("chk")                      ''Idx Array
Dim ipFileNo : ipFileNo = request("ipFileNo")
rw LTp
rw chk

chk = split(chk,",")

dim i,jj
Dim payIdx, paramInfo, sqlStr
Dim retParamInfo, RetErr, retErrStr, retErpLinkType
Dim errALL
Dim retARR,iTR_MAST_SEQ,iSETTLE_REAL_DATE, iSETTLE_YYYYMM
Dim AssignedRow, tmpAssign, erpLinkType
Dim ierpLinkSeq, iExt5

IF (LTp="A") THEN
    For i=UBound(chk) to LBound(chk) STEP -1
        IF Trim(chk(i))<>"" THEN
            if Trim(chk(i))<70 and Trim(chk(i))<>67 THen
                response.write "과거내역 업로드 불가"
                response.end
            end if
        END IF
    Next
ENd IF


IF (LTp="A") THEN
  For i=UBound(chk) to LBound(chk) STEP -1   ''역으로 등록.
    payIdx = Trim(chk(i))
    IF (payIdx<>"") Then
        ''rw payIdx
        ''검증은 ERP 프로시져에서.
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@payRequestIdx"	,adInteger, adParamInput,, payIdx) _
            ,Array("@erpLinkType"	,adVarchar, adParamOutput,1, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_payREQ_2015" 
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 생성된 송장번호 
        retErpLinkType = GetValue(retParamInfo, "@erpLinkType")   'S:영업,F:자금수지,C:회계
        
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
        
  if (NOT isTESTMODE) then     
        IF (RetErr<0) THEN
            errALL = errALL&"["&payIdx&"] (ERR:"&RetErr&") "&retErrStr&VbCRLF
        ELSE
            sqlStr = "Update db_Partner.dbo.tbl_eAppPayRequest"
            sqlStr = sqlStr & " SET erpLinkKey='"&RetErr&"'"
            sqlStr = sqlStr & " ,erpLinkType='"&retErpLinkType&"'"
            sqlStr = sqlStr & " ,payRequestState=8"
            sqlStr = sqlStr & " WHERE payREquestIdx="&payIdx
            
            dbget.Execute sqlStr
            
            IF (retErpLinkType="S") then  ''문서 업데이트
                sqlStr = " update D"
                sqlStr = sqlStr & " set erpDocLinkType='"&retErpLinkType&"'"
                sqlStr = sqlStr & " ,erpDocLinkKey="&RetErr
                sqlStr = sqlStr & " ,erpDocSendDate=getdate()"
                sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppPayDoc D"
                sqlStr = sqlStr & " WHERE payREquestIdx="&payIdx
                
                dbget.Execute sqlStr
            END IF
                
        END IF
  end if
    END IF
    
    
  Next
ELSEIF (LTp="AF") THEN          '''입금확정 File - 업체 정기결제.
    IF (ipFileNo<>"") then
        ''검증은 ERP 프로시져에서.
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@ipFileNo"	,adInteger, adParamInput,, ipFileNo) _
            ,Array("@erpLinkType"	,adVarchar, adParamOutput,1, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_ICheFile_2015" 
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE")         ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")           '   
        retErpLinkType = GetValue(retParamInfo, "@erpLinkType")     ' S:영업,F:자금수지,C:회계
        
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    if (NOT isTESTMODE) then    
        IF (RetErr<0) THEN
            errALL = errALL&"["&ipFileNo&"] (ERR:"&RetErr&") "&retErrStr&VbCRLF
        ELSE
            sqlStr = "Update db_jungsan.dbo.tbl_jungsan_ipkumFile_Master"
            sqlStr = sqlStr & " SET ipFileState=3"
            sqlStr = sqlStr & " ,ERP_TrMastSeq="&RetErr                                '''자금 수지 Key
            sqlStr = sqlStr & " WHERE ipFileNo="&ipFileNo
            
            dbget.Execute sqlStr
                
        END IF
    end if	    
        ''이체등록내역 update
   if (FALSE) then  ''필요 없을듯 함..     
        IF application("Svr_Info")="Dev" THEN
            sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_2015_TEST (4,"&RetErr&") "
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_2015 (4,"&RetErr&") "
        END IF
        
    	dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    	IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
    		retARR = dbiTms_rsget.getRows()   
    	END IF
    	dbiTms_rsget.close
    	
    	AssignedRow = 0
    	''  d.plan_seq,d.plan_list_seq,d.plan_amt,d.SCM_DTLIDX 
    if (NOT isTESTMODE) then   
    
    	IF (IsARRAY(retARR)) THEN
    	    For jj = 0 To UBound(retARR,2)
    	        ierpLinkSeq         = Trim(retARR(1,jj)) ''CStr(retARR(0,jj))+CStr(Format00(5,Trim(retARR(1,jj))))
    	        iExt5               = Trim(retARR(3,jj))
    	        
    	        if (CStr(iExt5)<>"") THEN
        	        sqlStr = " Update db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
        	        sqlStr = sqlStr & " SET erpDTLSeq='"&ierpLinkSeq&"'"
        	        sqlStr = sqlStr & " ,ipFileDetailState=1"                                    ''전송
        	        sqlStr = sqlStr & " WHERE ipFileNo="&ipFileNo
        	        sqlStr = sqlStr & " and isNULL(refipFileDetailIDx,ipFileDetailIDx)="&iExt5
        
                    dbget.Execute sqlStr,tmpAssign
                    
                    AssignedRow = AssignedRow + tmpAssign
                end if
    	    Next
    	END IF
     end if
    end if	
    	rw AssignedRow&"건 반영됨"
    	
    end if
ELSEIF (LTp="R") THEN
ELSE
    rw "지정안됨 LTp="&LTp
END IF
rw errALL
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->