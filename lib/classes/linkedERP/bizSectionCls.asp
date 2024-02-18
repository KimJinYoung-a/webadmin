<%
'############################
' Description : ERP 연동 부서관리
' History : 2011.04.21 정윤정  생성
'############################

Class CBizSection
public FBS_NM
public FUSE_YN
public FOnlySub
public FSale
public FView
public FYYYYMM
public Fpart_sn
public FGRP_YN
public FBizsection_cd
public FisRegularMember
public FSearchType
public FSearchText
public Fdepartment_id
public Finc_subdepartment

	'erp 부서리스트
	' /admin/linkedERP/Biz/popGetBizOne.asp		' /admin/linkedERP/Biz/popGetBiz.asp
	public Function fnGetBizSectionList
		Dim strSql

		strSql = "db_partner.dbo.sp_Ten_TMS_BA_BIZSECTION_getList('"&FBS_NM&"','"&FUSE_YN&"','"&FOnlySub&"','"&FSale&"','"&FView&"')"

		'response.write strSql & "<br>"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizSectionList = rsget.getRows()
		END IF
		rsget.close
	End Function


'erp 사용자 부서 전체리스트
	public Function fnGetBizSectionAllList
		Dim strSql
		IF Fpart_sn = "" THEN Fpart_sn = 0
		strSql = "db_partner.dbo.sp_Ten_user_Bizsection_getAllList('"&FYYYYMM&"', "&Fpart_sn&",'"&FBizsection_cd&"', '"&FUSE_YN&"', '"&FisRegularMember&"','"&FSearchType&"','"&FSearchText&"', '" + CStr(Fdepartment_id) + "', '" + CStr(Finc_subdepartment) + "' )"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizSectionAllList = rsget.getRows()
		END IF
		rsget.close
	End Function



'erp 부서리스트 - 월별 손익데이터 있는 부서만
	public Function fnGetBizMonthProftist
		Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_TMS_BA_BIZSECTION_getMonthProfitList]('"&FGRP_YN&"','"&FYYYYMM&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProftist = rsget.getRows()
		END IF
		rsget.close
	End Function

'erp 부서리스트 - 월별 손익데이터 업무비율 구분  있는 부서만
	public Function fnGetBizMonthBizList
		Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_TMS_BA_BIZSECTION_getMonthBizList]('"&FYYYYMM&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthBizList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'erp 부서리스트 -  월별 손익데이터 상세  업무비율 구분  있는 지원부서만
	public Function fnGetBizMonthUserBizList
	Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_user_Bizsection_getBizList]('"&FYYYYMM&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthUserBizList = rsget.getRows()
		END IF
		rsget.close
	End Function


	public Function fnGetManualBizList
	Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_userBizsection_avg_manualGetList]('"&FYYYYMM&"','"&FBizsection_cd&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetManualBizList = rsget.getRows()
		END IF
		rsget.close
	End Function
END Class

'//이익부서만 보여줄것인지 전체부서 보여줄것인지 계정과목으로 체크
'//input : 계정과목, 계정과목그룹코드
'//output : Y-이익부서만, N-전체
Function fnCheckBizSale(ByVal acc_use_cd, ByVal acc_grp_cd)
Dim blnSale : blnSale = "Y" '이익부서
	IF acc_grp_cd = "230" OR acc_use_cd = "13800" OR acc_use_cd ="93300" OR acc_use_cd ="21200" OR acc_use_cd ="14600" OR acc_use_cd="21900" OR acc_use_cd="25300" OR acc_use_cd="13410" OR acc_use_cd="81740" THEN '(230:판관비, 13800:전도금,93300:기부금,21200:비품,21900:시설장치(유형자산),13410 가지급금_동호회비, 81740 세금과공과)
	blnSale = "N" '전체표시
	END IF
	fnCheckBizSale = blnSale
End Function

function DrawBizSectionGain(itargetArr,compname,compVal,showTp)
    Dim retStr
    retStr = "<select name='"&compname&"'>"
    retStr = retStr & "<option value=''>전체"
    IF InStr(itargetArr,"O") then
        retStr = retStr & "<option value='0000000101' "&CHKIIF(compVal="0000000101","selected","")&">온라인"
    end if
    IF InStr(itargetArr,"F") then
        retStr = retStr & "<option value='0000000201' "&CHKIIF(compVal="0000000201","selected","")&">오프라인"
    end if
    IF InStr(itargetArr,"T") then
        retStr = retStr & "<option value='0000000301' "&CHKIIF(compVal="0000000301","selected","")&">아이띵소"
    end if
     IF InStr(itargetArr,"C") then
        retStr = retStr & "<option value='0000000401' "&CHKIIF(compVal="0000000401","selected","")&">아카데미"
    end if
    retStr = retStr & "</select>"

    response.write retStr
end function
%>
