<%
'############################
' Description : ERP ���� �μ�����
' History : 2011.04.21 ������  ����
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

	'erp �μ�����Ʈ
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


'erp ����� �μ� ��ü����Ʈ
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



'erp �μ�����Ʈ - ���� ���͵����� �ִ� �μ���
	public Function fnGetBizMonthProftist
		Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_TMS_BA_BIZSECTION_getMonthProfitList]('"&FGRP_YN&"','"&FYYYYMM&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthProftist = rsget.getRows()
		END IF
		rsget.close
	End Function

'erp �μ�����Ʈ - ���� ���͵����� �������� ����  �ִ� �μ���
	public Function fnGetBizMonthBizList
		Dim strSql
		strSql = "db_partner.dbo.[sp_Ten_TMS_BA_BIZSECTION_getMonthBizList]('"&FYYYYMM&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBizMonthBizList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'erp �μ�����Ʈ -  ���� ���͵����� ��  �������� ����  �ִ� �����μ���
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

'//���ͺμ��� �����ٰ����� ��ü�μ� �����ٰ����� ������������ üũ
'//input : ��������, ��������׷��ڵ�
'//output : Y-���ͺμ���, N-��ü
Function fnCheckBizSale(ByVal acc_use_cd, ByVal acc_grp_cd)
Dim blnSale : blnSale = "Y" '���ͺμ�
	IF acc_grp_cd = "230" OR acc_use_cd = "13800" OR acc_use_cd ="93300" OR acc_use_cd ="21200" OR acc_use_cd ="14600" OR acc_use_cd="21900" OR acc_use_cd="25300" OR acc_use_cd="13410" OR acc_use_cd="81740" THEN '(230:�ǰ���, 13800:������,93300:��α�,21200:��ǰ,21900:�ü���ġ(�����ڻ�),13410 �����ޱ�_��ȣȸ��, 81740 ���ݰ�����)
	blnSale = "N" '��üǥ��
	END IF
	fnCheckBizSale = blnSale
End Function

function DrawBizSectionGain(itargetArr,compname,compVal,showTp)
    Dim retStr
    retStr = "<select name='"&compname&"'>"
    retStr = retStr & "<option value=''>��ü"
    IF InStr(itargetArr,"O") then
        retStr = retStr & "<option value='0000000101' "&CHKIIF(compVal="0000000101","selected","")&">�¶���"
    end if
    IF InStr(itargetArr,"F") then
        retStr = retStr & "<option value='0000000201' "&CHKIIF(compVal="0000000201","selected","")&">��������"
    end if
    IF InStr(itargetArr,"T") then
        retStr = retStr & "<option value='0000000301' "&CHKIIF(compVal="0000000301","selected","")&">���̶��"
    end if
     IF InStr(itargetArr,"C") then
        retStr = retStr & "<option value='0000000401' "&CHKIIF(compVal="0000000401","selected","")&">��ī����"
    end if
    retStr = retStr & "</select>"

    response.write retStr
end function
%>
