<%
'############################
'ERP ���� �ŷ�ó ����Ʈ
'############################

Class CCust
public FCUSTgbn	'�ŷ�ó����
public FCUSTtype	'�ŷ�ó�з�
public FSearchType	'�˻����
public FSearchText	'�˻���
PUBLIC FARAP_TYPE	'�Ա�/���� ����
public FCUSTBRNTYPE
public FRectAllacct ''�⺻���¿���

public FRectEmpno
public FRectBankNo
public FRectAcctNo

public FTotCnt
public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage

public FCustCD
public FCORPYN
public FARYN
public FAPYN
public FCustNM
public FBizNo
public FCeoNM
public FEMAIL
public FTELNO
public FFAXNO
public FTAXTYPE
public FBSCD
public FINTP
public FPostCD
public FADDR
public FDispSeq

public FEMP_NO
public FEMP_NM
public FPos
public FDEPT_NM
public FSTelNo
public FHP_NO
public FSEmail
public FBank_cd
public Facct_no
public Fsav_mn
public FPSGB

	public Function fnGetCustList
		Dim strSql
		IF  FCUSTgbn = "" THEN  FCUSTgbn = 0
		IF  FSearchType = "" THEN  FSearchType = 0
		strSql = "db_partner.dbo.sp_Ten_TMS_BA_CUST_getListCnt("&FCUSTgbn&",'"&FCUSTtype&"','"&FARAP_TYPE&"',"&FSearchType&",'"&FSearchText&"',"&CHKIIF(FRectAllacct="on","1","0")&")"

		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_CUST_getList("&FCUSTgbn&",'"&FCUSTtype&"','"&FARAP_TYPE&"',"&FSearchType&",'"&FSearchText&"',"&CHKIIF(FRectAllacct="on","1","0")&","&FSPageNo&","&FEPageNo&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCustList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

 public Function fnGetCustData
 Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_CUST_getData('"&FCustCD&"','"&FRectEmpno&"','"&FRectBankNo&"','"&FRectAcctNo&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			 FCustCD = rsget("cust_cd")
			 FCUSTBRNTYPE = rsget("CUST_BRN_TYPE")
			 FCORPYN 	= rsget("CORP_YN")
			 FARYN	 	= rsget("CUST_AR_YN")
			 FAPYN	 	= rsget("CUST_AP_YN")
			 FCustNM 	= rsget("CUST_NM")
			 FBizNo 	= rsget("BIZ_NO")
			 FCeoNM 	= rsget("CEO_NM")
			 FEMAIL 	= rsget("EMAIL")
			 FTELNO 	= rsget("TEL_NO")
			 FFAXNO 	= rsget("FAX_NO")
			 FTAXTYPE = rsget("TAX_TYPE")
			 FBSCD 		= rsget("BSCD")
			 FINTP 		= rsget("INTP")
			 FPostCD 	= rsget("Post_CD")
			 FADDR 		= rsget("ADDR")
			 FDispSeq = rsget("Disp_Seq")
			 FEMP_NO	= rsget("EMP_NO")
			 FEMP_NM	= rsget("EMP_NM")
			 FPos			= rsget("Pos")
			 FDEPT_NM	= rsget("DEPT_NM")
			 FSTelNo		= rsget("STelNo")
			 FHP_NO		= rsget("HP_NO")
			 FSEmail 	= rsget("SEmail")
			 FBank_cd	= rsget("Bank_cd")
			 Facct_no = rsget("acct_no")
			 Fsav_mn	= rsget("sav_mn")
			 FPSGB		= rsget("PERSON_SITE_GB")
		END IF
		rsget.close
	End Function

public Function fnGetCustSaleorList
 	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_CUST_SALEOR_getList('"&FCustCD&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCustSaleorList = rsget.getRows()
		END IF
		rsget.close
End Function

public Function fnGetCustAcctList
 	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_CUST_ACCT_getList('"&FCustCD&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetCustAcctList = rsget.getRows()
		END IF
		rsget.close
End Function

	'//����� ����Ʈ ��������
	public Function fnGetBankList
	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_TMS_BA_COM_CD_getList"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetBankList = rsget.getRows()
		END IF
		rsget.close
	End Function
END Class

'//�ŷ�ó �з�
Function fnGetCustTypeName(ByVal CustType)
 Dim strTypeName
	IF CustType = "1" THEN
		strTypeName = "����"
	ELSEIF CustType = "4" THEN
		strTypeName = "��Ÿ(�� DuZon)"
	ELSEIF CustType = "5" THEN
		strTypeName = "�¶��ΰŷ�ó"
	ELSEIF CustType = "7" THEN
		strTypeName = "����/���/��ȣȸ"
	ELSEIF CustType = "0" THEN
		strTypeName = "����ŷ�ó"
	ELSEIF CustType = "9" THEN
		strTypeName = "�Һ��ڸ���ŷ�ó"
	ELSEIF IsNull(CustType) THEN
		strTypeName = "NULL"
	ELSE
		strTypeName = CStr(CustType)
	END IF
	fnGetCustTypeName = strTypeName
End Function

sub sbOptCustType(ByVal sCUSTtype)
%>
	<option value="0" <%IF sCUSTtype="0" THEN%>selected<%END IF%>>����ŷ�ó</option>
	<option value="7" <%IF sCUSTtype="7" THEN%>selected<%END IF%>>����/���/��ȣȸ</option>
	<option value="4" <%IF sCUSTtype="4" THEN%>selected<%END IF%>>��Ÿ(�� DuZon)</option>
	<option value="5" <%IF sCUSTtype="5" THEN%>selected<%END IF%>>�¶��ΰŷ�ó</option>
	<option value="9" <%IF sCUSTtype="9" THEN%>selected<%END IF%>>�Һ��ڸ���ŷ�ó</option>
<%
End Sub
%>
