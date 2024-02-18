<%
Class COpExpPart
public FPartTypeidx
public FOpExpPartidx
public FOpExpPartName
public FPartTypeName
public FIsUsing
public FadminID
public Fusername
public Fpart_sn
public Fjob_sn
public Fjobname
public FOutBank
public FOutBankName
public FOutAccNo
public FOutName

public FOrderNo
public FRectUserid
public FRectPartsn
public FRectDepartmentID

public FSPageNo
public FEPageNo
public FPageSize
public FCurrPage
public FTotCnt

public Fbizsection_cd
public Fbizsection_nm
public Farap_cd
public FARAP_NM
public Fcust_cd
public Fcust_nm
public FCardCo
public FCardNo

Public FRectIncNo

	'운영비 운영부서  리스트 가져오기
	public Function fnGetOpExpPartList
		IF FPartTypeidx = "" THEN FPartTypeidx = 0
		Dim strSql
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpPart_getListCnt]("&FPartTypeidx&",'"&FOpExpPartName&"', '" + CStr(FRectIncNo) + "')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_getList("&FPartTypeidx&",'"&FOpExpPartName&"',"&FSPageNo&","&FEPageNo&", '" + CStr(FRectIncNo) + "')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpPartList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'운영비사용처구분 리스트
	public Function fnGetOpExpPartTypeList
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPartType_getList('"&FRectUserid&"',"&FRectPartsn&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpPartTypeList = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetOpExpPartTypeListNew
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPartType_getListNew('"&FRectUserid&"','"&FRectDepartmentID&"','" & FIsUsing & "')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpPartTypeListNew = rsget.getRows()
		END IF
		rsget.close
	End Function

	'운영비카드사용처구분 리스트
	 public Function fnGetOpExpPartTypeCardList
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpPartType_getCardList]('"&FRectUserid&"',"&FRectPartsn&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpPartTypeCardList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'// 부서NEW 사용
	 public Function fnGetOpExpPartTypeCardListNew
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].[sp_Ten_OpExpPartType_getCardListNew]('"&FRectUserid&"','"&FRectDepartmentID&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpPartTypeCardListNew = rsget.getRows()
		END IF
		rsget.close
	End Function

		'운영비 팀 전체 리스트
	public Function fnGetOpExppartAllList
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_getOptList("&FPartTypeidx&",'"&FRectUserid&"',"&FRectPartsn&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExppartAllList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'// 부서NEW 사용
	public Function fnGetOpExppartAllListNew
		Dim strSql
		IF FRectPartsn = "" THEN FRectPartsn = 0
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_getOptListNew("&FPartTypeidx&",'"&FRectUserid&"','"&FRectDepartmentID&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExppartAllListNew = rsget.getRows()
		END IF
		rsget.close
	End Function

	'운영비 사용처 번호 가져오기
	public Function fnGetOpExpPart
	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_GetIdx('"&FCardCo&"','"&FCardNo&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FOpExpPartIdx = rsget(0)
		END IF
	rsget.close
  End Function

	'운영비관리 팀 내용 가져오기
	public Function fnGetOpExpPartData
	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_getData("&FOpExpPartidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		  FPartTypeIdx 	= rsget("PartTypeIdx")
		  FPartTypeName = rsget("PartTypeName")
		  FOpExpPartName= rsget("OpExpPartName")
		  FOutBank		= rsget("OutBank")
		  FOutAccNo		= rsget("OutAccNo")
		  FOutName		= rsget("OutName")
		  Fbizsection_cd	= rsget("BIZSECTION_CD")
		  Fbizsection_nm	= rsget("bizsection_nm")
		  FARAP_cd		= rsget("ARAP_cd")
		  FARAP_NM		= rsget("ARAP_NM")
		  FOrderNo		= rsget("OrderNo")
		  FIsUsing 		= rsget("IsUsing")
		  FadminID		= rsget("adminID")
		  Fusername		= rsget("UserName")
		  Fpart_sn		= rsget("part_sn")
		  Fjob_sn			= rsget("job_sn")
		  Fjobname		= rsget("job_name")
		  Fcust_cd		= rsget("cust_cd")
		  Fcust_nm		= rsget("cust_nm")
		  Fcardco			= rsget("cardCo")
		  FcardNo			= rsget("cardNo")
		END IF
		rsget.close
	End Function

	'구분 데이터
	public Function fnGetOpExpPartTypeData
	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPartType_getData("&FPartTypeIdx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
		  FPartTypeIdx 	= rsget("PartTypeIdx")
		  FPartTypeName = rsget("PartTypeName")
		  FIsUsing 		= rsget("IsUsing")
		END IF
		rsget.close
	End Function


	'운영비 관리 팀 연동 부서 리스트 가져오기
	public Function fnGetOpExppartInfoList
		Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPartInfo_getList("&FOpExpPartidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExppartInfoList = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetOpExpDepartmentInfoList
		Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpDepartmentInfo_getList("&FOpExpPartidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOpExpDepartmentInfoList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'운영비 팀 이름 가져오기
	public Function fnGetOpExpPartName
	Dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_OpExpPart_getName("&FOpExpPartidx&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FOpExpPartName	= rsget("OpExpPartName")
			FPartTypeName	= rsget("PartTypeName")
			'FeappPartIdx	= rsget("eappPartIdx")
		END IF
		rsget.close

	End Function
End Class

'운영비 운영구분
	public Sub sbOptPartType(arrList,iValue)
		Dim  intLoop
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(iValue)=Cstr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
			<%
			Next
		END IF
	End Sub

' 운영비 팀 전체리스트
	public Sub sbOptPart(arrList,iValue)
		Dim  intLoop
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			%>
		<option value="<%=arrList(0,intLoop)%>" <%IF Cstr(iValue)=Cstr(arrList(0,intLoop)) THEN%>selected<%END IF%>><%=arrList(1,intLoop)%></option>
			<%
			Next
		END IF
	End Sub
%>
