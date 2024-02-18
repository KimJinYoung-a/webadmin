<%

class CWaitItemlist2014
	public FListType
	public Fcurrstate
	public FSort

	public FTotCnt
	public FSPageNo
	public FEPageNo
	public FPageSize
	public FCurrPage

	public Fcatecode
	public Fmakerid
	public Fitemname
 	public Fitemid

	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectctrState

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FTotCnt =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	'// ���δ���ǰ����Ʈ
	public Function fnGetSummaryList
		dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_getSummrayList]('"&FListType&"', '"&FcurrState&"','"&FSort&"','"&Fmakerid&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetSummaryList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//���δ�� ��ǰ �󼼸���Ʈ
	' /admin/itemmaster/item_confirm.asp
	public Function fnGetWaitItemList
		Dim strSql

		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_getItemListCnt] '"&Fcatecode&"','"&Fmakerid&"','"&Fitemname&"','"&Fcurrstate&"','"&FItemid&"', '" + CStr(FRectCate_Large) + "', '" + CStr(FRectCate_Mid) + "', '" + CStr(FRectCate_Small) + "','"&FRectctrState&"'"

		'response.write strSql & "<Br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_temp].[dbo].sp_Ten_wait_item_getItemList '"&Fcatecode&"','"&Fmakerid&"','"&Fitemname&"','"&Fcurrstate&"','"&FSort&"','"&FItemid&"',"&FSPageNo&","&FEPageNo&", '" + CStr(FRectCate_Large) + "', '" + CStr(FRectCate_Mid) + "', '" + CStr(FRectCate_Small) + "','"&FRectctrState&"'"

		'response.write strSql & "<Br>"		
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	'//���� �������� �α�
	public Function fnGetWaitItemLog
		Dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_log_getItemList]("&Fitemid&")"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemLog = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//API �������� �α�
	public Function fnGetWaitItemApiLog
		Dim strSql
		strSql ="[db_temp].[dbo].[sp_Ten_wait_item_log_getItemList]("&Fitemid&", 'Y')"
	 	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetWaitItemApiLog = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetOldWaitItemLog
		Dim strSql
		strSql =" select  rejectdate, rejectmsg, reregdate, reregmsg, currstate From db_temp.dbo.tbl_wait_item where itemid =" &Fitemid&" and currstate in (2, 0, 5 ) "
		rsget.Open strSql,dbget,1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetOldWaitItemLog = rsget.getRows()
		END IF
		rsget.close
	End Function
end Class


'//������� �Լ�
Sub sbOptItemWaitStatus(currstate)
	%>
	<option value="A" <%IF currstate="A" THEN%>selected<%END IF%>>��ü</option>
	<option value="1" <%IF currstate="1" THEN%>selected<%END IF%>>���δ��</option>
	<option value="5" <%IF currstate="5" THEN%>selected<%END IF%>>���δ�� (����)</option>
	<option value="2" <%IF currstate="2" THEN%>selected<%END IF%>>���κ��� (���Ͽ�û)</option>
	<option value="0" <%IF currstate="0" THEN%>selected<%END IF%>>���ιݷ� (���ϺҰ�)</option>
	<option value="7" <%IF currstate="7" THEN%>selected<%END IF%>>���οϷ�</option>
	<%
End Sub

	function GetCurrStateColor(ByVal FCurrState)
		GetCurrStateColor = "#000000"
		if FCurrState="1" then
			GetCurrStateColor = "#000000"
		elseif FCurrState="2" then
			GetCurrStateColor = "#FF0000"
		elseif FCurrState="3" then
			GetCurrStateColor = "#DD0000"
		elseif FCurrState="4" then
			GetCurrStateColor = "#DD0000"
		elseif FCurrState="7" then
			GetCurrStateColor = "#0000FF"
		elseif FCurrState="5" then
			GetCurrStateColor = "#008800"
		else
			GetCurrStateColor = "#000000"
		end if
	end function

 function GetCurrStateName(ByVal FCurrState)
		GetCurrStateName = ""
		if FCurrState="1" then
			GetCurrStateName = "���δ��"
		elseif FCurrState="2" then
			GetCurrStateName = "���κ���<Br>(���Ͽ�û)"
		elseif FCurrState="3" then
			GetCurrStateName = "ó�����<Br>(���Ͽ�û)"
		elseif FCurrState="4" then
			GetCurrStateName = "ó������<Br>(���Ͽ�û)"
		elseif FCurrState="7" then
			GetCurrStateName = "���οϷ�"
		elseif FCurrState="5" then
			GetCurrStateName = "���δ��<Br>(����)"
		elseif FCurrState="0" then
			GetCurrStateName = "���ιݷ�<Br>(���ϺҰ�)"
		elseif FCurrState="9" then
			GetCurrStateName = "��ü����"
		else
			GetCurrStateName = ""
		end if
	end function

	 function GetCurrStateContsName(ByVal FCurrState)
		GetCurrStateContsName = ""
		if FCurrState="1" then
			GetCurrStateContsName = "���δ��"
		elseif FCurrState="2" then
			GetCurrStateContsName = "���κ���(���Ͽ�û)"
		elseif FCurrState="7" then
			GetCurrStateContsName = "���οϷ�"
		elseif FCurrState="5" then
			GetCurrStateContsName = "���δ��(����)"
		elseif FCurrState="0" then
			GetCurrStateContsName = "���ιݷ�(���ϺҰ�)"
		elseif FCurrState="9" then
			GetCurrStateContsName = "��ü����"
		else
			GetCurrStateContsName = ""
		end if
	end function

	function fnGetCurrStateShortName(ByVal FCurrState)
			fnGetCurrStateShortName = ""
		if FCurrState="1" then
			fnGetCurrStateShortName = "���"
		elseif FCurrState="2" then
			fnGetCurrStateShortName = "����"
		elseif FCurrState="7" then
			fnGetCurrStateShortName = "�Ϸ�"
		elseif FCurrState="5" then
			fnGetCurrStateShortName = "����"
		elseif FCurrState="0" then
			fnGetCurrStateShortName = "�ݷ�"
		elseif FCurrState="9" then
			fnGetCurrStateShortName = "����"
		elseif FCurrState="S" then
			fnGetCurrStateShortName = "��������ó��"
		elseif FCurrState="I" then
			fnGetCurrStateShortName = "�̹���ó��"
		else
			fnGetCurrStateShortName = ""
		end if
	End Function
%>
