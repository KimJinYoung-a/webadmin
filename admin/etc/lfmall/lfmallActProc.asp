<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/lfmall/lfmallCls.asp"-->
<!-- #include virtual="/admin/etc/lfmall/inclfmallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oLfmall, failCnt, chgSellYn, arrRows, isItemIdChk, mustPrice
Dim iErrStr, strSql, SumErrStr, SumOKStr, i, strparam, mrgnRate, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

If action = "REG" Then									'��ǰ���
	SET oLfmall = new CLfmall
		oLfmall.FRectItemID	= itemid
		oLfmall.getLfmallNotRegOneItem
	    If (oLfmall.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||��ϰ����� ��ǰ�� �ƴմϴ�."
			Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Lfmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			If oLfmall.FOneItem.checkTenItemOptionValid Then
				Call fnLfmallItemReg(itemid, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[��ǰ���] �ɼǰ˻� ����"
				Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
			End If
		End If
	SET oLfmall = nothing
ElseIf action = "EditSellYn" Then						'���º���
	SET oLfmall = new CLfmall
		oLfmall.FRectItemID	= itemid
		oLfmall.getLfmallEditOneItem
		If oLfmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���¼��� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		Else
			Call fnLfmallSellYN(itemid, chgSellYn, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				Call fnLfmallStatChk(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		End If
	SET oLfmall = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("lfmall", itemid, SumErrStr)
		Call SugiQueLogInsert("lfmall", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lfmall_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("lfmall", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "CHKSTAT" Then							'��ǰ��ȸ
	SET oLfmall = new CLfmall
		oLfmall.FRectItemID	= itemid
		oLfmall.getLfmallEditOneItem
		If oLfmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||��ǰ��ȸ �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		Else
			Call fnLfmallStatChk(itemid, iErrStr)
		End If
	SET oLfmall = nothing
ElseIf action = "EDITITEM" Then								'��ǰ����
	SET oLfmall = new CLfmall
		oLfmall.FRectItemID	= itemid
		oLfmall.getLfmallEditOneItem
		If oLfmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		Else
			Call fnLfmallItemEdit(itemid, iErrStr)
		End If
	SET oLfmall = nothing
ElseIf action = "EDIT" Then								'��ǰ ����
	SET oLfmall = new CLfmall
		oLfmall.FRectItemID	= itemid
		oLfmall.getLfmallEditOneItem
		If oLfmall.FResultCount = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||���� �� ��ǰ�� ��ϵǾ� ���� �ʽ��ϴ�."
			Call SugiQueLogInsert("lfmall", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		Else
            If (oLfmall.FOneItem.FmaySoldOut = "Y") OR (oLfmall.FOneItem.IsMayLimitSoldout = "Y") Then
				Call fnLfmallSellYN(itemid, "N", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
			'############## Lfmall ��ǰ ���� #################
				Call fnLfmallItemEdit(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			'############## Lfmall ���� ���� #################
				If failCnt = 0 Then
					Call fnLfmallSellYN(itemid, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			'############## Lfmall �� ��ȸ #################
				If failCnt = 0 Then
					Call fnLfmallStatChk(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If
		End If
	SET oLfmall = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("lfmall", itemid, SumErrStr)
		Call SugiQueLogInsert("lfmall", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lfmall_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("lfmall", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
End If


response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&iErrStr&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->