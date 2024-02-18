<%
'###########################################################
' Description :  ���� ������ ���� Ŭ����
' History : 2012.12.04 ������ ����
'			2020.05.15 �ѿ�� ����
'###########################################################

Class CMailzineOne
	Public FTitle						'��������
	Public FStartdate
	Public FEnddate
	Public FReenddate					'�߼۳�¥
	Public FTotalcnt
	Public FRealcnt						'�����߼����
	Public FRealpct						'�����߼����(%)
	Public FFilteringcnt
	Public FFilteringpct
	Public FSuccesscnt					'�����߼����
	Public FSuccesspct					'�����߼����(%)
	Public FFailcnt
	Public FFailpct
	Public FOpencnt						'�������
	Public FOpenpct						'�������(%)
	Public FClickCnt					'Ŭ�����
	Public FClickPct					'Ŭ������(%)
	Public FNoopencnt
	Public FNoopenpct
	Public fGubun
	public fmailergubun

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End class

Class CMailzinelist
	Public Flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public FRectyyyy
	Public FRectmm
	Public FGubun

	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then
			frecttotnew = " AND left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " AND reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
		Dim strSql, i

		strSql = "SELECT title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,clickcnt,gubun,mailergubun" & VBCRLF
		strSql = strSql & " FROM [db_log].[dbo].tbl_mailing_data with (readuncommitted)" & VBCRLF
		strSql = strSql & " WHERE 1=1 "& frecttotnew &" AND gubun = '"&FGubun&"' "

		'response.write strSql & "<Br>"
		rsget.open strSql,dbget,1
		FTotalCount = rsget.recordcount
	
		Redim flist(FTotalCount)
		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				Set FList(i) = new CMailzineOne
					FList(i).Ftitle = rsget("title")
					FList(i).Freenddate = rsget("reenddate")
					FList(i).Frealcnt = rsget("realcnt")
					FList(i).Fsuccesscnt = rsget("successcnt")
					FList(i).Fopencnt = rsget("opencnt")
					FList(i).FClickCnt = rsget("clickcnt")
					FList(i).fgubun = rsget("gubun")
					FList(i).fmailergubun = rsget("mailergubun")
					rsget.movenext
				i = i + 1
			Loop
		End If
		rsget.close
	End Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

Class CMailzinelistgraph
	Public FList
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public Frectyyyy
	Public Frectmm
	Public FGubun

	Public function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then
			frecttotnew = " AND left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " AND reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
		Dim strSql , i

		strSql = "SELECT title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,clickpct,gubun" & VBCRLF
		strSql = strSql & " FROM [db_log].[dbo].tbl_mailing_data with (readuncommitted)" & VBCRLF
		strSql = strSql & " WHERE 1=1 "& frecttotnew &" AND gubun = '"&FGubun&"'"

		'response.write strSql & "<Br>"
		rsget.open strSql, dbget, 1
		FTotalCount = rsget.RecordCount
		Redim FList(FTotalCount)
			i = 0
			If not rsget.EOF Then
				Do until rsget.EOF
					Set FList(i) = new CMailzineOne
						FList(i).Ftitle = rsget("title")
						FList(i).Freenddate = rsget("reenddate")
						FList(i).Frealpct = rsget("realpct")
						FList(i).Fsuccesspct = rsget("successpct")
						FList(i).Fopenpct = rsget("openpct")
						FList(i).FClickPct = rsget("clickpct")
						FList(i).fgubun = rsget("gubun")
						rsget.movenext
					i = i + 1
				Loop
			End If
			rsget.close
	End Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class
%>