<%
'###########################################################
' Description :  ���� ������ ���� Ŭ����
' History : 2007.08.27 �ѿ�� ����
'			2012.05.09 ������ ����
'###########################################################

Class CMailzineOne
	Public Ftitle						'��������
	Public Fstartdate
	Public Fenddate
	Public Freenddate					'�߼۳�¥
	Public Ftotalcnt
	Public Frealcnt						'�����߼����
	Public Frealpct						'�����߼����(%)
	Public Ffilteringcnt
	Public Ffilteringpct
	Public Fsuccesscnt					'�����߼����
	Public Fsuccesspct					'�����߼����(%)
	Public Ffailcnt
	Public Ffailpct
	Public Fopencnt						'�������
	Public Fopenpct						'�������(%)
	Public Fnoopencnt
	Public Fnoopenpct
	Public fgubun

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End class

'mailzine ���� �߼� ����� ���� Ŭ����
Class CMailzinelist
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	Public Function frecttot()				'����� ���� ������ �޾ƿͼ� ��ħ...
		frecttot = frectyyyy & "-" & frectmm
	End Function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'��¥���� �ִٸ�
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'���� ������ �˻� �ɼ��� ���δ�
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'�˻����� ���ٸ� �⺻��0�� �ְ�, �˻� �ɼ��� ���δ�.ùȭ�鿡�� ��絥���Ͱ� �ٻѷ����°��� ����...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine'"						'������ �������� �͸� �̾ƿ´�.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof					'���ڵ��� ������ ���� ����
				set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealcnt = rsget("realcnt")					'�����߼����
					flist(i).Fsuccesscnt = rsget("successcnt")				'�����߼����
					flist(i).Fopencnt = rsget("opencnt")					'�������
					flist(i).fgubun = rsget("gubun")
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
End class

'mailzine �׷����� Ŭ����
Class CMailzinelistgraph1
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ...
	Public function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'��¥���� �ִٸ�
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'���� ������ �˻� �ɼ��� ���δ�
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'�˻����� ���ٸ� �⺻��0�� �ְ�, �˻� �ɼ��� ���δ�.ùȭ�鿡�� ��絥���Ͱ� �ٻѷ����°��� ����...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"     & vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																									& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine'"								'������ �������� �͸� �����´�.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then						'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof						'���ڵ��� ������ ���� ����
				Set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealpct = rsget("realpct")					'�����߼����(%)
					flist(i).Fsuccesspct = rsget("successpct")				'�����߼����(%)
					flist(i).Fopenpct = rsget("openpct")					'�������(%)
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			Loop
		End If
		rsget.close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'Fingers ���� �߼� ����� ���� Ŭ����
Class Cacademyfingerslist
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"    	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'fingers'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof then
			Do until rsget.eof
				Set flist(i) = new CMailzineOne
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealcnt = rsget("realcnt")					'�����߼����
					flist(i).Fsuccesscnt = rsget("successcnt")				'�����߼����
					flist(i).Fopencnt = rsget("opencnt")					'�������
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			loop
		End If
		rsget.close
	End Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'Fingers �׷����� Ŭ����
Class Cacademyfingerslistgraph1
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'fingers'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof					'���ڵ��� ������ ���� ����
				Set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealpct = rsget("realpct")					'�����߼����(%)
					flist(i).Fsuccesspct = rsget("successpct")				'�����߼����(%)
					flist(i).Fopenpct = rsget("openpct")					'�������(%)
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			Loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'mailzine_not ���� �߼� ����� ���� Ŭ����
Class CMailzinelist_not
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	Public Function frecttot()				'����� ���� ������ �޾ƿͼ� ��ħ...
		frecttot = frectyyyy & "-" & frectmm
	End Function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'��¥���� �ִٸ�
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'���� ������ �˻� �ɼ��� ���δ�
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'�˻����� ���ٸ� �⺻��0�� �ְ�, �˻� �ɼ��� ���δ�.ùȭ�鿡�� ��絥���Ͱ� �ٻѷ����°��� ����...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine_not'"						'������ �������� �͸� �̾ƿ´�.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof					'���ڵ��� ������ ���� ����
				set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealcnt = rsget("realcnt")					'�����߼����
					flist(i).Fsuccesscnt = rsget("successcnt")				'�����߼����
					flist(i).Fopencnt = rsget("opencnt")					'�������
					flist(i).fgubun = rsget("gubun")
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
End class

'mailzine �׷����� Ŭ����
Class CMailzinelistgraph1_not
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ...
	Public function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'��¥���� �ִٸ�
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'���� ������ �˻� �ɼ��� ���δ�
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'�˻����� ���ٸ� �⺻��0�� �ְ�, �˻� �ɼ��� ���δ�.ùȭ�鿡�� ��絥���Ͱ� �ٻѷ����°��� ����...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"     & vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																									& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine_not'"								'������ �������� �͸� �����´�.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then						'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof						'���ڵ��� ������ ���� ����
				Set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealpct = rsget("realpct")					'�����߼����(%)
					flist(i).Fsuccesspct = rsget("successpct")				'�����߼����(%)
					flist(i).Fopenpct = rsget("openpct")					'�������(%)
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			Loop
		End If
		rsget.close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class



'Fingers ���� �߼� ����� ���� Ŭ����
Class Cacademyfingerslist_not
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"    	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'fingers_not'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof then
			Do until rsget.eof
				Set flist(i) = new CMailzineOne
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealcnt = rsget("realcnt")					'�����߼����
					flist(i).Fsuccesscnt = rsget("successcnt")				'�����߼����
					flist(i).Fopencnt = rsget("opencnt")					'�������
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			loop
		End If
		rsget.close
	End Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'Fingers �׷����� Ŭ����
Class Cacademyfingerslistgraph1_not
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'fingers_not'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof					'���ڵ��� ������ ���� ����
				Set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealpct = rsget("realpct")					'�����߼����(%)
					flist(i).Fsuccesspct = rsget("successpct")				'�����߼����(%)
					flist(i).Fopenpct = rsget("openpct")					'�������(%)
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			Loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class


'OFFLINE ���� �߼� ����� ���� Ŭ����
Class COFFLINElist
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"    	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'OFFLINE'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof then
			Do until rsget.eof
				Set flist(i) = new CMailzineOne
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealcnt = rsget("realcnt")					'�����߼����
					flist(i).Fsuccesscnt = rsget("successcnt")				'�����߼����
					flist(i).Fopencnt = rsget("opencnt")					'�������
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			loop
		End If
		rsget.close
	End Sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class

'OFFLINE �׷����� Ŭ����
Class COFFLINElistgraph1
	Public flist
	Public FCurrPage
	Public FPageSize
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	Public FTotalPage
	Public frectyyyy
	Public frectmm

	'����� ���� ������ �޾ƿͼ� ��ħ
	Public Function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End Function

	public function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" then
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'OFFLINE'"
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'���ڵ��� ù��°�� �ƴ϶��
			Do until rsget.eof					'���ڵ��� ������ ���� ����
				Set flist(i) = new CMailzineOne 			'Ŭ������ �ְ�
					flist(i).Ftitle = rsget("title")						'��������
					flist(i).Freenddate = rsget("reenddate")				'�߼۳�¥
					flist(i).Frealpct = rsget("realpct")					'�����߼����(%)
					flist(i).Fsuccesspct = rsget("successpct")				'�����߼����(%)
					flist(i).Fopenpct = rsget("openpct")					'�������(%)
					flist(i).fgubun = rsget("gubun")
					rsget.movenext
				i = i + 1
			Loop
		end if
		rsget.close
	end sub

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class
%>
