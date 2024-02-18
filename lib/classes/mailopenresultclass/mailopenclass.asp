<%
'###########################################################
' Description :  메일 오픈율 관리 클래스
' History : 2007.08.27 한용민 생성
'			2012.05.09 김진영 수정
'###########################################################

Class CMailzineOne
	Public Ftitle						'메일제목
	Public Fstartdate
	Public Fenddate
	Public Freenddate					'발송날짜
	Public Ftotalcnt
	Public Frealcnt						'실제발송통수
	Public Frealpct						'실제발송통수(%)
	Public Ffilteringcnt
	Public Ffilteringpct
	Public Fsuccesscnt					'성공발송통수
	Public Fsuccesspct					'성공발송통수(%)
	Public Ffailcnt
	Public Ffailpct
	Public Fopencnt						'오픈통수
	Public Fopenpct						'오픈통수(%)
	Public Fnoopencnt
	Public Fnoopenpct
	Public fgubun

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End class

'mailzine 메일 발송 결과에 따른 클래스
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

	Public Function frecttot()				'년수과 달을 폼에서 받아와서 합침...
		frecttot = frectyyyy & "-" & frectmm
	End Function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'날짜값이 있다면
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'위에 쿼리에 검색 옵션을 붙인다
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine'"						'구분이 메일진인 것만 뽑아온다.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'레코드의 첫번째가 아니라면
			Do until rsget.eof					'레코드의 끝까지 루프 ㄱㄱ
				set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealcnt = rsget("realcnt")					'실제발송통수
					flist(i).Fsuccesscnt = rsget("successcnt")				'성공발송통수
					flist(i).Fopencnt = rsget("opencnt")					'오픈통수
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

'mailzine 그래프용 클래스
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

	'년수과 달을 폼에서 받아와서 합침...
	Public function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'날짜값이 있다면
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'위에 쿼리에 검색 옵션을 붙인다
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"     & vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																									& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine'"								'구분이 메일진인 것만 가져온다.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then						'레코드의 첫번째가 아니라면
			Do until rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				Set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealpct = rsget("realpct")					'실제발송통수(%)
					flist(i).Fsuccesspct = rsget("successpct")				'성공발송통수(%)
					flist(i).Fopenpct = rsget("openpct")					'오픈통수(%)
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

'Fingers 메일 발송 결과에 따른 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealcnt = rsget("realcnt")					'실제발송통수
					flist(i).Fsuccesscnt = rsget("successcnt")				'성공발송통수
					flist(i).Fopencnt = rsget("opencnt")					'오픈통수
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

'Fingers 그래프용 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
		If not rsget.eof Then					'레코드의 첫번째가 아니라면
			Do until rsget.eof					'레코드의 끝까지 루프 ㄱㄱ
				Set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealpct = rsget("realpct")					'실제발송통수(%)
					flist(i).Fsuccesspct = rsget("successpct")				'성공발송통수(%)
					flist(i).Fopenpct = rsget("openpct")					'오픈통수(%)
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

'mailzine_not 메일 발송 결과에 따른 클래스
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

	Public Function frecttot()				'년수과 달을 폼에서 받아와서 합침...
		frecttot = frectyyyy & "-" & frectmm
	End Function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'날짜값이 있다면
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'위에 쿼리에 검색 옵션을 붙인다
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realcnt,successcnt,opencnt,gubun"   	& vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																			& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine_not'"						'구분이 메일진인 것만 뽑아온다.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then					'레코드의 첫번째가 아니라면
			Do until rsget.eof					'레코드의 끝까지 루프 ㄱㄱ
				set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealcnt = rsget("realcnt")					'실제발송통수
					flist(i).Fsuccesscnt = rsget("successcnt")				'성공발송통수
					flist(i).Fopencnt = rsget("opencnt")					'오픈통수
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

'mailzine 그래프용 클래스
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

	'년수과 달을 폼에서 받아와서 합침...
	Public function frecttot()
		frecttot = frectyyyy & "-" & frectmm
	End function

	Public Function frecttotnew()
		If frectyyyy <>"" and frectmm <> "" Then													'날짜값이 있다면
			frecttotnew = " and left(convert(varchar(15),reenddate,121),7) = '" & frecttot & "'"	'위에 쿼리에 검색 옵션을 붙인다
		Else
			frecttotnew = " and reenddate = '" & 0 &"'"												'검색값이 업다면 기본값0을 주고, 검색 옵션을 붙인다.첫화면에서 모든데이터가 다뿌려지는것을 방지...
		End If
	End Function

	Public Sub FMailzinelist
	Dim sql , i
	sql = sql & "select title,left(convert(varchar(15),reenddate,121),10) as reenddate ,realpct,successpct,openpct,gubun"     & vbcrlf
	sql = sql & " from [db_log].[dbo].tbl_mailing_data with (readuncommitted)"																									& vbcrlf
	sql = sql & " where 1=1 "& frecttotnew &" and gubun = 'mailzine_not'"								'구분이 메일진인 것만 가져온다.
	rsget.open sql,dbget,1
	FTotalCount = rsget.recordcount
	Redim flist(FTotalCount)
		i = 0
		If not rsget.eof Then						'레코드의 첫번째가 아니라면
			Do until rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				Set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealpct = rsget("realpct")					'실제발송통수(%)
					flist(i).Fsuccesspct = rsget("successpct")				'성공발송통수(%)
					flist(i).Fopenpct = rsget("openpct")					'오픈통수(%)
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



'Fingers 메일 발송 결과에 따른 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealcnt = rsget("realcnt")					'실제발송통수
					flist(i).Fsuccesscnt = rsget("successcnt")				'성공발송통수
					flist(i).Fopencnt = rsget("opencnt")					'오픈통수
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

'Fingers 그래프용 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
		If not rsget.eof Then					'레코드의 첫번째가 아니라면
			Do until rsget.eof					'레코드의 끝까지 루프 ㄱㄱ
				Set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealpct = rsget("realpct")					'실제발송통수(%)
					flist(i).Fsuccesspct = rsget("successpct")				'성공발송통수(%)
					flist(i).Fopenpct = rsget("openpct")					'오픈통수(%)
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


'OFFLINE 메일 발송 결과에 따른 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealcnt = rsget("realcnt")					'실제발송통수
					flist(i).Fsuccesscnt = rsget("successcnt")				'성공발송통수
					flist(i).Fopencnt = rsget("opencnt")					'오픈통수
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

'OFFLINE 그래프용 클래스
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

	'년수과 달을 폼에서 받아와서 합침
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
		If not rsget.eof Then					'레코드의 첫번째가 아니라면
			Do until rsget.eof					'레코드의 끝까지 루프 ㄱㄱ
				Set flist(i) = new CMailzineOne 			'클래스를 넣고
					flist(i).Ftitle = rsget("title")						'메일제목
					flist(i).Freenddate = rsget("reenddate")				'발송날짜
					flist(i).Frealpct = rsget("realpct")					'실제발송통수(%)
					flist(i).Fsuccesspct = rsget("successpct")				'성공발송통수(%)
					flist(i).Fopenpct = rsget("openpct")					'오픈통수(%)
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
