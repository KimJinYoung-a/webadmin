<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	'### 보너스 코인 : gubuncd 가 13
	Dim sql, tmp, userid, userid_m, contents, savecoin, org_savecoin, vError, vIdx, vInsertGubun, vTemp, i, vAlert
	vIdx = Request("idx")
	userid = Request("giveuserid")
	contents = requestCheckVar(Request("givecontent"),100)
	savecoin = Request("givecoin")
	org_savecoin = savecoin	'### - 코인이 현재 코인보다 클 경우 계산하기 위해 따로 변수에 하나 저장.
	vInsertGubun = Request("insertgubun")
	userid_m = Request("giveuserid_m")
	vError = "x"

	
	'tozzinet
	On Error Resume Next
	dbget.beginTrans

	
	'####### 아이디 한개일 경우 #######
	If vInsertGubun = "one" Then
			sql = ""
			'//코인 내용 존재여부 확인	
			If Err.Number = 0 Then
			        vError = "1"
			end if
				
			sql = "select savecoin, currentcoin from db_momo.dbo.tbl_coin_current where userid = '"&userid&"' and isusing='Y'"
			
			'response.write sql &"<br>"
			rsget.open sql,dbget,1
				if not(rsget.bof or rsget.eof) then
					tmp = 1
					IF Left(savecoin,1) = "-" Then
						If (CDbl(savecoin) + (CDbl(rsget("savecoin")) - CDbl(rsget("currentcoin")))) < 0 Then
							savecoin = CDbl("-" & (CDbl(rsget("savecoin")) - CDbl(rsget("currentcoin"))))
							vAlert = "" & userid & " 고객님은 기존 코인이 " & (CDbl(rsget("savecoin")) - CDbl(rsget("currentcoin"))) & " 여서 - 가 되므로 0 으로 처리합니다.\n\n"
						End IF
					End If
				end if
			rsget.close

			'//코인 저장	
			sql = ""
			if tmp > 0 then
				
				If vIdx = "" Then
					If Err.Number = 0 Then
					        vError = "2"
					end if
					
					sql = ""
					sql = "update db_momo.dbo.tbl_coin_current set" + vbcrlf
					IF Left(savecoin,1) = "-" Then
						sql = sql & " currentcoin = currentcoin + "& Replace(savecoin,"-","")&" , lastupdate=getdate()" + vbcrlf
					Else
						sql = sql & " savecoin = savecoin + "&savecoin&" , lastupdate=getdate()" + vbcrlf
					End If
					sql = sql & " where userid = '"&userid&"' and isusing='Y'" + vbcrlf
					
					dbget.execute sql
				End IF
				
				
				If vIdx = "" Then
					If Err.Number = 0 Then
					        vError = "3"
					end if
					
					'//코인 로그 저장
					sql = ""
					sql = "insert into db_momo.dbo.tbl_coin_log" + vbcrlf
					sql = sql & " (userid,coin,gubuncd,gubun,deleteyn) values" + vbcrlf
					sql = sql & " (" + vbcrlf		
					sql = sql & " '" & userid & "'" + vbcrlf
					sql = sql & " , '" & savecoin & "' " + vbcrlf
					sql = sql & " , '13' " + vbcrlf
					sql = sql & " , '" & contents & "' " + vbcrlf		
					sql = sql & " , 'N' " + vbcrlf	
					sql = sql & " )" + vbcrlf
							
					dbget.execute sql
			'		response.write sql &"<br>"
			'		dbget.RollBackTrans
			'		response.end
				Else
					If Err.Number = 0 Then
					        vError = "3"
					end if
					
					sql = ""
					sql = "update db_momo.dbo.tbl_coin_log set gubun = '" & contents & "' where id = '" & vIdx & "'" + vbcrlf
					dbget.execute sql
				End IF
				
			else
				If Err.Number = 0 Then
				        vError = "2"
				end if
				'### tbl_coin_current 에 회원 아이디 없음	'//코인 신규 추가	'//코인 로그 저장
				sql = "SELECT count(userid) FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '" & userid & "'"
				rsget.open sql,dbget,1
				If rsget(0) > 0 Then
					If Err.Number = 0 Then
					        vError = "3"
					end if
					
					sql = ""
					sql = "EXECUTE [db_momo].dbo.ten_momo_coin_insert '"&userid&"',"&savecoin&" " 
					dbget.execute sql
					
					If Err.Number = 0 Then
					        vError = "4"
					end if
					sql = ""
					sql = "insert into db_momo.dbo.tbl_coin_log" + vbcrlf
					sql = sql & " (userid,coin,gubuncd,gubun,deleteyn) values" + vbcrlf
					sql = sql & " (" + vbcrlf		
					sql = sql & " '" & userid & "'" + vbcrlf
					sql = sql & " , '" & savecoin & "' " + vbcrlf
					sql = sql & " , '13' " + vbcrlf
					sql = sql & " , '" & contents & "' " + vbcrlf		
					sql = sql & " , 'N' " + vbcrlf	
					sql = sql & " )" + vbcrlf
					dbget.execute sql
				Else
					'### 회원db에 없는 아이디임.
					vError = "o"
				End If
				rsget.close					
			end if
	Else
	'####### 아이디 여러개일 경우 #######
			vTemp = Split(userid_m, ",")
			For i = 0 To ubound(vTemp)
			
					sql = ""
					tmp = 0
					userid = ""
					userid = Trim(vTemp(i))
					
					'//코인 내용 존재여부 확인	
					If Err.Number = 0 Then
					        vError = "1"
					end if
						
					sql = "select savecoin, currentcoin from db_momo.dbo.tbl_coin_current where userid = '"&userid&"' and isusing='Y'"
					
					'response.write sql &"<br>"
					rsget.open sql,dbget,1
						if not(rsget.bof or rsget.eof) then
							tmp = 1
							IF Left(savecoin,1) = "-" Then
								If (CDbl(savecoin) + (CDbl(rsget("savecoin"))-CDbl(rsget("currentcoin")))) < 0 Then
									savecoin = CDbl("-" & (CDbl(rsget("savecoin"))-CDbl(rsget("currentcoin"))))
									vAlert = vAlert & "" & userid & " 고객님은 기존 코인이 " & (CDbl(rsget("savecoin"))-CDbl(rsget("currentcoin"))) & " 여서 - 가 되므로 0 으로 처리합니다.\n"
								End IF
							End If
						end if
					rsget.close
					
					'//코인 저장	
					sql = ""
					if tmp > 0 then
						
						If vIdx = "" Then
							If Err.Number = 0 Then
							        vError = "2"
							end if
							
							sql = ""
							sql = "update db_momo.dbo.tbl_coin_current set" + vbcrlf
							IF Left(savecoin,1) = "-" Then
								sql = sql & " currentcoin = currentcoin + "& Replace(savecoin,"-","")&" , lastupdate=getdate()" + vbcrlf
							Else
								sql = sql & " savecoin = savecoin + "&savecoin&" , lastupdate=getdate()" + vbcrlf
							End If
							sql = sql & " where userid = '"&userid&"' and isusing='Y'" + vbcrlf
							
							dbget.execute sql
						End IF
						
						
						If vIdx = "" Then
							If Err.Number = 0 Then
							        vError = "3"
							end if
							
							'//코인 로그 저장
							sql = ""
							sql = "insert into db_momo.dbo.tbl_coin_log" + vbcrlf
							sql = sql & " (userid,coin,gubuncd,gubun,deleteyn) values" + vbcrlf
							sql = sql & " (" + vbcrlf		
							sql = sql & " '" & userid & "'" + vbcrlf
							sql = sql & " , '" & savecoin & "' " + vbcrlf
							sql = sql & " , '13' " + vbcrlf
							sql = sql & " , '" & contents & "' " + vbcrlf		
							sql = sql & " , 'N' " + vbcrlf	
							sql = sql & " )" + vbcrlf
									
							dbget.execute sql
					'		response.write sql &"<br>"
					'		dbget.RollBackTrans
					'		response.end
						Else
							If Err.Number = 0 Then
							        vError = "3"
							end if
							
							sql = ""
							sql = "update db_momo.dbo.tbl_coin_log set gubun = '" & contents & "' where id = '" & vIdx & "'" + vbcrlf
							dbget.execute sql
						End IF
						
					else

						If Err.Number = 0 Then
						        vError = "2"
						end if
						'### tbl_coin_current 에 회원 아이디 없음	'//코인 신규 추가	'//코인 로그 저장
						sql = "SELECT count(userid) FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '" & userid & "'"
						rsget.open sql,dbget,1
						If rsget(0) > 0 Then
							If Err.Number = 0 Then
							        vError = "3"
							end if
							
							sql = ""
							sql = "EXECUTE [db_momo].dbo.ten_momo_coin_insert '"&userid&"',"&savecoin&" " 
							dbget.execute sql
							
							If Err.Number = 0 Then
							        vError = "4"
							end if
							sql = ""
							sql = "insert into db_momo.dbo.tbl_coin_log" + vbcrlf
							sql = sql & " (userid,coin,gubuncd,gubun,deleteyn) values" + vbcrlf
							sql = sql & " (" + vbcrlf		
							sql = sql & " '" & userid & "'" + vbcrlf
							sql = sql & " , '" & savecoin & "' " + vbcrlf
							sql = sql & " , '13' " + vbcrlf
							sql = sql & " , '" & contents & "' " + vbcrlf		
							sql = sql & " , 'N' " + vbcrlf	
							sql = sql & " )" + vbcrlf
							dbget.execute sql
						Else
							'### 회원db에 없는 아이디임.
							vError = "o"
							Exit For
						End If
						rsget.close
					end if
					
				savecoin = org_savecoin
			
			Next
	End If
		

	If Err.Number = 0 Then
	        dbget.CommitTrans
	Else
	        dbget.RollBackTrans
	        response.write "<script>alert('[" & vError & "]데이타를 저장하는 도중에 에러가 발생하였습니다.')</script>"
	        response.write "<script>history.back()</script>"
	        response.end
	End If
	
on error Goto 0

	If vError = "o" Then
		Response.Write "<script>alert('"&userid&" 아이디가 존재하지 않습니다.');location.href='bonus_coin_give.asp';</script>"
		dbget.close()
		Response.End
	Else
		If vIdx = "" Then
			If vInsertGubun = "one" Then
				Response.Write "<script>alert('" & vAlert & "" & userid & "님께 " & savecoin & " 코인이 지급되었습니다.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
			Else
				Response.Write "<script>alert('" & vAlert & "\n" & savecoin & " 코인이 지급되었습니다.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
			End If
		Else
			Response.Write "<script>alert('보내스내역이 수정되었습니다.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
		End If
		dbget.close()
		Response.End
	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->