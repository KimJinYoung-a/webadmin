<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������
' Hieditor : 2009.11.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
	'### ���ʽ� ���� : gubuncd �� 13
	Dim sql, tmp, userid, userid_m, contents, savecoin, org_savecoin, vError, vIdx, vInsertGubun, vTemp, i, vAlert
	vIdx = Request("idx")
	userid = Request("giveuserid")
	contents = requestCheckVar(Request("givecontent"),100)
	savecoin = Request("givecoin")
	org_savecoin = savecoin	'### - ������ ���� ���κ��� Ŭ ��� ����ϱ� ���� ���� ������ �ϳ� ����.
	vInsertGubun = Request("insertgubun")
	userid_m = Request("giveuserid_m")
	vError = "x"

	
	'tozzinet
	On Error Resume Next
	dbget.beginTrans

	
	'####### ���̵� �Ѱ��� ��� #######
	If vInsertGubun = "one" Then
			sql = ""
			'//���� ���� ���翩�� Ȯ��	
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
							vAlert = "" & userid & " ������ ���� ������ " & (CDbl(rsget("savecoin")) - CDbl(rsget("currentcoin"))) & " ���� - �� �ǹǷ� 0 ���� ó���մϴ�.\n\n"
						End IF
					End If
				end if
			rsget.close

			'//���� ����	
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
					
					'//���� �α� ����
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
				'### tbl_coin_current �� ȸ�� ���̵� ����	'//���� �ű� �߰�	'//���� �α� ����
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
					'### ȸ��db�� ���� ���̵���.
					vError = "o"
				End If
				rsget.close					
			end if
	Else
	'####### ���̵� �������� ��� #######
			vTemp = Split(userid_m, ",")
			For i = 0 To ubound(vTemp)
			
					sql = ""
					tmp = 0
					userid = ""
					userid = Trim(vTemp(i))
					
					'//���� ���� ���翩�� Ȯ��	
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
									vAlert = vAlert & "" & userid & " ������ ���� ������ " & (CDbl(rsget("savecoin"))-CDbl(rsget("currentcoin"))) & " ���� - �� �ǹǷ� 0 ���� ó���մϴ�.\n"
								End IF
							End If
						end if
					rsget.close
					
					'//���� ����	
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
							
							'//���� �α� ����
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
						'### tbl_coin_current �� ȸ�� ���̵� ����	'//���� �ű� �߰�	'//���� �α� ����
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
							'### ȸ��db�� ���� ���̵���.
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
	        response.write "<script>alert('[" & vError & "]����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.')</script>"
	        response.write "<script>history.back()</script>"
	        response.end
	End If
	
on error Goto 0

	If vError = "o" Then
		Response.Write "<script>alert('"&userid&" ���̵� �������� �ʽ��ϴ�.');location.href='bonus_coin_give.asp';</script>"
		dbget.close()
		Response.End
	Else
		If vIdx = "" Then
			If vInsertGubun = "one" Then
				Response.Write "<script>alert('" & vAlert & "" & userid & "�Բ� " & savecoin & " ������ ���޵Ǿ����ϴ�.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
			Else
				Response.Write "<script>alert('" & vAlert & "\n" & savecoin & " ������ ���޵Ǿ����ϴ�.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
			End If
		Else
			Response.Write "<script>alert('������������ �����Ǿ����ϴ�.');parent.location.href='bonus_coin_list.asp?menupos=1156';</script>"
		End If
		dbget.close()
		Response.End
	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->