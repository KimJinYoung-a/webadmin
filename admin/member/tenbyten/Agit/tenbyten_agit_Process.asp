<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim mode, idx
	Dim AreaDiv, userid, username, posit_sn, part_sn, userPhone, userHP, ChkStart, ChkEnd, usePersonNo, etcComment
	Dim SQL, strMsg,department_id
	dim ipoint, smoney , empno, regType, maxCnt, rstCnt
	dim penaltyStartDate, penaltyEndDate, penaltyCause, isTakePoint, penaltyPoint
	
	mode 		= requestCheckvar(request("mode"),5)
	idx 		= requestCheckvar(request("idx"),8)

	AreaDiv		= requestCheckvar(request("AreaDiv"),1)
	empno 		= requestCheckvar(request("sEn"),32)
	userid		= requestCheckvar(request("userid"),32)
	username	= requestCheckvar(request("username"),20)
	posit_sn	= requestCheckvar(request("posit_sn"),8)
	part_sn		= requestCheckvar(request("part_sn"),8)
	userPhone	= requestCheckvar(request("userPhone"),18)
	userHP		= requestCheckvar(request("userHP"),18)
	ChkStart	= requestCheckvar(request("ChkStart"),10) & " " & requestCheckvar(request("ChkSTime"),5)
	ChkEnd		= requestCheckvar(request("ChkEnd"),10) & " " & requestCheckvar(request("ChkETime"),5)
	usePersonNo	= requestCheckvar(request("usePersonNo"),8)
	ipoint	= requestCheckvar(request("ipoint"),2)
	smoney	= requestCheckvar(request("smoney"),10)
	
	etcComment	= html2db(request("etcComment"))
	department_id = requestCheckvar(request("department_id"),10)

	regType = requestCheckvar(request("regType"),1)	'������ ��� ��忡�� �Ѿ��

	penaltyStartDate = requestCheckvar(request("psdate"),10)
	penaltyEndDate	= requestCheckvar(request("pedate"),10)
	penaltyCause	= html2db(request("penaltyCause"))
	isTakePoint		= requestCheckvar(request("isTakePoint"),1)
	penaltyPoint	= requestCheckvar(request("penaltyPoint"),6)

	maxCnt = 2	'�̿����� ���� �ִ� ���� ���Ѽ�

	'// ó�� �б� //
	Select Case mode
		Case "add"
		 
			'// ����Ʈ�� ���ݾ� Ȯ��
			if userid="admin" and empno = "00000000000000" then
				ipoint ="0"
				smoney="0"
			else	
					if ipoint ="" or smoney ="" then 
						%>
						<script type="text/javascript">
							alert("����Ʈ�� �ݾ��� �Է����ּ���");
							history.go(-1);
						</script>
						<%
					dbget.close: response.end
					end if
			end if

			'// ���� ���� Ȯ��
			sql = " select idx "
			sql = sql & " from  db_partner.dbo.tbl_TenAgit_Booking as b "
			sql = sql & "  inner join ( "
			sql = sql & "		select empno, userid" & vbcrlf
			sql = sql & "		from db_partner.dbo.tbl_user_tenbyten" & vbcrlf
			sql = sql & "		where isusing=1" & vbcrlf

			' ��翹���� ó��	' 2018.10.16 �ѿ��
			sql = sql & " 		and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
			sql = sql & "		union all" & vbcrlf
			sql = sql & "		select '00000000000000' as empno, 'admin' as userid" & vbcrlf
			sql = sql & "	) as t  on (b.empno = t.empno) or (b.userid = t.userid) "
			sql = sql & "	where  b.chkstart <='"&ChkEnd&"' and b.chkend >='"&ChkStart&"' and AreaDiv='"&AreaDiv&"'"		
			sql = sql & "		and  b.statdiv=1 and b.isusing = 'Y' " 

			'response.write sql & "<br>"
			rsget.Open sql, dbget, 1
			if not rsget.eof then
				idx = rsget("idx")
			end if
			rsget.close

			if not (idx = "" or  isNull(idx)) then 
 		%>
			<script type="text/javascript">
				alert("�̹� ��û�� �Ⱓ�Դϴ�. �ٸ� ��¥�� �������ּ���");
				history.go(-1);
			</script>
		<%
				dbget.close: response.end
			end if
					
			'// �г�Ƽ �Ⱓ Ȯ��
			sql = " select idx from db_partner.dbo.tbl_TenAgit_penalty where empno ='"&empno&"' and   enddate >='"&ChkStart&"' and startdate<='"&ChkEnd&"'"
			rsget.Open sql, dbget, 1
			if not rsget.eof then
				idx = rsget("idx")
			end if
			rsget.close

			if not (idx = "" or  isNull(idx)) then 
		%>
			<script type="text/javascript">
				alert("�г�Ƽ �Ⱓ�Դϴ�. ��ϺҰ����մϴ�.");
				history.go(-1);
			</script>
		<%
				dbget.close: response.end
			end if
					
			'// ��û ���ɼ� Ȯ��; �⿹���(�̿���) �ܿ��� (2018.04.05 ������)
			if Not(userid="admin" and empno = "00000000000000") and regType="" then
				sql = "SELECT COUNT(Idx) as cnt FROM db_partner.dbo.tbl_TenAgit_Booking "
				sql = sql & "WHERE UserId='" & userid & "' and EmpNo='" & empno & "' "
				sql = sql & " and IsUsing='Y' and StatDiv='1' and ChkStart>=getdate() "
				rsget.Open sql, dbget, 1
				if Not rsget.eof then
					rstCnt = rsget("cnt")
				end if
				rsget.close
				
				if rstCnt>=maxCnt then
			%>
				<script type="text/javascript">
					alert("���� �̿����� ���� ������ �ֽ��ϴ�.\n�����Ͻ� ������ Ȯ�����ּ���.(�ִ� <%=maxCnt%>��)");
					history.go(-1);
				</script>
			<%
				dbget.close: response.end
				end if
			end if
			

			'// ��� ó��
			strMsg = "��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.dbo.tbl_TenAgit_Booking " &_
					" (AreaDiv, empno,userid,   userPhone, userHP, ChkStart, ChkEnd, usePersonNo, etcComment,usepoint, usemoney, isIpkum, lastupdate , adminid) values " &_
					" ('" & AreaDiv & "'" &_
					" ,'" & empno & "'" &_
					" ,'" & userid & "'" &_  
					" ,'" & userPhone & "'" &_
					" ,'" & userHP & "'" &_
					" ,'" & ChkStart & "'" &_
					" ,'" & ChkEnd & "'" &_
					" ,'" & usePersonNo & "'" &_
					" ,'" & etcComment& "'"&_
					" , " & ipoint &_
					" ,'" & smoney& "'"&_
					" , 0, getdate(), '"&session("ssBctId")&"') "  
			dbget.Execute(SQL)
			
			'// ����Ʈ ���� ó��
			if userid<>"admin" then
				SQL ="update db_partner.dbo.tbl_TenAgit_Point set usePoint = usePoint +"&ipoint
				SQL = SQL & " WHERE empno = '"&empno&"' and isusing =1 and startday <='"&ChkStart&"' and endday >='"&ChkEnd&"'"
				dbget.Execute(SQL)
			end if	


'		Case "modi"
'			strMsg = "�����Ǿ����ϴ�."
'			SQL =	"Update db_partner.dbo.tbl_TenAgit_Booking Set " &_
'					"	AreaDiv = '" & AreaDiv & "' " &_
'					"	,userid = '" & userid & "' " &_
'					"	,username = '" & username & "' " &_
'					"	,posit_sn = '" & posit_sn & "' " &_
'					"	,part_sn = '" & part_sn & "' " &_
'					"	,userPhone = '" & userPhone & "' " &_
'					"	,userHP = '" & userHP & "' " &_
'					"	,ChkStart = '" & ChkStart & "' " &_
'					"	,ChkEnd = '" & ChkEnd & "' " &_
'					"	,usePersonNo = '" & usePersonNo & "' " &_
'					"	,etcComment = '" & etcComment & "' " &_
'					" ,department_id = '" & department_id & "' " &_
'					"Where idx=" & idx
'			dbget.Execute(SQL)

		Case "del"
			'// ���� ��ҽ� ����Ʈ ���� �� �Ⱓ���� ���� �г�Ƽ ����
			if empno <> "00000000000000" and userid <>"admin" then
'				// 2020�� ����Ʈ ��å�������� �г�Ƽ ����
'				if datediff("d",date(),ChkStart) <= 5 and datediff("d",date(),ChkStart)>0 then '5���� ��� > 3������ �̿�Ұ�, ��û����Ʈ ����
'					SQL = "insert into db_partner.[dbo].[tbl_TenAgit_penalty] (idx, empno, penaltykind,startdate, enddate, adminid)" 
'		 			SQL = SQL & " values("&idx&",'"&empno&"',1, convert(varchar(10),getdate(),121), convert(varchar(10),dateadd(month,3,getdate()),121),'"&userid&"')"        
'					dbget.Execute(SQL)
'					
'		 		elseif 	datediff("d",date(),ChkStart) =0 then '���� ��� > 6������ �̿�Ұ�, ��û ����Ʈ����, ȯ�ҺҰ�
'		 			SQL = "insert into db_partner.[dbo].[tbl_TenAgit_penalty] (idx, empno, penaltykind,startdate, enddate, adminid)" 
'		 			SQL = SQL & " values("&idx&",'"&empno&"',2,convert(varchar(10),getdate(),121),  convert(varchar(10),dateadd(month,6,getdate()),121),'"&userid&"')"        
'					dbget.Execute(SQL)
'
'				else	'���� ��Ҵ� ����Ʈ ȯ��
'					SQL =	"Update db_partner.dbo.tbl_TenAgit_Point Set "  
'					SQL = SQL &	"	usePoint = usePoint- "&ipoint  
'					SQL = SQL & " WHERE empno = '"&empno&"' and isusing =1 and startday <='"&ChkStart&"' and endday >='"&ChkEnd&"'"
'					dbget.Execute(SQL)
'				end if
				'���� ��� ����Ʈ ȯ��
				SQL =	"Update db_partner.dbo.tbl_TenAgit_Point Set "  
				SQL = SQL &	"	usePoint = usePoint- "&ipoint  
				SQL = SQL & " WHERE empno = '"&empno&"' and isusing =1 and startday <='"&ChkStart&"' and endday >='"&ChkEnd&"'"
				dbget.Execute(SQL)
		 	end if
			strMsg = "��û��� �Ǿ����ϴ�."	
	 		SQL =	"Update db_partner.dbo.tbl_TenAgit_Booking Set " &_
			"	isUsing = 'N' , canceldate = getdate() , lastupdate =getdate() " &_
			"Where idx=" & idx
			dbget.Execute(SQL)
		
					
		Case "pt"
			SQL = "insert into db_partner.[dbo].[tbl_TenAgit_penalty] (idx, empno, penaltykind,startdate, enddate, adminid)" 
 			SQL = SQL & " values("&idx&",'"&empno&"',4,convert(varchar(10),getdate(),121),  convert(varchar(10),dateadd(year,1,getdate()),121),'"&userid&"')"        
			dbget.Execute(SQL)	
			strMsg = "������ �г�Ƽ�� ��� �Ǿ����ϴ�."	

		Case "ptAdd"
			SQL = "insert into db_partner.[dbo].[tbl_TenAgit_penalty] (idx, empno, penaltykind, startdate, enddate, adminid, penaltyCause, penaltyPoint)" 
 			SQL = SQL & " values("&idx&",'"&empno&"',4,'" & penaltyStartDate & "','" & penaltyEndDate & "','"&userid&"','"&penaltyCause&"','"&penaltyPoint&"')"
			dbget.Execute(SQL)	
			
			if isTakePoint="1" and penaltyPoint>0 then
				SQL =	"Update db_partner.dbo.tbl_TenAgit_Point Set "  
				SQL = SQL &	"	usePoint = usePoint+ " & penaltyPoint  
				SQL = SQL & " WHERE empno = '" & empno & "' and isusing =1 and yyyy='" & Year(Date) & "'"
				dbget.Execute(SQL)
			end if
			
			strMsg = "������ �г�Ƽ�� ��� �Ǿ����ϴ�."	
			

		case "cal"	'�޷� ���ϵ�� 
			dim sHolidayname
			strMsg = "��ϵǾ����ϴ�."
			sHolidayname = requestCheckvar(request("sHnm"),20)

			SQL =" Update db_sitemaster.dbo.LunarToSolar set holiday =2, holiday_name ='"&sHolidayname&"' where solar_date ='"&requestCheckvar(request("ChkStart"),10)&"'"
			dbget.Execute(SQL)
	End Select

	response.write	"<script type='text/javascript'>" &_
					"	alert('" & strMsg & "');" &_
					"	opener.history.go(0);" &_
					"	self.close();" &_
					"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->