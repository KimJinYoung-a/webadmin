<%

Class QnaListItem
	public FQdate
	public FRedate  
	public FDelayCount
	public FAvgTime
	
	Private sub Class_Intialize()
	
	End sub
	
	Private Sub Class_Terminate()
	
	End Sub
	
end class

Class QnaList
	public Compdate()
	public FTotalCount
	public FReplyCount
	public yyyy1, mm1, dd1
	public yyyy2, mm2, dd2
	public delay
	public FAvgTime
	
	Private sub Class_Intialize()
		delay=0
		TempAvgTime=0
		FAvgTime=0
	End sub
	
	Private Sub Class_Terminate()
	
	End Sub
	 
	public function FormatCS(Byval v)
		dim d
		d=CStr(year(v))
		
		if (month(v) < 10) then
			d = d + "0" + CStr(month(v))
       	else
			d = d + CStr(month(v))
		end if
		
		if (day(v) < 10) then
			d = d + "0" + CStr(day(v))
       	else
			d = d + CStr(day(v))
		end if
		
		if (hour(v) < 10) then
			d = d + "0" + CStr(hour(v))
       	else
			d = d + CStr(hour(v))
		end if
		
		if (minute(v) < 10) then
			d = d + "0" + CStr(minute(v))
       	else
			d = d + CStr(minute(v))
		end if
		
		if (second(v) < 10) then
			d = d + "0" + CStr(second(v))
       	else
			d = d + CStr(second(v))
		end if
		
		FormatCS=d
	end function
	
	Public Sub GetQnaCount()
	
		dim strSql,i
		dim Questiondate,Redate,TempTime
		dim timeA,timeB,time1,time2,time3,time4,time5
		dim limitA,limitB,limit1,limit2,limit3,limit4,limit5
		dim d,d2,t1,t2,t3,t4,t5,t6
				
		limitA = yyyy1 & "-" & mm1 & "-" & dd1-1 &"  17:30:00"
		limitB = yyyy1 & "-" & mm1 & "-" & dd1 & " 17:30:00"
		limit1 = yyyy1 & "-" & mm1 & "-" & dd1 & " 09:00:00"
		limit2 = yyyy1 & "-" & mm1 & "-" & dd1 & " 11:00:00"
		limit3 = yyyy1 & "-" & mm1 & "-" & dd1 & " 12:00:00"
		limit4 = yyyy1 & "-" & mm1 & "-" & dd1 & " 19:00:00"
		limit5 = yyyy1 & "-" & mm1 & "-" & dd1-1 & " 19:00:00"
		
		'rsponse.write limitA
		'dbget.close()	:	response.End
		timeA = FormatCs(limitA)
		timeB = FormatCs(limitB)
		time1 = FormatCs(limit1)
		time2 = FormatCs(limit2)
		time3 = FormatCs(limit3)
		time4 = FormatCs(limit4)
		time5 = FormatCs(limit5)
		
			strSql="select regdate,replydate from [db_cs].[10x10].tbl_myqna"
			strSql=strSql + " where regdate between '" & Dateserial(Cstr(yyyy1),CStr(mm1),CStr((dd1-1))) & " " & Formatdatetime(limitA,4) & "'" + vbcrlf
			strSql=strSql + " and '" & Dateserial(Cstr(yyyy1),CStr(mm1),CStr(dd1)) & " " & Formatdatetime(limitB,4) & "'" + vbcrlf
			strSql=strSql + " order by regdate asc"
			'response.write strSql
			rsget.open strSql,dbget,1
				
			FTotalCount=rsget.recordcount
			
			if not rsget.eof then
				redim preserve Compdate(FTotalcount)
				i=0		
			
				do until rsget.eof
				set Compdate(i) = new QnaListItem
				
				Compdate(i).FQdate	= rsget("regdate")
				Compdate(i).FRedate = rsget("replydate")
				
				Questiondate=FormatCS(Compdate(i).FQdate)
								
				if Compdate(i).FRedate<>"" then
					Redate=FormatCs(Compdate(i).FRedate)
				else
					Redate=""
				end if
					
				'response.write timeA
				if Redate="" then
					 'delay=delay+1
				else
					if timeA<= Questiondate and Questiondate < time1 then							'' 전날 17시 30분~ 다음날 9시
						if Redate <= time5 or Redate <= time3  then											''전날 19시 이전 or 다음날 12이전 
							FReplyCount=FReplyCount+1	
							response.write "<br>ok --1 "
						else
							delay=delay+1																				
							FReplyCount=FReplyCount+1															
						response.write "<br>" & Questiondate
						response.write ".1." & Redate
						end if 
					
					elseif time1 <= Questiondate and Questiondate < time2 then             				''다음날  9시 ~11시
					 	if time1 <= Redate and Redate <= time3 then   		 										''다음날 9~12시 
							FReplyCount=FReplyCount+1
							response.write "<br>ok --2 "	
						else							
							delay=delay+1																				
							FReplyCount=FReplyCount+1															
						response.write "<br>" & Questiondate
						response.write ".2." & Redate 
						end if 				
						
					elseif time2 <= Questiondate and Questiondate < timeB then       						'' 다음날 11시~ 17시30분 
						if  time2 <= Redate and Redate <= time4 then   		 									'' 다음날 11시 ~19시  
							FReplyCount=FReplyCount+1
							response.write "<br>ok --3 "
						else
							delay=delay+1																				
							FReplyCount=FReplyCount+1															
							response.write "<br>" & Questiondate
							response.write ".3." & Redate 
						end if
					else 
						response.write "<br>" & Questiondate
						response.write ".4.." & Redate 
					end if
					TempTime = TempTime + datediff("h",Compdate(i).FQdate,Compdate(i).FRedate)
				end if
				
				rsget.movenext
				i=i+1
			loop
			rsget.close
			end if
	
	end sub
		
end class







Class CsTotalitems

	public FDay
	public FRegcnt
	public FFincnt
	public FDelaycnt
	public FAvgtime
	
	Private sub Class_Intialize()
	
	End sub
	
	Private Sub Class_Terminate()
	
	End Sub
	
end class


Class CsTotal
	public items()
	public FTotalCount
	public yyyy1
	public mm1
	public dd1
	public yyyy2
	public mm2
	public dd2
	public regtotal,fintotal,delaytotal,avgtotal
	public maxregcnt,maxfincnt,maxdelaycnt
	
	Private sub Class_Intialize()
				
	End sub
	
	Private Sub Class_Terminate()
	
	End Sub
	
	Public Sub GetCsTotal()
		
		dim strSql,i
		strSql ="select max(regcnt) as maxregcnt,max(fincnt) as maxfincnt,max(delaycnt) as maxdelaycnt from [db_log].[dbo].tbl_cs_daily_report "
		strSql = strSql & " where yyyymmdd between '" &  Dateserial(Cstr(yyyy1),CStr(mm1),CStr(dd1)) & "'" + vbcrlf
		strSql = strSql & " and '" & Dateserial(Cstr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
		'response.write strSql
		'dbget.close()	:	response.End
		
		rsget.open strSql,dbget,1
		
		maxregcnt=rsget("maxregcnt")
		maxfincnt=rsget("maxfincnt")
		maxdelaycnt=rsget("maxdelaycnt")			
		
		rsget.close
		
		strSql ="select yyyymmdd,regcnt,fincnt,delaycnt,isnull(avgtime,0) as avgtime from [db_log].[dbo].tbl_cs_daily_report "
		strSql = strSql & " where yyyymmdd between '" &  Dateserial(Cstr(yyyy1),CStr(mm1),CStr(dd1)) & "'" + vbcrlf
		strSql = strSql & " and '" & Dateserial(Cstr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
		
		rsget.open strSql,dbget,1
		
		FTotalCount=rsget.recordcount
		
		if not rsget.eof then
		i=0
		regtotal=0
		fintotal=0
		delaytotal=0
		avgtotal=0
			redim preserve items(FTotalCount)
		
			do until	rsget.eof
		
				Set items(i) = new CsTotalitems
		
				items(i).Fday=rsget("yyyymmdd")
				items(i).FRegcnt=rsget("regcnt")
				items(i).FFincnt=rsget("fincnt")
				items(i).FDelaycnt=rsget("delaycnt")
				items(i).FAvgtime=rsget("avgtime")
				
				regtotal=regtotal+Cint(items(i).FRegcnt)
				fintotal=fintotal+Cint(items(i).FFincnt)
				delaytotal=delaytotal+Cint(items(i).FDelaycnt)
				avgtotal=avgtotal+CLng(items(i).FAvgtime)
			rsget.movenext
			i=i+1
			loop
		rsget.close
		avgtotal=avgtotal/i
		
		
	end if
	
	End Sub

end class 

%>

