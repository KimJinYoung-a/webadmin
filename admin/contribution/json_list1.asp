<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"--> 
<!-- #include virtual="/lib/classes/contribution/contributionCls1.asp"--> 
<%
dim sdatetype
dim dstdate
dim deddate
dim sdispcate
dim itotCnt

dim clsCMeachulLog
dim licomm10,licommP,pgcomm,cardcomm,cpscomm,ptcomm
dim totcomm10, totcommP, totcommPer10, totcommPerP
dim pgcommper, cardcommper ,cpscommper, licomm10per,licommPper,ptcommper
Dim   sdt, edt ,dategbn 
	dstdate   = requestcheckvar(request("sdt"),10)
	deddate     = requestcheckvar(request("edt"),10) 
    dategbn     = requestCheckvar(request("dategbn"),32)

 
 if dategbn="" then dategbn="ActDate"
 
 sdatetype = dategbn 
 sdispcate = "N"
 

dim i, tmpJSON, j
dim gubunNM
dim buyPer, BCPer, MPPer, MPerPer

Set tmpJSON = New aspJSON 
With tmpJSON.data 
	set clsCMeachulLog = new CMeachulLog
	clsCMeachulLog.FdateType =sdatetype
	clsCMeachulLog.FstDate =dstdate
	clsCMeachulLog.FedDate =deddate
	clsCMeachulLog.FDispCate =sdispcate
	clsCMeachulLog.fnGetOrerLogData
	iTotCnt = clsCMeachulLog.FTotCnt

	if iTotCnt > 0 then	 
		 '------------------구매총액 ----------------------------------------------------
		j = 0
		For i = 0 to iTotCnt 
		
			buyPer= 0 
			if clsCMeachulLog.Fbuytot_sum > 0 then 
			buyPer = round((clsCMeachulLog.Fbuytot(i)/clsCMeachulLog.Fbuytot_sum)*100)
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "1.구매총액"
			.Add "제휴구분", clsCMeachulLog.Fsitename(i)  
			.Add "구분1",   getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))
			.Add "날짜", clsCMeachulLog.FDDate(i)
			.Add "전체_금액", clsCMeachulLog.Fbuytot(i)  
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   

			j = j+ 1
		Next
		 '------------------/구매총액 ----------------------------------------------------
		'------------------보너스쿠폰 ----------------------------------------------------
		For i = 0 to iTotCnt 
		
			BCPer= 0 
			if clsCMeachulLog.FBCtot_sum > 0 then 
			BCPer = round((clsCMeachulLog.FBCtot(i)/clsCMeachulLog.FBCtot_sum)*100)
			end if

			.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "2.보너스쿠폰"
			.Add "제휴구분", clsCMeachulLog.Fsitename(i)  
			.Add "구분1",   getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))
			.Add "날짜", clsCMeachulLog.FDDate(i)
			.Add "전체_금액", clsCMeachulLog.FBCtot(i)
			.Add "전체_율", BCPer/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
			 j= j+1
		Next
	 	'------------------/ 보너스쿠폰  ----------------------------------------------------
	 	'------------------취급액 ----------------------------------------------------
		For i = 0 to iTotCnt 
		
			MPPer= 0 
			if clsCMeachulLog.FMPricetot_sum > 0 then 
			MPPer = round((clsCMeachulLog.FMPricetot(i)/clsCMeachulLog.FMPricetot_sum)*100)
			end if

			.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "3.취급액"
			.Add "제휴구분", clsCMeachulLog.Fsitename(i)  
			.Add "구분1",   getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))
			.Add "날짜", clsCMeachulLog.FDDate(i)
			.Add "전체_금액", clsCMeachulLog.FMPricetot(i)
			.Add "전체_율", MPPer/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
			 j= j+1
		Next
		 '------------------/ 취급액  ----------------------------------------------------
	 	 '------------------취급액 수익----------------------------------------------------
		For i = 0 to iTotCnt 
		
			MPerPer= 0 
			if clsCMeachulLog.FBCtot_sum > 0 then 
			MPerPer = round((clsCMeachulLog.FMPertot(i)/clsCMeachulLog.FMPertot_sum)*100)
			end if

			.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "4.취급액 수익"
			.Add "제휴구분", clsCMeachulLog.Fsitename(i)  
			.Add "구분1",   getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))
			.Add "날짜", clsCMeachulLog.FDDate(i)
			.Add "전체_금액", clsCMeachulLog.FMPertot(i) 
			.Add "전체_율", MPerPer/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1
		Next
		
	 '------------------/ 취급액 수익 ----------------------------------------------------
 	else
			.Add 0, tmpJSON.Collection()
				With .item(0) 
				.Add "매출구분", "x"
				.Add "제휴구분",   "x"
				.Add "구분1",   "x"
				.Add "날짜",""
				.Add "전체_금액", 0
				.Add "전체_율", 0
				if sdispcate ="Y" then
				.Add "디자인문구",0	 
				end if
				End With  

	end if	
	
	'------ 변동비1 ----------	
	 clsCMeachulLog.fnGetLinceComm
	 licomm10 = round(clsCMeachulLog.FliComm10)
	 licommP = round(clsCMeachulLog.FliCommP)
	 clsCMeachulLog.fnGetComm
	 pgcomm  = 	round(clsCMeachulLog.FpgComm)
	 cardcomm  = 	round(clsCMeachulLog.FcardComm)
	 cpscomm  = 	round(clsCMeachulLog.FcpsComm )	
 	  totcomm10 = licomm10 +pgcomm+ cardcomm+cpscomm 
	  totcommP = licommP + ptcomm


	  totcommPer10 = 0
	  pgcommper = 0
	  cardcommper = 0
	  cpscommper = 0
	  licomm10per = 0
	if totcomm10 > 0 then 
		pgcommper = round((pgcomm/totcomm10)*100)
		cardcommper = round((cardcomm/totcomm10)*100)
		cpscommper = round((cpscomm/totcomm10)*100)
		licomm10per = round((licomm10/totcomm10)*100)
		  totcommPer10 =  pgcommper+ cardcommper +cpscommper+ licomm10per
	end if
	j = j + 1
	 	.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "판매수수료" 
			.Add "날짜", ""
			.Add "전체_금액", totcomm10
			.Add "전체_율", totcommPer10/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","PG수수료"
			.Add "날짜", ""
			.Add "전체_금액", pgcomm
			.Add "전체_율", pgcommper/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		 .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","신용카드수수료"
			.Add "날짜", ""
			.Add "전체_금액", cardcomm
			.Add "전체_율",  cardcommper/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1
	 
		.Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","CPS수수료"
			.Add "날짜", ""
			.Add "전체_금액", cpscomm
			.Add "전체_율", cpscommper/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","라이센스수수료"
			.Add "날짜", ""
			.Add "전체_금액", licomm10
			.Add "전체_율",  licomm10per/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		 .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2",""
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		   .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2","상품배송비"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2","판매포장비"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직급여(물류)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직급여(연차수당)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "10X10" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직퇴직급여(물류)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

'변동비1 - 제휴
		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "판매수수료" 
			.Add "날짜", ""
			.Add "전체_금액", totcommP
			.Add "전체_율", totcommPerP/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","제휴수수료"
			.Add "날짜", ""
			.Add "전체_금액", ptcomm
			.Add "전체_율", ptcommper/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

	 

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "판매수수료"
			.Add "구분2","라이센스수수료"
			.Add "날짜", ""
			.Add "전체_금액", licommP
			.Add "전체_율",  licommPper/100
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		 .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2",""
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		   .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2","상품배송비"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2","판매포장비"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직급여(물류)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직급여(연차수당)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "5.변동비1"
			.Add "제휴구분", "제휴" 
			.Add "구분1",  "물류비"
			.Add "구분2","계약직퇴직급여(물류)"
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		  '공헌이익
		   .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "6.공헌이익1(전체)"
			.Add "제휴구분", "10x10"  
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1

		    '공헌이익
		   .Add j, tmpJSON.Collection()
			With .item(j) 
			.Add "매출구분", "6.공헌이익1(전체)"
			.Add "제휴구분", "제휴"  
			.Add "날짜", ""
			.Add "전체_금액", "" 
			.Add "전체_율",  ""
			if sdispcate ="Y" then
			.Add "디자인문구", "1500"	 
			end if
			End With   
		  j= j+1
	set clsCMeachulLog = nothing	  
End With
	Response.Write tmpJSON.JSONoutput() 
	
	Set tmpJSON = Nothing
 
 %> 