<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<%
	public function IsSailItem()
		if NowEventDoing then
			IsSailItem = (((rsget("eventprice")<>0) and (rsget("sellcash")>rsget("eventprice"))) or ((rsget("sailYn")="Y") and (FOrgPrice>rsget("sellcash")))) or ((rsget("SpecialUserItem")>0) and (getUserLevel()>0))
		else
			IsSailItem = ((rsget("sailYn")="Y") and (FOrgPrice>rsget("sellcash"))) or ((rsget("SpecialUserItem")>0) and (getUserLevel()>0))
		end if
	end function

	public function getUserLevel()
		dim uselevel
		uselevel = request.cookies("uinfo")("userlevel")
		if uselevel="" then
			getUserLevel = "0"
		else
			getUserLevel = uselevel
		end if
	end function

	public function IsFreeBeasongCoupon()
		if Fitemcoupontype ="3" or Fdeliverytype >= 2 then
			IsFreeBeasongCoupon	=	true
		else
			IsFreeBeasongCoupon	=	false
	end if
	end function


	public function getOrgPrice()
		if FOrgPrice=0 then
			getOrgPrice = rsget("sellcash")
		else
			getOrgPrice = FOrgPrice
		end if
	end function

	public function getRealPrice()
		if NowEventDoing then
			if rsget("eventprice")=0 then
				getRealPrice = rsget("sellcash")
			else
				getRealPrice = rsget("eventprice")
			end if

		else
			getRealPrice = rsget("sellcash")
		end if

	end function

	public function IsSpecialUserItem()
		IsSpecialUserItem = (rsget("SpecialUserItem")>0)
	end function

	public Function FormatStr(value)
		if value<10 then
			FormatStr="0" & CStr(value)
		else
			FormatStr = CStr(value)
		end if
	end function


Dim SQL,insertSQL
dim Strgubun,rank
dim firstdate,lastdate

lastdate=FormatStr(year(now())) & "-" & FormatStr(month(now())) & "-" & FormatStr(day(now()))
firstdate=FormatStr(year(now())) & "-" & FormatStr(month(now())) & "-" & FormatStr(day(now())-7)
'lastdate="2006-05-13"
'firstdate="2006-05-19"

dim fso,tFile,filecontents,savePath,filename
dim strbestHtml

'======= left best10 생성 ========
savePath = server.mappath("/Diary_collection_2007/") + "\"
filename="diary_left_bestItem10.asp"

Set fso= Server.CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(savePath & FileName )


'SQL =	" SELECT TOP 10 "&_
'			" i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem "&_
'			" ,dm.idx ,dm.icon_img "&_
'			" ,sum(d.itemno) as total ,d.itemcost "&_
'			" FROM  [db_order].[10x10].tbl_order_master m "&_
'			" JOIN [db_order].[10x10].tbl_order_detail d "&_
'			" 	ON m.orderserial=d.orderserial "&_
'			" 	and m.ipkumdiv>3 and m.cancelyn='N' and m.jumundiv<>9 and d.itemid<>0 and d.cancelyn<>'Y'  "&_
'			" JOIN db_item.[10x10].tbl_item i "&_
'			" 	ON i.itemid=d.itemid and i.itemid<>0 "&_
'			" JOIN db_contents.dbo.tbl_diary_master dm	 "&_
'			" 	ON i.itemid=dm.itemid "&_
'			" WHERE m.regdate BETWEEN '" & firstdate & "' AND '" & lastdate & "' /* 월요일 ~일요일 */"&_
'			" GROUP BY i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem  "&_
'			" 		,dm.idx,dm.icon_img,d.itemcost "&_
'			" ORDER BY TOTAL DESC "

		SQL =	" SELECT top 10  "&_
					" i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem  "&_
					" ,dm.idx ,dm.icon_img "&_
					" FROM [db_item].[10x10].tbl_item i "&_
					" JOIN [db_contents].[dbo].tbl_diary_master dm "&_
					" 	ON i.itemid=dm.itemid and dm.itemid<>0 "&_
					" ORDER BY i.recentsellcount desc "


'response.write sql
			rsget.open SQL,dbget,1

			strGubun=FormatStr(year(now())) & FormatStr(month(now())) & FormatStr(day(now())) &_
							 FormatStr(Hour(now())) & FormatStr(Minute(now())) & FormatStr(Second(now()))

			rank=0
			insertSQL=""

			if not rsget.eof then

					'//db 입력
					do until rsget.eof


							insertSQL =	insertSQL & " insert into [db_contents].[dbo].tbl_diary_bestitem10 (gubun,rank,diaryidx,itemid,itemname,sellcash,icon_img) " &_
																			" values( " &_
																			"'" & strGubun & "' " &_
																			",'" & rank & "' " &_
																			",'" & rsget("idx") & "' " &_
																			",'" & rsget("itemid") & "' " &_
																			",'" & rsget("itemname") & "' " &_
																			",'" & getRealPrice & "' " &_
																			",'" & rsget("icon_img") & "' " &_
																			")"
							rsget.movenext
							rank=rank+1

					loop

					dbget.Execute(insertSQL)

				'// html 만들기

					rsget.movefirst
					rank=0

					strbestHtml	=	" <table width='72' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#FFFFFF'> " &_
				  							" <tr> "&_
				    						" 	<td valign='top'><img src='http://fiximage.10x10.co.kr/images/diary_collection_2007/best10_title.gif' width='182' height='38'></td> " &_
				  							" </tr> "

					do until rsget.eof

					strbestHtml	=	strbestHtml & "" &_
												"		<tr>" &_
					    					" 		<td  style='border-bottom: 1px solid #ffffff'> " & _
					      				"				<table width='182' border='0' cellspacing='2'> " & _
					        			"					<tr> " & _
												"						<td width='12'><img src='http://fiximage.10x10.co.kr/images/diary_collection_2007/best_" & FormatStr(rank+1) & ".gif' width='12' height='12' alt='' /></td> " & _
												"						<td width='52'> " & _
												"							<table width='50' border='0' cellpadding='0' cellspacing='1' bgcolor='cccccc'> " & _
												"								<tr> " & _
												"									<td><a href='/diary_collection_2007/diary_prd.asp?idx=" & rsget("idx") & "'><img src='http://imgstatic.10x10.co.kr/contents/diary/icon/" & rsget("icon_img") & "' width='50' height='50' border='0' alt='' /></a></td> " & _
												"								</tr> " & _
												"							</table> " & _
												"						</td> " & _
												"						<td> " & _
												"							<table border='0' width='100%' cellpadding='0' cellspacing='1'> " & _
												"								<tr> " & _
												"									<td><a href='/diary_collection_2007/diary_prd.asp?idx=" & rsget("idx") & "'><font color='#666666'>" & db2html(rsget("itemname")) & "</font></a></td> " & _
												"								</tr> " & _
												"								<tr> " & _
												"									<td class='prd_price'>" & FormatNumber(getRealPrice,0) & "원</td> " & _
												"								</tr> " & _
												"							</table> " & _
												"						</td> " & _
												"					</tr> " & _
												"				</table> " & _
												"			</td> " & _
												"		</tr> "


				  rsget.movenext
					rank=rank+1
					loop

					strbestHtml	=	strbestHtml & "</table>"

					tFile.Write strbestHtml

					tFile.Close
					Set tFile = Nothing
					Set fso = Nothing

			end if

			rsget.close


'======= 메인 best 15 생성 ========

savePath = server.mappath("/Diary_collection_2007/") + "\"
filename="diary_main_bestItem15.asp"

Set fso= Server.CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(savePath & FileName )



'		SQL =	" SELECT TOP 15 "&_
'					" i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem "&_
'					" ,dm.idx ,dm.list_img "&_
'					" ,sum(d.itemno) as total ,d.itemcost "&_
'					" FROM  [db_order].[10x10].tbl_order_master m "&_
'					" JOIN [db_order].[10x10].tbl_order_detail d "&_
'					" 	ON m.orderserial=d.orderserial "&_
'					" 	and m.ipkumdiv>3 and m.cancelyn='N' and m.jumundiv<>9 and d.itemid<>0 and d.cancelyn<>'Y'  "&_
'					" JOIN db_item.[10x10].tbl_item i "&_
'					" 	ON i.itemid=d.itemid and i.itemid<>0 "&_
'					" JOIN db_contents.dbo.tbl_diary_master dm	 "&_
'					" 	ON i.itemid=dm.itemid "&_
'					" WHERE m.regdate BETWEEN '" & firstdate & "' AND '" & lastdate & "' /* 월요일 ~일요일 */"&_
'					" GROUP BY i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem  "&_
'					" 		,dm.idx,dm.list_img,d.itemcost "&_
'					" ORDER BY TOTAL DESC "

		SQL =	" SELECT top 15  "&_
					" i.itemid,i.itemname ,i.sellcash ,i.sailYn ,i.eventprice ,i.orgprice ,i.specialuseritem  "&_
					" ,dm.idx ,dm.list_img "&_
					" FROM [db_item].[10x10].tbl_item i "&_
					" JOIN [db_contents].[dbo].tbl_diary_master dm "&_
					" 	ON i.itemid=dm.itemid and dm.itemid<>0 "&_
					" ORDER BY i.recentsellcount desc "


					rsget.open SQL,dbget,1


					if not rsget.eof then
					i=1
					strbestHtml	=	"" &_
				  							"<table width='630' border='0' cellpadding='0' cellspacing='0'> "&_
				  							"	<tr> "

				  do until rsget.eof

					  strbestHtml	=	strbestHtml &	"		<td width='126' valign='top'> "&_
																  			"			<div align='center'> "&_
																  			"				<table width='38' border='0' cellpadding='0' cellspacing='2'> "&_
																  			"					<tr> "&_
																  			"						<td> "&_
																  			"							<div align='left'><img src='http://fiximage.10x10.co.kr/images/diary_collection_2007/best15_" & CStr(i) & ".gif' width='20' height='20'></div> "&_
																  			"						</td> "&_
																  			"					</tr> "&_
																  			"					<tr> "&_
																  			"						<td> "&_
																  			"							<table width='100' border='0' cellpadding='3' cellspacing='1' bgcolor='#CCCCCC'> "&_
																  			"								<tr> "&_
																  			"									<td bgcolor='#FFFFFF'><a href='/diary_collection_2007/diary_prd.asp?idx=" & rsget("idx") & "'><img src='http://imgstatic.10x10.co.kr/contents/diary/list/" & rsget("list_img") & "' width='100' height='100' border='0'></a></td> "&_
																  			"								</tr> "&_
																  			"							</table> "&_
																  			"						</td> "&_
																  			"					</tr> "&_
																  			"					<tr> "&_
																  			"						<td style='padding-top:5'> "&_
																  			"							<div align='center'><a href='/diary_collection_2007/diary_prd.asp?idx=" & rsget("idx") & "' class='link_kor'>" & db2html(rsget("itemname")) & "</a></div> "&_
																  			"						</td> "&_
																  			"					</tr> "&_
																  			"					<tr> "&_
																  			"						<td  height='12'> "&_
																  			"							<div align='center' class='prd_price'>" & FormatNumber(getRealPrice,0) & "원 </div> "&_
																  			"						</td> "&_
																  			"					</tr> "&_
																  			"				</table> "&_
																  			"			</div> "&_
																  			"		</td> "
					if i mod 5 = 0 then
						strbestHtml	=	strbestHtml &	" </tr>" &_
																				"</table> " &_
																				"<table width='630' border='0' cellpadding='0' cellspacing='0'> "&_
				  															"	<tr> "
				  end if

					rsget.movenext
					i=i+1
					loop
				  strbestHtml	=	strbestHtml &	"	</tr> "&_
				  														"</table> "

					tFile.Write strbestHtml

					tFile.Close
					Set tFile = Nothing
					Set fso = Nothing

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->