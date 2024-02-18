<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'########################################################### 
' Description :  이벤트 개요 데이터처리 - 등록, 수정, 삭제
' History : 2017.11.09 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
 <%
 dim sMode,menupos
 dim strSql
 dim evtNm, evtSD, evtED, evtSale, evtCoupon , evtGift, evtdisp1, evtdisp2, evtTag
 dim makerid ,evtKind, evtManager, evtState
 dim evtCode
 dim evtGCode, evtGDesc, evtGSort
 dim evtSalePer, evtsalecper
 dim retStep
 dim arritem
 dim etcitemimg,evtmolistbanner,evtNmW,evtNmM,subcopyK,evtsubname,mdtheme,mdthememo,thmcolor,thmcolormo,txtbgcolor,txtbgcolormo
 dim gUsing,gtext1,gimg1,gtext2, gimg2,gtext3, gimg3
 dim slideimg1, slideimg2,slideimg3,slideimgm1, slideimgm2,slideimgm3
 dim adminid
 dim etext  ,intLoop
 dim vChangeContents
 dim dispCate
 
 sMode = requestCheckVar(Request("hidM"),2)
 menupos= requestCheckVar(Request("menupos"),10)
 evtCode = requestCheckVar(Request("eC"),10)
 evtNm = requestCheckVar(Request("evtNm"),64)
 evtSD = requestCheckVar(Request("evtSD"),10)
 evtED = requestCheckVar(Request("evtED"),10)
 
 evtSale = requestCheckVar(Request("evtSale"),1)
 evtCoupon = requestCheckVar(Request("evtCoupon"),1)
 evtGift = requestCheckVar(Request("evtGift"),1)
 
 dispCate = requestCheckVar(Request("disp"),64)
 evtTag = requestCheckVar(Request("evtTag"),300)
      
 evtKind = 1 '쇼핑찬스
 evtManager =1 '텐텐
 evtState =0 '업체 등록중 
 adminid = session("ssBctID")
 
 if evtSale ="" then evtSale = 0
 if evtCoupon ="" then evtCoupon = 0
 if evtGift ="" then evtGift = 0
 
 
 evtGCode =   requestCheckVar(Request("eGC"),10)	
 
 evtSalePer = 	requestCheckVar(Request("eSP"),10)
 evtsalecper = 	requestCheckVar(Request("eCP"),10)
retStep= 	requestCheckVar(Request("retStep"),1)

Dim vCnt : vCnt = Request.Form("cksel").count
	Dim sdiv : sdiv = requestCheckVar(Request.Form("sdiv"),1)
	Dim upback : upback = requestCheckVar(Request.Form("upback"),1)
	dim k 
	dim page, itemid, itemname, sellyn,sailyn
 page= requestCheckVar( Request.Form("page") ,10) 
itemid  	= requestCheckvar(request("itemid"),255)
itemname 	= RequestCheckVar(request("itemname"),32)
sellyn		= RequestCheckVar(request("sellyn"),10)
sailyn 		= RequestCheckVar(request("sailyn"),10)
dispCate 	= requestCheckvar(request("disp"),16)
makerid   = requestCheckvar(request("makerid"),32)
select Case sMode
 
Case "GA" 'step2 그룹추가
 
 	  evtGDesc =  requestCheckVar(Request("hidGNm"),64)
 	  evtGSort =  requestCheckVar(Request("hidGS"),10)
 	   ' 최상위 그룹추가 
 	   dim evtPGcode ,evtGdepth
 	   evtPGcode =0
 	   strSql ="select isNull(evtgroup_code,0)  from db_event.dbo.tbl_partner_eventitem_group where evt_code = "&evtCode&" and evtgroup_pcode =0 "
 	   	rsget.Open strSql,dbget,1
 	   	if not rsget.eof then
 	   		evtPGcode = rsget(0)
	 	  end if
	 	  rsget.close
	 	   
	 	  
	 	  if evtPGcode = 0 then
		 	   strSql = "INSERT INTO db_event.dbo.tbl_partner_eventitem_group (evt_code,evtgroup_desc,evtgroup_sort,evtgroup_depth,evtgroup_pcode)"
		 	   strSql = strSql & " VALUES("&evtCode&",'최상위',0,100,0)"   	 
		 	   dbget.Execute strSql 
		 	   
		 	   strSql = "select SCOPE_IDENTITY()" 
				rsget.Open strSql,dbget,0
					evtPGcode = rsget(0)
				rsget.close
 	   end if
 	    
 	   	strSql = "select isnull(max(evtgroup_depth),0)+1 FROM  [db_event].[dbo].[tbl_partner_eventitem_group] WHERE evt_code = "&evtCode&" and (evtgroup_code = "& evtPGcode&" OR evtgroup_pcode ="&evtPGcode&")  "
 	   	 rsget.Open strSql, dbget,1
			IF not (rsget.EOF or rsget.BOF) THEN
				evtGdepth = 	rsget(0)
			END IF	
			rsget.Close
			 
 	    strSql = "INSERT INTO db_event.dbo.tbl_partner_eventitem_group (evt_code,evtgroup_desc,evtgroup_sort,evtgroup_depth,evtgroup_pcode)"
		 	strSql = strSql & " VALUES("&evtCode&",'"&evtGDesc&"',"&evtGSort&","&evtGdepth&","&evtPGcode&")"   
		 	dbget.Execute strSql  
 	   
 	  %>
 	  <script type="text/javascript">
 		parent.jsSetGList('','A');  	
 		location.href = "about:blank";
 	</script>
 		<%
 	    response.end
Case "GM"    '그룹수정
 	dim evtModGDesc
 	 
 	evtModGDesc = requestCheckVar(Request("hidGNm"),64)
 	strSql = "update db_event.dbo.tbl_partner_eventitem_group set evtgroup_desc ='"&evtModGDesc&"' where evt_Code = "&evtCode&" and evtgroup_code ="&evtGCode 	 
 	  dbget.Execute strSql 
 	 
 	%>
 	<script type="text/javascript">
 		parent.jsSetGList('','M');  		
 		location.href = "about:blank";
 	</script>
 	<%
 	response.end
Case "GD"    '그룹삭제
 	  strSql =" update db_event.dbo.tbl_partner_eventitem_group set evtgroup_using ='N' where evt_Code = "&evtCode&" and evtgroup_code ="&evtGCode
 	  dbget.Execute strSql 
 	  	  %>
 	  <script type="text/javascript">
 	parent.jsSetGList('','D'); 
 	location.href = "about:blank";
 	</script>
 		<%
 	     response.end
Case "GS" '그룹순서변경
 	dim arrEGC,  i, arrEGS
 	arrEGC = 	requestCheckVar(Request("arrGC"),200)
 	arrEGS = 	requestCheckVar(Request("arrGS"),200)
 	 evtGCode = split(arrEGC,",")
 	 evtGSort = split(arrEGS,",")
 	 
 	 for i =0 To UBound(evtGCode)
 		strSql = "update db_event.dbo.tbl_partner_eventitem_group set  evtgroup_sort="&trim(evtGSort(i))&" where evtgroup_code="&trim(evtGCode(i))      	
 		  dbget.Execute strSql 
 	next
 	%>
 	   <script type="text/javascript">
 	parent.jsSetGList('',''); 
 	parent.jsViewGS();
 	location.href = "about:blank";
 	</script>
 		<%
 	     response.end   
 	    
 CASE "IA" '그룹상품등록
 	
	if Request.Form("cksel") <> "" then
		if checkNotValidHTML(Request.Form("cksel")) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
 
 	dim gitemCnt
 	gitemCnt=0
 	strSql = " select count(itemid) from [db_event].[dbo].[tbl_partner_eventitem] where evt_code = "&evtCode&" and evtitem_isusing =1 and evtgroup_code =  "&evtGcode
 	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		gitemCnt = rsget(0)
	end if
	rsget.close
	 
	if (gitemCnt+ vCnt) >105 then 
 	%>
  	<script>
  	alert("상품등록 최대 105개까지 가능합니다.");  	
  	</script>
  	<%
	response.end
	end if
'배열로 처리
redim arritemcode(vCnt)
dim errid,arrerrid
arrerrid = ""

for i=1 to vCnt
	errid =""
	arritemcode(i) = requestCheckVar(Request.Form("cksel")(i),10)
	strSql ="SELECT itemid "
	strSql = strSql &"				FROM [db_event].[dbo].[tbl_partner_eventitem] as i"
	strSql = strSql &"				inner join db_event.dbo.tbl_partner_eventitem_group as g on i.evt_code = g.evt_code and i.evtgroup_code = g.evtgroup_code and g.evtgroup_using = 'Y'" 
	strSql = strSql &"	 where  i.itemid='" & arritemcode(i) & "' and i.evt_code="&evtCode&" and i.evtgroup_code <> '" & evtGcode & "' and i.evtitem_isusing =1 "		 
	rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		errid = rsget(0)
 	end if
 	rsget.close
 	if arrerrid =""  then 
 		arrerrid = errid
 	else
 		arrerrid = arrerrid &"," &errid
	end if	
  
		strSql = " IF Not Exists(SELECT evt_code FROM [db_event].[dbo].[tbl_partner_eventitem] WHERE itemid='" & arritemcode(i) & "' and evt_code="&evtCode&")"			
		strSql = strSql & "	BEGIN "
		strSql = strSql & " 			INSERT INTO [db_event].[dbo].[tbl_partner_eventitem] (evt_code,  itemid, evtgroup_code, evtitem_sort)"
		strSql = strSql & "     	VALUES (" & evtCode & ", " & arritemcode(i) &",'" & evtGcode & "', " & i & ")"
		strSql = strSql &" 	END "
		strSql = strSql & " ELSE "
		strSql = strSql & " 	BEGIN "			
		strSql = strSql & "			UPDATE i "
		strSql = strSql & " 		SET evtgroup_code ='"&evtGcode&"', evtitem_sort ='" & i & "', evtitem_isusing =1"
		strSql = strSql &"				FROM [db_event].[dbo].[tbl_partner_eventitem] as i"
		strSql = strSql &"				inner join db_event.dbo.tbl_partner_eventitem_group as g on i.evt_code = g.evt_code and i.evtgroup_code = g.evtgroup_code  " 
		strSql = strSql & " 		WHERE i.evt_code = '" & evtCode & "' "
		strSql = strSql & " 		and i.itemid ="&arritemcode(i)&" and( (g.evtgroup_using = 'Y' and  i.evtitem_isusing =0) or(g.evtgroup_using ='N')) " 
		strSql = strSql & " 	END "  
		dbget.execute strSql
next	   
  if arrerrid <> "" THEN
  	%>
  	<script>
  	alert("상품코드: <%=arrerrid%>\n 이미 등록된 상품입니다. 그룹 변경을 원하시면 기존 그룹코드에서 삭제 후 새로 등록해주세요");  	
  	</script>
  	<%
end if
%>
<script>
	parent.location.href='/admin/eventmanage/wait/popRegDispItem.asp?ec=<%=evtCode%>&eGc=<%=evtGcode%>&page=<%=page%>&itemname=<%=itemname%>&itemid=<%=itemid%>&sellyn=<%=sellyn%>&sailyn=<%=sailyn%>&disp=<%=dispcate%>'; 
</script>
<%
 	response.end 	
Case "IU" '그룹상품 상품삭제 
 arritem = requestCheckVar( Request.Form("delid") ,100)
	strSql = " update  [db_event].[dbo].[tbl_partner_eventitem] set evtitem_isusing = 0 WHERE evt_code=" & evtCode & " and evtgroup_code='" & evtGcode & "' and itemid in (" &arritem& ")" 
	dbget.execute strSql

%>
	<script>	
		alert('삭제 되었습니다.');		
		parent.location.href='/admin/eventmanage/wait/popRegDispItem.asp?ec=<%=evtCode%>&eGc=<%=evtGcode%>&page=<%=page%>&itemname=<%=itemname%>&itemid=<%=itemid%>&sellyn=<%=sellyn%>&sailyn=<%=sailyn%>&disp=<%=dispcate%>'; 
		</script>
	<%
	response.End   	
Case "IS" '그룹상품 상품정렬순서 변경
dim eImgSize
eImgSize= requestCheckVar( Request.Form("eImgSize") ,10)
if eImgSize ="" then eImgSize = 240
redim arritemcode(vCnt)
redim arritemsort(vCnt)
 For i=1 to vCnt
	errid =""
	arritemcode(i) = requestCheckVar(Request.Form("cksel")(i),10)
	arritemsort(i) = requestCheckVar(Request.Form("iSort")(i),10)
	
	strSql ="update [db_event].[dbo].[tbl_partner_eventitem] set evtitem_sort = "&arritemsort(i)&" , evtitem_imgsize  = "&eImgSize
	strSql= strSql & "	where evt_code ="&evtCode&" and itemid ="&arritemcode(i) &" and evtgroup_code="&evtGCode 
	dbget.execute strSql
Next 
%>
	<script>	
		alert('저장되었습니다.');				
		$(opener.document).find("#btnItem<%=evtGCode%>").val("상품(<%=vCnt%>)");
		self.close();
		//parent.location.href='/partner/event/plan/popRegDispItem.asp?ec=<%=evtCode%>&eGc=<%=evtGcode%>&page=<%=page%>&itemname=<%=itemname%>&itemid=<%=itemid%>&sellyn=<%=sellyn%>&sailyn=<%=sailyn%>&disp=<%=dispcate%>'; 
		</script>
	<%
	response.End 	
	 
 	
 Case "U"
 
 etcitemimg = requestCheckVar(Request("hiddf"),128)
 evtmolistbanner= requestCheckVar(Request("hidwb"),128)
 evtNmW= requestCheckVar(Request("evtNmW"),128)
 evtNmM= requestCheckVar(Request("evtNmM"),128)
 subcopyK= requestCheckVar(Request("subcopyK"),128)
 evtsubname= requestCheckVar(Request("evtsubname"),128)
 mdtheme= requestCheckVar(Request("mdTm"),1) 
 thmcolor= requestCheckVar(Request("tmc"),10)
 thmcolormo= requestCheckVar(Request("tmcmo"),10)
 txtbgcolor= requestCheckVar(Request("tbgc"),10) 
 gUsing= requestCheckVar(Request("gUsing"),4)
 gtext1= requestCheckVar(Request("gtext1"),128)
 gimg1= requestCheckVar(Request("hidg1"),128)
 gtext2= requestCheckVar(Request("gtext2"),128)
 gimg2= requestCheckVar(Request("hidg2"),128)
 gtext3= requestCheckVar(Request("gtext3"),128)
 gimg3= requestCheckVar(Request("hidg3"),128)
 
  	strSql = "update db_event.dbo.tbl_partner_event "
  	strSql = strSql & " set  "
  	strSql = strSql & " evt_name ='"&evtNm&"' , evt_startdate='"&evtSD&"', evt_enddate ='"&evtED&"', evt_lastupdate =getdate() "
	strSql = strSql & ",[evt_dispcate]="&dispCate&",[issale]="&evtSale&",[isgift]="&evtGift&",[iscoupon]="&evtCoupon&",[brand]='"&makerid&"',[evt_tag]='"&evtTag&"' "
	strSql = strSql & " ,saleper ='"&evtSalePer&"' , salecper='"&evtsalecper&"' "
  	strSql = strSql & " ,etc_itemimg ='"&etcitemimg&"' , evt_mo_listbanner  ='"&evtmolistbanner&"', title_pc ='"&evtNmW&"', title_mo ='"&evtNmM&"', evt_subcopyK ='"&subcopyK&"',evt_subname ='"&evtsubname&"'"
	strSql = strSql & "	,mdtheme ='"&mdtheme&"',themecolor ='"&thmcolor&"',themecolormo ='"&thmcolormo&"',textbgcolor ='"&txtbgcolor&"'"
	strSql = strSql & "	,gift_isusing ='"&gUsing&"',gift_text1 ='"&gtext1&"',gift_img1 ='"&gimg1&"',gift_text2 ='"&gtext2&"',gift_img2 ='"&gimg2&"',gift_text3 ='"&gtext3&"',gift_img3 ='"&gimg3&"'"
	strSql = strSql & "	from db_event.dbo.tbl_partner_event "
  	strSql = strSql & "  where  evt_Code  ="&evtCode 
	 dbget.Execute strSql
 		 
 		 if mdtheme =2 THEN
 		 	slideimg1 =requestCheckVar(Request("hidtms1"),128) 
 		 	slideimg2 =requestCheckVar(Request("hidtms2"),128) 
 		 	slideimg3 =requestCheckVar(Request("hidtms3"),128) 
 		 	slideimgm1 =requestCheckVar(Request("hidtmsm1"),128) 
 		 	slideimgm2 =requestCheckVar(Request("hidtmsm2"),128) 
 		 	slideimgm3 =requestCheckVar(Request("hidtmsm3"),128) 
 		  
 		 	if slideimg1 <> "" then
 		 		strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =1 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimg1&"', isusing='Y' where evt_code ="&evtCode&" and device='W' and sorting =1 "  &vbCrlf
 		 		strSql = strSql &"ELSE"  &vbCrlf
	 		 	strSql = strSql & "INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'W','"&slideimg1&"',1)"
	 			dbget.Execute strSql	 	
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =1 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='W' and sorting =1 "  &vbCrlf	
 		 			dbget.Execute strSql	 	
	 		end if
	 		if slideimg2 <> "" then
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =2 )"  &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimg2&"' where evt_code ="&evtCode&" and device='W' and sorting =2"  &vbCrlf
 		 		strSql = strSql &"ELSE"  &vbCrlf
	 		 	strSql = strSql & "INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'W','"&slideimg2&"',2)"
	 			dbget.Execute strSql	 	
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =2 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='W' and sorting =2 "  &vbCrlf		
 		 		dbget.Execute strSql	 	
	 		end if
	 		if slideimg3 <> "" then
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =3 )"  &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimg3&"' where evt_code ="&evtCode&" and device='W' and sorting =3 "  &vbCrlf
 		 		strSql = strSql &"ELSE"  &vbCrlf
	 		 	strSql = strSql &"INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'W','"&slideimg3&"',3)"
	 			dbget.Execute strSql	 
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='W' and sorting =3 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='W' and sorting =3 "  &vbCrlf		
 		 		dbget.Execute strSql	 			
	 		end if
	 			if slideimgm1 <> "" then
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='M' and sorting =1 )"  &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimgm1&"' where evt_code ="&evtCode&" and device='M' and sorting =1 "  &vbCrlf
 		 		strSql = strSql &"ELSE" 	 &vbCrlf
	 		 	strSql = strSql &"INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'M','"&slideimgm1&"',1)" &vbCrlf 
	 			dbget.Execute strSql	 	
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='M' and sorting =1 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='M' and sorting =1 "  &vbCrlf		
 		 		dbget.Execute strSql	 		
	 		end if
	 		if slideimgm2 <> "" then
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='M' and sorting =2 )"  &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimgm2&"' where evt_code ="&evtCode&" and device='M' and sorting =2 "  &vbCrlf
 		 		strSql = strSql &"ELSE"  &vbCrlf
	 		 	strSql = strSql & "INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'M','"&slideimgm2&"',2)" 
	 			dbget.Execute strSql	 	
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='M' and sorting =2 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='M' and sorting =2 "  &vbCrlf		
 		 		dbget.Execute strSql	 		
	 		end if
	 		if slideimgm3 <> "" then
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device= 'M' and sorting =3 )"  &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET slideimg ='"&slideimgm3&"' where evt_code ="&evtCode&" and device='M' and sorting =3 "  &vbCrlf
 		 		strSql = strSql &"ELSE"  &vbCrlf
	 		 	strSql = strSql &"INSERT INTO db_event.dbo.tbl_partner_event_slide_addimage (evt_code, device, slideimg, sorting )" &vbCrlf
	 		 	strSql = strSql & " VALUES ("&evtCode&" ,'M','"&slideimgm3&"',3)" 
	 			dbget.Execute strSql	 	
	 		ELSE
	 			strSql ="IF EXISTS(SELECT idx FROM db_event.dbo.tbl_partner_event_slide_addimage where evt_code ="&evtCode&" and device='M' and sorting =3 )" &vbCrlf
 		 		strSql = strSql &"UPDATE db_event.dbo.tbl_partner_event_slide_addimage  SET isusing ='N' where evt_code ="&evtCode&" and device='M' and sorting =3 "  &vbCrlf		
 		 		dbget.Execute strSql	 		
	 		end if
	 	end if	 
 	 response.redirect "/admin/eventmanage/wait/contEvent.asp?menupos="&menupos&"&eC="&evtCode
 	response.end	
Case "TB" 	'상품 테마 등록
	
	if Request.Form("cksel") <> "" then
		if checkNotValidHTML(Request.Form("cksel")) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
  
'배열로 처리
redim arritemcode(vCnt)

for i=1 to vCnt
	arritemcode(i) = requestCheckVar(Request.Form("cksel")(i),10)
 
		strSql = " IF Not Exists(SELECT IDX FROM [db_event].[dbo].[tbl_partner_event_itembanner] WHERE itemid='" & arritemcode(i) & "' and evt_code="&evtCode&" and sdiv='"&sdiv&"')"			
		strSql = strSql + "	BEGIN "
		strSql = strSql+ " 			INSERT INTO [db_event].[dbo].[tbl_partner_event_itembanner] (evt_code, sdiv, itemid,  viewidx)"
		strSql = strSql + "     	VALUES (" & evtCode & ", '" & sdiv & "', " & arritemcode(i) &"," & i & ")"
		strSql = strSql + " 	END "
		strSql = strSql + " ELSE "
		strSql = strSql + " 	BEGIN "			
		strSql = strSql + "			UPDATE [db_event].[dbo].[tbl_partner_event_itembanner]"
		strSql = strSql + " 		SET viewidx ='" & i & "'"
		strSql = strSql + " 		WHERE evt_code = '" & evtCode & "' "
		strSql = strSql + " 		and itemid ="&arritemcode(i)&" and sdiv ='"&sdiv&"'"
		strSql = strSql + " 	END " 
		dbget.execute strSql
	
 next
 %>
 	<script>	 
		parent.location.href='/partner/event/plan/popRegItem.asp?ec=<%=evtCode%>&sdiv=<%=sdiv%>&page=<%=page%>&itemname=<%=itemname%>&itemid=<%=itemid%>&sellyn=<%=sellyn%>&sailyn=<%=sailyn%>&disp=<%=dispcate%>'; 
		</script>
	<%
 	response.end

CASE "TD" '상품테마 상품삭제
 
 arritem = requestCheckVar( Request.Form("delid") ,100)
	strSql = " delete FROM [db_event].[dbo].[tbl_partner_event_itembanner] WHERE evt_code=" & evtCode & " and sdiv='" & sdiv & "' and itemid in (" &arritem& ")" 
	dbget.execute strSql

%>
	<script>	
		alert('삭제 되었습니다.');		
		parent.location.href='/admin/eventmanage/wait/popRegItem.asp?ec=<%=evtCode%>&sdiv=<%=sdiv%>&page=<%=page%>&itemname=<%=itemname%>&itemid=<%=itemid%>&sellyn=<%=sellyn%>&sailyn=<%=sailyn%>&disp=<%=dispcate%>'; 
		</script>
	<%
	response.End  
Case "TS" '상품테마 미리보기세팅
	dim ClsEvt
	dim arrimg
set ClsEvt = new CEvent 
	ClsEvt.FevtCode = evtCode
 	ClsEvt.Fsdiv =sdiv
 	arrimg 		= ClsEvt.fnGetEventItemImg
set ClsEvt = nothing
	%>
	<div id="pvSlide">
		<%  if sdiv ="W" then %>
		<div class="slide">
		<%
			if isArray(arrimg) then
				for intLoop = 0 To UBound(arrimg,2)
		%>
			<div><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrimg(0,intLoop)%>" ></div>		
		<%	next
			end if
	 	%>
		</div>
		<%	else%>
		<div class="swiper-container">
			<div class="swiper-wrapper" id="tmsmImg">  	
			 <%
			if isArray(arrimg) then
				for intLoop = 0 To UBound(arrimg,2)
			%>
				<div class="swiper-slide">																				
					<div class="thumbnail"><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrimg(0,intLoop)%>" ></div>																			
				</div>
			<%	next
			end if
	 		%>
			</div> 
		<div class="pagination-line"></div>
			<button type="button" class="btnNav btnPrev">이전</button>
			<button type="button" class="btnNav btnNext">다음</button>
		</div>
		<%end if%>
	</div>
	<script type="text/javascript" src="/js/jquery.slides.min2.js"></script>
	<script type="text/javascript">
		var sValue = $("#pvSlide").html();
		<%if sdiv ="W" then%>
			 $(opener.document).find(".fullTemplatev17 .fullContV17 .slide").remove(); 	
	 		 $(opener.document).find(".fullTemplatev17 .fullContV17").append(sValue);   
			opener.jsRollingbg('C');
		 <%else%> 
		 	 $(opener.document).find("#mdRolling .swiper-container").remove(); 
		 	 $(opener.document).find("#mdRolling").append(sValue);   
		 	opener.jsRollingbgM('C'); 
		 <%end if%>
 self.close();
	</script>
	<%
response.end
 
Case "R" '반려
dbget.begintrans
	
	etext =  requestCheckVar( Request.Form("etext") ,200)	 
	strSql = "UPDATE db_event.dbo.tbl_partner_event  set evt_state=3, evt_lastupdate =getdate() " 
	strSql = strSql & " where evt_code ="&evtCode&" and evt_state = 5 "
	dbget.execute strSql
	
	strSql = "INSERT INTO db_event.dbo.tbl_partner_eventStateLog (evt_code, evt_state, evt_text, evt_manager, regid)"
	strSql = strSql & " VALUES ("&evtCode&",3, '"&etext&"',"&evtManager&",'"&adminid&"')"
	dbget.execute strSql
	
		IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
		response.End 
	END IF
	dbget.committrans
	
	Call Alert_move("반려처리를 완료하였습니다.","/admin/eventmanage/wait/?menupos="&menupos&"&eC="&evtCode)
Case "C" '승인
	dim eScope,elevel  ,eState,sOpenDate,sCloseDate,sImgregdate,isWeb,isMobile,isApp,eISort
	dim blnFull,blnWide,blnIteminfo,blnItemprice,eDateView
	dim evtCodeR
	dim eSalePer, eSaleCPer
	dim estrSale
	eSalePer = requestCheckVar(Request("eSP"),8)
	eSaleCPer = requestCheckVar(Request("eCP"),8) 
  if eSalePer <> "" or eSalePer <>"0" THEN
  	estrSale = "|"&eSalePer
  elseif 	eSaleCPer <> "" or eSaleCPer <> "0" then
  	estrSale = "|"&eSaleCPer
  end if
	eScope 	 =2
	elevel = 2 '중요도 보통으로 임시 설정
	eState 		= requestCheckVar(Request("eventstate"),4)
	sOpenDate = "null"
	sCloseDate = "null"
	sImgregdate = "null"
	
	IF eState = 7 THEN
		sOpenDate = "getdate()"
	ELSEIF eState = 9 THEN
		sCloseDate = "getdate()"
	ELSEIF eState = 3 THEN
	    sImgregdate = "getdate()"	
	END IF
	isWeb =1
	isMobile =1
	isApp =1
	eISort = 3'지정번호순
	blnFull = 1
  	blnWide = 0
  	blnIteminfo =1
  	blnItemprice = 0
  	eDateView = 0
  	strSql =" select realevt_code FROM db_event.dbo.tbl_partner_event where  evt_code ="&evtCode
  		rsget.Open strSql, dbget, 1
  		if  rsget.eof then
  		evtCodeR = rsget(0)
  		end if
  		rsget.close
  		
  		if evtCodeR <> "" and evtCodeR<>"0" then
  			Call sbAlertMsg ("이미 등록된 이벤트 코드입니다. 확인해주세요.[0]", "back", "")
  	  end if
  	
	dbget.begintrans
	'--1.master등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event] (evt_kind, evt_manager, evt_scope, evt_name, evt_startdate, evt_enddate, evt_prizedate,  evt_level, evt_state, evt_regdate, opendate, closedate, evt_lastupdate, adminid,evt_subcopyK,evt_sortNo , evt_subname, isWeb, isMobile, isApp ,evt_imgregdate, evt_type, isConfirm) "&vbCrlf		
		strSql = strSql &" ( SELECT evt_kind,evt_manager,"&escope&",evt_name+'"&estrSale&"',evt_startdate,evt_enddate,'', "&eLevel&","&eState&",evt_regdate,"&sOpenDate&","&sCloseDate&",getdate(),adminid,evt_subcopyK,0,evt_subname,"&isWeb&","&isMobile&","&isApp&","&sImgregdate&",80,0"&vbCrlf
		strSql = strSql &" FROM db_event.dbo.tbl_partner_event "&vbCrlf
      	strSql =strSql &" WHERE evt_code ="&evtCode&" and evt_state = 5 and evt_using ='Y' ) " 
		dbget.execute strSql
		
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
		response.End 
	END IF
				
		strSql = "select SCOPE_IDENTITY()"
		rsget.Open strSql, dbget, 0
		evtCodeR = rsget(0)
		rsget.Close

		'--2.disply등록 
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_display] "&vbCrlf
		strSql = strSql &		" (evt_code, evt_dispCate, brand,evt_template,evt_template_mo,partMdid, designerid "&vbCrlf
		strSql = strSql &		"	,issale,isgift,iscoupon,isOnlyTen,isOneplusone,isFreedelivery,isbookingsell, isDiary,isNew,iscomment,isbbs,isitemps,isapply,isGetBlogURL "&vbCrlf
		strSql = strSql &		"	,evt_itemsort,evt_tag, evt_fullyn, evt_wideyn, evt_iteminfoyn, evt_itempriceyn,evt_dateview,etc_itemimg, evt_mo_listbanner "&vbCrlf
		strSql = strSql &		"	, SalePer, SaleCPer, mdtheme, mdthememo, themecolor, themecolormo, textbgcolor, textbgcolormo, mdbntype, mdbntypemo,evt_itemlisttype, eventtype_pc, eventtype_mo)" & vbCrlf
		strSql = strSql & " ( SELECT "&evtCodeR&",evt_dispcate,brand,9,9,'"&adminid&"','' "&vbCrlf
		strSql = strSql & " ,issale,isgift,iscoupon,0,0,0,0,0,0,0,0,0,0,0 "&vbCrlf
      	strSql = strSql &"		,"&eISort&"	,evt_tag,"&blnFull&","&blnWide&","&blnIteminfo&","&blnItemprice&",'"&eDateView&"',etc_itemimg,evt_mo_listbanner "&vbCrlf
      	strSql = strSql & ",'"&eSalePer&"','"&eSaleCPer&"',isNull(mdtheme,1) as mdtheme,isNull(mdtheme,1) as mdthememo, themecolor,themecolormo,textbgcolor, textbgcolor as textbgcolormo"&vbCrlf
      	strSql = strSql & "	,case mdtheme when '2' then 'D' when '3' then 'T' else '' end as mdbntype ,case mdtheme when '2' then 'D' when '3' then 'T' else '' end as mdbntypemo ,'1', 80, 80  " &vbCrlf
      	strSql = strSql &" FROM db_event.dbo.tbl_partner_event "&vbCrlf
      	strSql =strSql &" WHERE evt_code ="&evtCode&" and evt_state = 5 and evt_using ='Y') "     
		dbget.execute strSql				
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
		response.End 
	END IF	

		'--3.MD 등록 테마 정보 등록
		strSql = "INSERT INTO [db_event].[dbo].[tbl_event_md_theme] (evt_code, comm_isusing, comm_start,comm_end, gift_isusing, gift_text1, gift_img1, gift_text2, gift_img2, gift_text3, gift_img3, usinginfo, title_pc, title_mo) "&vbCrlf 
		strSql = strSql & " ( SELECT "&evtCodeR&" , 'N','','', gift_isusing,gift_text1,gift_img1,gift_text2,gift_img2,gift_text3,gift_img3,0, title_pc,title_mo "&vbCrlf 
		strSql = strSql &" FROM db_event.dbo.tbl_partner_event "&vbCrlf
      	strSql =strSql &" WHERE evt_code ="&evtCode&" and evt_state = 5  and evt_using ='Y') "      
      	 	
		dbget.execute strSql
		
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[3]", "back", "")
		response.End 
	END IF
 
	'--4. 그룹등록
	dim arrgroup ,evtGCodeR, tmpGPcode
	dim tmpgdesc,tmpgsort,tmpgdepth, tmpgcode
	'//최상위 코드는 무조건 한개로 
	strSql ="SELECT top 1 evtgroup_code, evtgroup_desc, evtgroup_sort, evtgroup_depth, evtgroup_pcode FROM db_event.dbo.tbl_partner_eventitem_group WHERE  evt_code ="&evtCode&" and evtgroup_using='Y' and evtgroup_pcode = 0 "
	  rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			 tmpgdesc =rsget("evtgroup_desc")
			 tmpgsort =rsget("evtgroup_sort")
			 tmpgdepth=rsget("evtgroup_depth")
			 tmpgcode=rsget("evtgroup_code")
		END IF
		rsget.Close	
		
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort,evtgroup_depth, evtgroup_pcode"&vbCrlf
		strSql = strSql & ", evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_depth_mo,evtgroup_pcode_mo,evtgroup_linkkind,evtgroup_linkkind_mo) "&vbCrlf
		strSql = strSql & " values ("&evtCodeR&", '"&tmpgdesc&"','"&tmpgsort&"',"&tmpgdepth&", 0 "&vbCrlf
		strSql = strSql & " ,  '"&tmpgdesc&"','"&tmpgsort&"',"&tmpgdepth&", 0,5,5 )" 
		dbget.execute strSql 
		
		strSql = "select SCOPE_IDENTITY()"
		rsget.Open strSql, dbget,0
		tmpGPcode = rsget(0)
		rsget.Close
		
			strSql = "update db_event.dbo.tbl_partner_eventitem_group set realevtg_code = "&tmpGPcode&" where evtgroup_code="&tmpgcode
    	dbget.execute strSql 
    	
	  strSql ="SELECT evtgroup_code, evtgroup_desc, evtgroup_sort, evtgroup_depth, evtgroup_pcode FROM db_event.dbo.tbl_partner_eventitem_group WHERE  evt_code ="&evtCode&" and evtgroup_using='Y' and evtgroup_pcode <> 0 "
	  rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			arrgroup = rsget.getrows()
		END IF
		rsget.Close	
		
		if isArray(arrgroup) then
			for intLoop = 0 To UBound(arrgroup,2)
			
		strSql = "INSERT INTO [db_event].[dbo].[tbl_eventitem_group] (evt_code,evtgroup_desc, evtgroup_sort,evtgroup_depth, evtgroup_pcode"&vbCrlf
		strSql = strSql & ", evtgroup_desc_mo, evtgroup_sort_mo,evtgroup_depth_mo,evtgroup_pcode_mo) "&vbCrlf
		strSql = strSql & " values ("&evtCodeR&", '"&arrgroup(1,intLoop)&"','"&arrgroup(2,intLoop)&"',"&arrgroup(3,intLoop)&", "&tmpGPcode& vbCrlf
		strSql = strSql & " , '"&arrgroup(1,intLoop)&"','"&arrgroup(2,intLoop)&"', "&arrgroup(3,intLoop)&",  "&tmpGPcode&" )" 
		dbget.execute strSql 
		
		strSql = "select SCOPE_IDENTITY()"
		rsget.Open strSql, dbget,0
		evtGCodeR = rsget(0)
		rsget.Close
		
		strSql = "update db_event.dbo.tbl_partner_eventitem_group set realevtg_code = "&evtGCodeR&" where evtgroup_code="&arrgroup(0,intLoop)
    	dbget.execute strSql
    	next
    	end if
		
		strSql = "update [db_event].[dbo].[tbl_eventitem_group] set evtgroup_code_mo = evtgroup_code where evt_code = "&evtCodeR
		dbget.execute strSql 
		
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[3]", "back", "")
		response.End 
	END IF
	 
	'--5. 상품등록
	strSql =" insert into db_event.dbo.tbl_eventitem (evt_code, itemid, evtgroup_code, evtitem_sort, evtitem_imgsize, evtitem_sort_mo)"&vbCrlf
	strSql = strSql & " SELECT "&evtCodeR&",i.itemid, g.realevtg_code, i.evtitem_sort, i.evtitem_imgsize, i.evtitem_sort as evtitem_sort_mo "&vbCrlf
	strSql = strSql &  " FROM db_event.dbo.tbl_partner_eventitem as i "&vbCrlf
	strSql = strSql &  " inner join db_event.dbo.tbl_partner_eventitem_group as g on i.evt_code = g.evt_code and i.evtgroup_code = g.evtgroup_code "&vbCrlf
	strSql = strSql & " WHERE i.evt_code ="&evtCode&" and i.evtitem_isusing =1 and g.evtgroup_using ='Y' "
	dbget.execute strSql
	
	'--6. 이미지 테마 배너등록
	strSql ="INSERT INTO db_event.[dbo].[tbl_event_slide_addimage]([evt_code],[device],[slideimg],[sorting]) "&vbCrlf
     strSql = strSql & " SELECT "&evtCodeR&", device,slideimg, sorting "&vbCrlf
     strSql = strSql & " FROM db_event.dbo.tbl_partner_event_slide_addimage "&vbCrlf
    strSql = strSql & " WHERE evt_code ="&evtCode& " and isusing ='Y'" 
    	dbget.execute strSql
    	
  	'--6. 이미지 테마 배너등록
	strSql ="INSERT INTO db_event.dbo.tbl_event_slide_template (evt_code, device, topimg, btmYN, btmimg, btmcode, topaddimg, btmaddimg, pcadd1,  gubun) "&vbCrlf
     strSql = strSql & " SELECT "&evtCodeR&", device, '','','','','','','',0 "&vbCrlf
     strSql = strSql & " FROM db_event.dbo.tbl_partner_event_slide_addimage "&vbCrlf
    strSql = strSql & " WHERE evt_code ="&evtCode& " and isusing ='Y'" 
    	dbget.execute strSql
    	
    	  	
	'--7. 상품테마 상품등록
	strSql = "INSERT INTO db_event.[dbo].[tbl_event_itembanner] ([evt_code],[sdiv],[itemid],[itemname],[viewidx])"
	strSql = strSql & " SELECT  "&evtCodeR&",sdiv,e.itemid,i.itemname,e.viewidx "
	strSql = strSql & " FROM db_event.[dbo].[tbl_partner_event_itembanner] as e"
	strSql = strSql & " 	INNER JOIN db_item.dbo.tbl_item as i on e.itemid = i.itemid "
	strSql = strSql & " WHERE evt_code ="&evtCode
	dbget.execute strSql
	'-- 파트너 정보 업데이트
	strSql = "UPDATE db_event.dbo.tbl_partner_event  set evt_state=7,  realevt_code = "&evtCodeR&" ,evt_lastupdate =getdate() "
	strSql = strSql & " where evt_code ="&evtCode&" and evt_state = 5 "
	dbget.execute strSql
	
	etext = "승인되었습니다."
	strSql = "INSERT INTO db_event.dbo.tbl_partner_eventStateLog (evt_code, evt_state, evt_text, evt_manager, regid)"
	strSql = strSql & " VALUES ("&evtCode&",7, '"&etext&"',"&evtManager&",'"&adminid&"')"
	dbget.execute strSql
	
		IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
		response.End 
	END IF
	dbget.committrans
	dim vSCMChangeSQL
	vChangeContents = vChangeContents & "이벤트 INSERT " & vbCrLf
	vChangeContents = vChangeContents & "- 이벤트명 :  evt_code = " & evtCodeR & vbCrLf
	vChangeContents = vChangeContents & "- 파트너이벤트 승인 partenr_evt_code = "&evtCode& vbCrLf 

    	'### 수정 로그 저장(event)
    	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
    	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'event', '" & evtCodeR & "', '" & menupos & "', "
    	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
    	dbget.execute(vSCMChangeSQL)
    	
	Call Alert_move("승인되었습니다.","/admin/eventmanage/wait/?menupos="&menupos&"&eC="&evtCode)
case  Else
	 	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.-error: case else")
END select
 	
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->