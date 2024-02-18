<!--METADATA TYPE="typelib" NAME="ADODB Type Library" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

'######## Class #################


class CEvaluateSearcherItem
	public FId
	public FUserID
	public FPoint
	public FUesdContents
	public FPoint_fun
	public FPoint_dgn
	public FPoint_prc
	public FPoint_stf
	public Fimgsmall
	public FItemID
	public Fimglist

	public FRegdate

	public Fusingtitle
	public Flinkimg1
	public Flinkimg2

	public FBrandname
	public Fgubun
	public FItemname

	public function getUsingTitle()
		if (isNULL(Fusingtitle) or (Fusingtitle="")) then
			getUsingTitle = FUesdContents
			if Len(getUsingTitle)>40 then
				getUsingTitle = Left(getUsingTitle,40) + "..."
			end if
		else
			getUsingTitle = Fusingtitle
		end if
	end function

	public function getFingersUsingTitle()
		if (isNULL(Fusingtitle) or (Fusingtitle="")) then
			getFingersUsingTitle = FUesdContents
			if Len(getUsingTitle)>30 then
				getFingersUsingTitle = Left(getUsingTitle,30) + "..."
			end if
		else
			getUsingTitle = Fusingtitle
		end if
	end function

	public function IsPhotoExist()
		IsPhotoExist = (Flinkimg1<>"") or (Flinkimg2<>"")
	end function


	Private Sub Class_Terminate()

	End Sub

	public sub Class_Initialize()

	end sub
end Class

Class CEvaluateSearcher
	public FItemList()

	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectUserID
	public FRectItemID
	public FDiscountRate
	public FRectStarN
	public OrderMethod
	public FPointValue
	public FBrandValue
	public FCateValue
	public FDateYn
	public FRectDate1
	public FRectDate2
	public FRectEventId

	Private Sub Class_Initialize()
		redim preserve FItemList(0)

		FCurrPage     = 1
		FPageSize     = 12
		FResultCount  = 0
		FScrollCount  = 10


		FDiscountRate = 1
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public function GetImageFolerName(byval i)
		GetImageFolerName = "0" + CStr(Clng(FItemList(i).FItemID\10000))
	end function

	public sub getItemEvalList()
		dim sql,i,othersql


		sql = " select count(u.id) as cnt " &_
					" from [db_board].[10x10].tbl_user_goodusing u " &_
					" join [db_item].[10x10].tbl_item i ON i.itemid=u.itemid "

		if FRectEventId<>"" then
		sql = sql + " join (select itemid " &_
								" 		from db_contents.dbo.tbl_event_detail " &_
								" 		where masteridx in (" & FRectEventId & ")) as ed  " &_
								" 	on ed.itemid=u.itemid "
		end if

		sql = sql +	" where u.isdelete='N' and u.itemid<>'0' "

		if (FRectStarN<>"") then
			sql = sql + " and u.tpoint>=" + Cstr(FRectStarN)
		end if

		if FRectItemid<>"" then
			sql = sql + " and u.itemid='" + CStr(FREctItemid) + "'" + vbcrlf
		end if
		if FPointValue<>"" then
			sql = sql + " and u.tpoint='" + CStr(FPointValue) + "'" + vbcrlf
		end if

		if FBrandValue<>"" then
			sql=sql + " and i.makerid='" + CStr(FBrandValue) + "'" + vbcrlf
		end if

		if FCateValue<>"" then
			sql = sql + " and i.itemserial_large='" + CStr(FCateValue) + "'" + vbcrlf
		end if

		if FDateYn="on" then
			if FRectDate1 <> "" and FRectDate2 <> "" then
				sql = sql + " and u.regdate between '" + FRectDate1 + "' and '" + FRectDate2 + "'"
			end if
		end if

		rsget.open sql,dbget,1

		if not rsget.eof then
			FTotalCount = rsget("cnt")
		end if
		rsget.Close


		sql = " SELECT top " + CStr(FPageSize*FCurrPage) + " u.id, u.itemid, u.userid, u.gubun, i.itemname " &_
					" , IsNULL(u.tpoint,0) as point, u.contents, IsNULL(u.point1,0) as point_fun " &_
					" , IsNULL(u.point2,0) as point_dgn, IsNULL(u.point3,0) as point_prc " &_
					" , IsNULL(u.point4,0) as point_stf " &_
					" , convert(varchar(10),u.regdate,21) as regdate " &_
					" , IsNULL(u.file1,'') as linkimg1, IsNULL(u.file2,'') as linkimg2, i.smallimage  " &_
					" FROM [db_board].[10x10].tbl_user_goodusing u " &_
					" JOIN [db_item].[10x10].tbl_item i ON i.itemid=u.itemid "
		if FRectEventId<>"" then
		sql = sql + " JOIN (SELECT itemid " &_
								" 		from db_contents.dbo.tbl_event_detail " &_
								" 		WHERE masteridx in (" & FRectEventId & ")) as ed  " &_
								" 	on ed.itemid=u.itemid "
		end if

		sql = sql +	" WHERE u.isdelete='N' and u.itemid<>'0' "


		if (FRectStarN<>"") then
			sql = sql + " and u.point>=" + Cstr(FRectStarN)
		end if

		if FRectItemid<>"" then
			sql = sql + " and u.itemid='" + CStr(FREctItemid) + "'" + vbcrlf
		end if

		if FPointValue<>"" then
			sql = sql + " and u.tpoint='" + CStr(FPointValue) + "'" + vbcrlf
		end if

		if FBrandValue<>"" then
			sql = sql + " and i.makerid='" + CStr(FBrandValue) + "'" + vbcrlf
		end if

		if FCateValue<>"" then
			sql = sql + " and i.itemserial_large='" + CStr(FCateValue) + "'" + vbcrlf
		end if

		if FDateYn="on" then
			if FRectDate1 <> "" and FRectDate2 <> "" then
				sql = sql + " and u.regdate between '" + FRectDate1 + "' and '" + FRectDate2 + "'"
			end if
		end if

		sql = sql + " order by u.itemid desc, u.regdate desc "


		'response.write sql
		'dbget.close()	:	response.End
		rsget.pagesize = FPageSize
		rsget.open sql,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))



		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			i=0
			redim preserve FItemList(FResultCount)
			do until rsget.eof
				set FItemList(i) = new CEvaluateSearcherItem
				FItemList(i).FId					 = rsget("id")
				FItemList(i).FUserID       = rsget("userid")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FPoint        = rsget("point")
				FItemList(i).FItemname	=db2html(rsget("itemname"))
				FItemList(i).FUesdContents = db2html(rsget("contents"))
				FItemList(i).FPoint_fun       = rsget("point_fun")
				FItemList(i).FPoint_dgn        = rsget("point_dgn")
				FItemList(i).FPoint_prc        = rsget("point_prc")
				FItemList(i).FPoint_stf        = rsget("point_stf")

				FItemList(i).FRegdate 	= rsget("regdate")

				'FItemList(i).Fusingtitle = db2html(rsget("usingtitle"))

				FItemList(i).Fgubun = rsget("gubun")
				FItemList(i).Flinkimg1	= rsget("linkimg1")
				FItemList(i).Flinkimg2	= rsget("linkimg2")

				If FItemList(i).Fgubun = "01" then
					 if FItemList(i).Flinkimg1<>"" then
						 FItemList(i).Flinkimg1 = "http://webimage.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + FItemList(i).Flinkimg1
					 end if

					 if FItemList(i).Flinkimg2<>"" then
						 FItemList(i).Flinkimg2 = "http://webimage.10x10.co.kr/goodsimage/" + GetImageFolerName(i) + "/" + FItemList(i).Flinkimg2
					 end If
				elseIf FItemList(i).Fgubun = "02" then
					 if FItemList(i).Flinkimg1<>"" then
						 FItemList(i).Flinkimg1 = "http://webimage.10x10.co.kr/contimage/album/" + FItemList(i).Flinkimg1
					 end if

					 if FItemList(i).Flinkimg2<>"" then
						 FItemList(i).Flinkimg2 = "http://webimage.10x10.co.kr/contimage/album/" + FItemList(i).Flinkimg2
					 end If
				elseIf FItemList(i).Fgubun = "03" then
					 if FItemList(i).Flinkimg1<>"" then
						 FItemList(i).Flinkimg1 = "http://webimage.10x10.co.kr/contimage/maniaimg/evaluate/file1/" + FItemList(i).Flinkimg1
					 end if

					 if FItemList(i).Flinkimg2<>"" then
						 FItemList(i).Flinkimg2 = "http://webimage.10x10.co.kr/contimage/maniaimg/evaluate/file2/" + FItemList(i).Flinkimg2
					 end If
				End If

				FItemList(i).Fimgsmall        = "http://webimage.10x10.co.kr/image/small/" + GetImageFolerName(i) + "/" + rsget("smallimage")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
End Class


'########        내용  		 #################


dim evaluates,ix,i,ordermethod,frectval,designer,eventid
dim pointvalue,catevalue,brandvalue,citemid,dateYN
Dim page
page = request("page")
If page = "" Then page=1

citemid=request("citemid")
pointvalue=request("pointvalue")
eventid= request("eventid")
if citemid="" then
brandvalue=request("brandvalue")
catevalue=request("catevalue")
end if

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate,toDate
dateYN=request("dateYN")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))


set evaluates = new CEvaluateSearcher

evaluates.FPageSize = 30
evaluates.FCurrpage=page
evaluates.FRectItemid=citemid
evaluates.FPointValue=pointvalue
evaluates.FBrandValue=brandvalue
evaluates.FCateValue=catevalue
evaluates.FDateYn=dateYN
evaluates.FRectDate1=fromDate
evaluates.FRectDate2=toDate
evaluates.FRectEventId=eventid
evaluates.getItemEvalList


Sub SelectBoxDesigner(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select name="<%=selectBoxName%>" onchange="javascript:delid();">
     <option value='' <%if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select userid,socname_kor from [db_user].[10x10].tbl_user_c  order by userid"
   ''query1 = query1 & " where a.userid = b.userid order by userid "
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"& rsget("userid") & "[" & replace(db2html(rsget("socname_kor")),"'","") & "]</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
<script language='javascript'>
function showhide(num, p_totcount)    {
	var imaxwidth = 600;

  	for (i=0; i<p_totcount; i++)   {
	  var menu=eval("document.all.evalu_block_"+i+".style");
	  if (num==i ) {
		if (menu.display=="block"){
			menu.display="none";
		}else{
		  menu.display="block";
		}
	  }else{
		menu.display="none";
	  }
	}

	FixImageSizeAll(600,"image_fix_1");
	FixImageSizeAll(600,"image_fix_2");
}

function showhide2(num, p_totcount)    {
	var imaxwidth = 600;

  	for (i=0; i<p_totcount; i++)   {
	  var menu=eval("document.all.photo_block_"+i+".style");
	  if (num==i ) {
		if (menu.display=="block"){
			menu.display="none";
		}else{
		  menu.display="block";
		}
	  }else{
		menu.display="none";
	  }
	}

	FixImageSizeAll(600,"photo_fix_1");
	FixImageSizeAll(600,"photo_fix_2");
}

function showhide3(num, p_totcount)    {
	var imaxwidth = 600;

  	for (i=0; i<p_totcount; i++)   {
	  var menu=eval("document.all.mania_block_"+i+".style");
	  if (num==i ) {
		if (menu.display=="block"){
			menu.display="none";
		}else{
		  menu.display="block";
		}
	  }else{
		menu.display="none";
	  }
	}

	FixImageSizeAll(600,"mania_fix_1");
	FixImageSizeAll(600,"mania_fix_2");
	FixImageSizeAll(600,"mania_fix_3");
	FixImageSizeAll(600,"mania_fix_4");
	FixImageSizeAll(600,"mania_fix_5");
}

function FixImageSizeAll(imaxwidth,icompname){
	var iimglist = eval("document.all." + icompname);

	try {
		if (iimglist.width>imaxwidth){
			iimglist.width = imaxwidth;
		}
	}catch(e){

	};

	try {
		for (var i=0;i<iimglist.length;i++){
			if (iimglist[i].width>imaxwidth){
				iimglist[i].width = imaxwidth;
		   }
		}
	}catch(e){

	}

}

function isArray(obj){
  return obj instanceof Array
}

function GPgae(page){
	document.orderfrm.page.value=page;
	document.orderfrm.submit();
}

function orderpage(){
	document.orderfrm.page.value='';
	document.orderfrm.submit();
}

function delid(){
	document.orderfrm.citemid.value='';
}

function AddBestEvaluate(iid){
	document.EvalFrm.iid.value=iid;
	document.EvalFrm.submit();

}
</script>
<link rel=stylesheet type="text/css" href="/css/tenten.css">
<script language="JavaScript" SRC="/js/tenbytencommon.js"></script>
	<table border="1" cellpadding="2" cellspacing="0" width="750" height="20">
	<tr>
		<td class="a">카테고리,브랜드,점수 동시에 검색 가능합니다.&nbsp;&nbsp;상품 코드,점수 동시에 검색 가능합니다.<br>
							<b><u>카테고리,상품코드 또는 브랜드,상품 코드 동시 검색 불가능 합니다. ^^;</u></b>
		</td>
	</tr>
	</table>
	<table border="0" cellpadding="2" cellspacing="0" width="750" height="50"  bgcolor="E6E6E6">
	<form name="orderfrm" method="get" action="">
		<input type="hidden" name="page" value="<%= page %>">
		<tr class="a">
			<td align="left" width="150">
				카테고리선택
			</td>
			<td align="left" width="100">
				브랜드 선택
			</td>
			<td align="left" width="180">
				점수 선택
			</td>
			<td align="left" width="80">
				상품 코드
			</td>
			<td rowspan="2" style="padding-top:9">
				<a href="javascript:orderpage();"><img src="http://webadmin.10x10.co.kr/admin/images/search2.gif" border="0"></a>
			</td>
		</tr>
		<tr class="a" valign="top">
			<td align="left">
				<select name="catevalue"  onchange="javascript:delid();">
					<option value="" <% if catevalue="" then response.write "selected" %>>전체</option>
					<option value="10" <% if catevalue="10" then response.write "selected" %>>디자인문구/개인소품</option>
					<option value="15" <% if catevalue="15" then response.write "selected" %>>인테리어/리빙데코</option>
					<option value="25" <% if catevalue="25" then response.write "selected" %>>주방/욕실/생활</option>
					<option value="30" <% if catevalue="30" then response.write "selected" %>>패션/잡화</option>
					<option value="40" <% if catevalue="40" then response.write "selected" %>>키덜트/얼리</option>
					<option value="20" <% if catevalue="20" then response.write "selected" %>>취미/여가</option>
					<option value="35" <% if catevalue="35" then response.write "selected" %>>주얼리</option>
					<option value="50" <% if catevalue="50" then response.write "selected" %>>플라워</option>
				</select>
			</td>
			<td align="left">
				<%	SelectBoxDesigner "brandvalue",brandvalue	%>
			</td>
			<td align="left">
				<input type="radio" name="pointvalue" <% if pointvalue="" then response.write "checked" %> value="">전체
				<input type="radio" name="pointvalue" <% if pointvalue="1" then response.write "checked" %> value="1">1<input type="radio" name="pointvalue" <% if pointvalue="2" then response.write "checked" %> value="2">2
				<input type="radio" name="pointvalue" <% if pointvalue="3" then response.write "checked" %> value="3">3<input type="radio" name="pointvalue" <% if pointvalue="4" then response.write "checked" %> value="4">4
			</td>
			<td align="left">
				<input type="text" name="citemid" value="<%= citemid %>" size="6">
			</td>
		</tr>
		<tr>
			<td colspan="5"><% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>(<font class="a">날짜검색 사용</font><input type="checkbox" name="dateYN" <% if DateYn="on" then response.write "checked" %> />)
				<font class="a">이벤트번호</font><input type="text" name="eventid" value="<%= eventid %>" /></td>
		</tr>
	</form>
	</table>
	<table  border="0" width="750" cellspacing="0" cellpadding="0">
        <tr valign="top">
          <td style="padding-top:5">
            <table border="0" width="750" height="61" cellpadding="0" cellspacing="0">
              <% if evaluates.FResultCount<1 then %>
              <tr>
              	<td height="60" align=center> <font color="0074EA">(해당 상품에 상품평이 없습니다.)</font></td>
              </tr>
              <% else %>
              <tr>
                <td>
                  <table border="0" width="100%" cellpadding="0" cellspacing="0">
                  <% for i = 0 to evaluates.FResultCount - 1 %>
						<% if i=0 then %>
							<tr>
								<td  width="100%" bgcolor="eeeeee">
									<img src="<%= evaluates.FItemList(i).Fimgsmall %>" border="0"><font class="a"><%= evaluates.FItemList(i).FItemName %>(상품번호: <font color="blue"><%= evaluates.FItemList(i).FItemID %>)</font></font>
								</td>
							</tr>
						<% else %>
							<% if evaluates.FItemList(i).FItemId<>evaluates.FItemList(i-1).FItemId then %>
								<tr><td width="100%" bgcolor="eeeeee">
									<img src="<%= evaluates.FItemList(i).Fimgsmall %>" border="0"><font class="a"><%= evaluates.FItemList(i).FItemName %>(상품번호: <font color="blue"><%= evaluates.FItemList(i).FItemID %>)</font></font>
								</td></tr>
							<% end if %>
						<% end if %>
                    <tr>
                      <td height="22">
                        <table style="border-bottom: 1px solid #d1d1d1" border="0" width="100%" cellpadding="3" cellspacing="0" height="28">
                        	<tr valign="middle">
                            <td width="64" class="verdana-small" align="left">
                            <% for ix=0 to evaluates.FItemList(i).FPoint -1 %><img src="http://www.10x10.co.kr/images/category/step.gif" width="9"><% next %>
                            </td>
                            <td>
                              <p><a href="javascript:showhide(<%= i %>,<%= evaluates.FResultCount %>);" onfocus="this.blur();"><font color="#666666"><%= evaluates.FItemList(i).getUsingTitle %></a></font></p>
                            </td>
                            <td width="69">
                            <% if evaluates.FItemList(i).IsPhotoExist then %>
                              <div align="center"><img src="http://www.10x10.co.kr/images/shopping/photo.gif" width="63" height="20"></div>
                            <% end if %>
                            </td>
                            <td width="75" class="verdana-small">
                              <div align="center"><font color="#666666"><%= evaluates.FItemList(i).FRegdate %></font></div>
                            </td>
                            <td width="75" class="verdana-small">
                              <div align="center"><font color="#333333"><span onclick="AddBestEvaluate('<%= evaluates.FItemList(i).FId %>');" style="cursor:pointer;"><%= evaluates.FItemList(i).FUserID %></span></font></div>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td>
                        <div id="evalu_block_<%= i %>" style="DISPLAY:none; xCURSOR:hand">
                          <table border="0" width="91%" cellpadding="0" cellspacing="7">
                            <tr>
                              <td width="120" valign="top">
                              <table width="120" border="0" cellspacing="0" cellpadding="0" class="a">
                              <tr>
                              	<td width="50">기능</td>
                              	<td><% for ix=0 to evaluates.FItemList(i).FPoint_fun -1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10"><% next %></td>
                              </tr>
                              <tr>
                              	<td width="50">디자인</td>
                              	<td><% for ix=0 to evaluates.FItemList(i).FPoint_dgn -1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10"><% next %></td>
                              </tr>
                              <tr>
                              	<td width="50">가격</td>
                              	<td><% for ix=0 to evaluates.FItemList(i).FPoint_prc -1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10"><% next %></td>
                              </tr>
                              <tr>
                              	<td width="50">만족도</td>
                              	<td><% for ix=0 to evaluates.FItemList(i).FPoint_stf -1 %><img src="http://www.10x10.co.kr/images/category/px_01.gif" width="10"><% next %></td>
                              </tr>
                              </table>
                              </td>
                              <td>
                              <font size="2" color="#666666">
                              <%= nl2br(evaluates.FItemList(i).FUesdContents) %>
                              </font>
                              </td>
                            </tr>
                            <% if evaluates.FItemList(i).Flinkimg1<>"" then %>
                            <tr>
                              <td width="100"></td>
                              <td><img name="image_fix_1" src="<%= evaluates.FItemList(i).Flinkimg1 %>" ></td>
                            </tr>
                            <% end if %>
                            <% if evaluates.FItemList(i).Flinkimg2<>"" then %>
                            <tr>
                              <td width="100"></td>
                              <td><img name="image_fix_2" src="<%= evaluates.FItemList(i).Flinkimg2 %>" ></td>
                            </tr>
                            <% end if %>
                          </table>
                        </div>
                      </td>
                    </tr>
                    <% next %>
                  </table>
                </td>
					</tr>
              <% end if %>
            </table>
          </td>
        </tr>
		  <tr>
				<td>
				 <% if evaluates.HasPreScroll then %>
					 <a href="javascript:GPgae('<%= evaluates.StarScrollPage-1 %>')">[pre]</a>
				 <% else %>
					 [pre]
				 <% end if %>

				 <% for i=0 + evaluates.StarScrollPage to evaluates.FScrollCount + evaluates.StarScrollPage - 1 %>
					 <% if i>evaluates.FTotalpage then Exit for %>
					 <% if CStr(page)=CStr(i) then %>
					 <font color="red">[<%= i %>]</font>
					 <% else %>
					 <a href="javascript:GPgae('<%= i %>')">[<%= i %>]</a>
					 <% end if %>
				 <% next %>

				 <% if evaluates.HasNextScroll then %>
					 <a href="javascript:GPgae('<%= i %>')">[next]</a>
				 <% else %>
					 [next]
				 <% end if %>
				<span align="right"><font class="a"  color="blue"> Total Page: <%= evaluates.FTotalPage %> / Total: <%= evaluates.FTotalCount %>건 </font></span>
				</td>
		  </tr>
      </table>
<form name="EvalFrm" method="post" target="FrameEval" action="do_Diary_event_evaluate.asp">
<input type="hidden" name="iid" value="" />
</form>
<iframe name="FrameEval" src="" frameborder="0" width="500" height="100"></iframe>
<%
set evaluates = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->