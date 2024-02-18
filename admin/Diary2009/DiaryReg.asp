<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  다이어리 리스트 어드민 신규등록 팝업
' History : 2015.09.14 유태욱 수정(상품코드로 찾기기능 추가,리미티드추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<!-- #include virtual="/admin/diary2009/lib/include_event_code.asp"-->
<%
dim Diaryid,mode, limited
Diaryid = request("id") 
Mode = request("mode")
dim oDiary

IF Mode = "" THEN Mode = "add"

dim CateCode,ItemID,RegDate,isUsing,ImageName,BasicImgUrl , commentyn ,commentImgName , commentImgUrl
dim event_code, eventgroup_code , event_start , event_end ,weight, ImageName2, ImageName3, MDPick, soonseo, StoryImg, BasicImg2Url, StoryImgUrl, storytext, nanumimg, nanumimgUrl, reservdate, mdpicksort
dim comStat, BasicImg3Url
MDPick = "x"
limited = "x"
mdpicksort = 0
IF Mode= "edit" THEN
	set oDiary = new DiaryCls
	oDiary.FRectDiaryID = Diaryid
	oDiary.getDiary()

	Diaryid = oDiary.FItem.FDiaryID
	CateCode = oDiary.FItem.FCateCode
	ItemID = oDiary.FItem.Fitemid
	RegDate = oDiary.FItem.FRegDate
	isUsing = oDiary.FItem.FisUsing

	commentyn = oDiary.FItem.fcommentyn
	commentImgName = oDiary.FItem.fcomment_img
	commentImgUrl = oDiary.FItem.Imgcomment
	eventgroup_code = oDiary.FItem.feventgroup_code
	event_code = oDiary.FItem.fevent_code
	event_start = oDiary.FItem.fevent_start
	event_end = oDiary.FItem.fevent_end
	weight = oDiary.Fitem.Fweight

	ImageName	= oDiary.FItem.FImg
	BasicImgUrl = oDiary.FItem.ImgBasic

	ImageName2 = oDiary.Fitem.FImg2
	BasicImg2Url = oDiary.FItem.ImgBasic2
	ImageName3 = oDiary.Fitem.FImg3
	BasicImg3Url = oDiary.FItem.ImgBasic3

	StoryImg = oDiary.Fitem.FImgStory
	StoryImgUrl = oDiary.Fitem.ImgStory

	MDPick = oDiary.Fitem.Fmdpick
	limited = oDiary.Fitem.Flimited
	soonseo = oDiary.Fitem.Fsoonseo
	storytext = oDiary.Fitem.FStoryText

	nanumimg = oDiary.Fitem.FImgNanum
	nanumimgUrl = oDiary.Fitem.ImgNanum

	reservdate = oDiary.Fitem.FReservdate
	If reservdate = "1900-01-01" Then
		reservdate = ""
	End IF

	mdpicksort = oDiary.FItem.Fsorting

	set oDiary = nothing

End IF
IF isUsing="" THEN isUsing="Y"
IF commentyn="" THEN commentyn="N"

IF commentyn="Y" and eventgroup_code<>"" and event_start<=datevalue(now) and datevalue(now) <= event_end Then
	comStat="오픈"
ELSEIF commentyn="Y" and eventgroup_code<>"" and datevalue(now) > event_end Then
	comStat ="종료"
'ELSEIF commentyn="Y" and eventgroup_code<>"" and datevalue(now) < event_start Then
'	comStat ="준비중"
ELSE
	comStat ="준비중"
End IF

if MDPick="" then MDPick="x"
if limited="" then limited="x"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">
	// 새상품 추가 팝업
	function findProd(){
			var popwin;
			popwin = window.open("/admin/Diary2009/pop_additemlist.asp", "popup_item", "width=900,height=600,scrollbars=yes,resizable=yes");
			popwin.focus();
	}

	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

	function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,thumb){

		window.open('','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = eval("document.getElementsByName('"+iptNm+"')").value;
//		document.imginputfrm.orgImgName.value = document.getElementById(iptNm).value;
		document.imginputfrm.target='imginput';
		document.imginputfrm.action='PopImgInput.asp';
		document.imginputfrm.submit();
	}

	function jsImgDel(divnm,iptNm,vPath){

		window.open('','imgdel','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
		document.imginputfrm.divName.value=divnm;
		document.imginputfrm.inputname.value=iptNm;
		document.imginputfrm.ImagePath.value = vPath;
		document.imginputfrm.maxFileSize.value = Fsize;
		document.imginputfrm.maxFileWidth.value = Fwidth;
		document.imginputfrm.makeThumbYn.value = thumb;
		document.imginputfrm.orgImgName.value = eval("document.getElementsByName('"+iptNm+"')").value;
	//	document.imginputfrm.orgImgName.value = document.getElementById(iptNm).value;
		document.imginputfrm.target='imgdel';
		document.imginputfrm.action='PopImgInput.asp';
		document.imginputfrm.submit();
	}

	document.domain = "10x10.co.kr";

	function jsComShow(v){

		var tmp = document.getElementById("comconf");

		if (v=='Y'){
			tmp.style.display="block";
		}else {
			tmp.style.display="none";
		}
	}

	function delimage(gubun)
	{
		var aa = eval("document.frmreg."+gubun+"");
		aa.value = "";
		frmreg.submit();
	}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frmreg" method="post" action="/admin/Diary2009/Lib/DiaryRegProc.asp">
		<input type="hidden" name="mode" value="<%= Mode %>">
		<input type="hidden" name="did" value="<%= Diaryid %>">
		<input type="hidden" name="event_code" value="<%=vEventCode%>">
		<table class="tbType1 listTb">
			<tr>
				<td>
					<table class="tbType1 listTb">
						<tr bgcolor="#FFFFFF" height="25">
							<td colspan="2" ><b>다이어리 신규등록</b></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> 구분</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<% SelectList "cate",CateCode %>
							</td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 다이어리</td>
							<td bgcolor="#FFFFFF" style="text-align:left;"><%= Diaryid %></td>

						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 상품코드</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" class="text" name="iid" id="iid" value="<%=ItemID%>">
								<input type="button" class="button" value="상품찾기" onClick="findProd();">
							</td>
						</tr>
						<!--
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 기본컷 이미지 (270x270)</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv','basicimgName','basic','2000','600','false');"/>
								(<b><font color="red">270x270</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
									<input type="hidden" name="basicimgName" value="<%= ImageName %>">
									<div align="right" id="imgdiv"><% IF ImageName<>"" THEN %><img src="<%= BasicImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= BasicImgUrl %>');"><a href="javascript:delimage('basicimgName');">[삭제]</a><% End IF %></div>
							</td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 활용컷 이미지 (372x270)</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv22','basicimgName2','basic2','2000','600','false');"/>
								(<b><font color="red">372x270</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
									<input type="hidden" name="basicimgName2" value="<%= ImageName2 %>">
									<div align="right" id="imgdiv22"><% IF ImageName2<>"" THEN %><img src="<%= BasicImg2Url %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= BasicImg2Url %>');"><a href="javascript:delimage('basicimgName2');">[삭제]</a><% End IF %></div>
							</td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 큰 가로 이미지 (470x290)</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv3','storyimgName','story','2000','750','false');"/>
								(<b><font color="red">470x290</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
									<input type="hidden" name="storyimgName" value="<%= StoryImg %>">
									<div align="right" id="imgdiv3"><% IF StoryImg<>"" THEN %><img src="<%= StoryImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= StoryImgUrl %>');"><a href="javascript:delimage('storyimgName');">[삭제]</a><% End IF %></div>
							</td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 회색배경 이미지 (225x290)</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv33','basicimgName3','basic3','2000','300','false');"/>
								(<b><font color="red">225x290</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
									<input type="hidden" name="basicimgName3" value="<%= ImageName3 %>">
									<div align="right" id="imgdiv33"><% IF ImageName3<>"" THEN %><img src="<%= BasicImg3Url %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= BasicImg3Url %>');"><a href="javascript:delimage('basicimgName3');">[삭제]</a><% End IF %></div>
							</td>
						</tr>
						-->
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> MDPick 여부</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="mdpick" value="o" <% IF MDPick="o" THEN %>checked<% END IF %>>MDPick 지정
								<input type="radio" name="mdpick" value="x" <% IF MDPick="x" THEN %>checked<% END IF %> >MDPick 지정안함
							</td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> limited 여부</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="limited" value="o" <% IF limited="o" THEN %>checked<% END IF %>>limited 지정
								<input type="radio" name="limited" value="x" <% IF limited="x" THEN %>checked<% END IF %> >limited 지정안함
							</td>
						</tr>
						<!--
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap width="150"> 나눔 이미지 (270x270)<br><font color="blue">나눔이미지를 등록하면<br>오픈일도 같이 등록해야<br>리스트에 나타납니다.</font></td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv4','nanumimgName','nanum','2000','750','false');"/>
								(<b><font color="red">270x270</font></b>,<b><font color="red">JPG,GIF</font></b>만가능)
									<input type="hidden" name="nanumimgName" value="<%= nanumimg %>">
									<div align="right" id="imgdiv4"><% IF nanumimg<>"" THEN %><img src="<%= nanumimgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= nanumimgUrl %>');"><a href="javascript:delimage('nanumimgName');">[삭제]</a><% End IF %></div>
							</td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>무게</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" name="wt" value="<%= weight %>">(g)</td>
						</tr>



						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>다이어리 속지순서</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" name="soonseo" value="<%= soonseo %>" size="50"></td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>다이어리 스토리</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<textarea name="storytext" rows="6" cols="50"><%=storytext%></textarea>
							</td>
						</tr>

						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>오픈일</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" name="reservdate" value="<%= reservdate %>" size="10" maxlength="10">(예: <%=date()%> )</td>
						</tr>


						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>코맨트사용여부</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="commentyn" value="Y" <% IF commentyn="Y" THEN %>checked<% END IF %> onClick="jsComShow(this.value);">사용
								<input type="radio" name="commentyn" value="N" <% IF commentyn="N" THEN %>checked<% END IF %> onClick="jsComShow(this.value);">사용안함
							</td>
						</tr>
						//-->
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap>전시순서</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" name="mdpicksort" value="<%= mdpicksort %>" size="5"></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<td nowrap> 사용여부</td>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="radio" name="ius" value="Y" <% IF isUsing="Y" THEN %>checked<% END IF %>>사용
								<input type="radio" name="ius" value="N" <% IF isUsing="N" THEN %>checked<% END IF %> >사용안함
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF" style="text-align:left;">
					<% IF commentyn="Y" Then %>
					<table class="tbType1 listTb" id="comconf" style="display:block;">
					<% ELSE %>
					<table class="tbType1 listTb" id="comconf" style="display:none;">
					<% End IF %>
					<tr  bgcolor="<%= adminColor("tabletop") %>">
						<td nowrap width="152">코멘트 진행상태</td>
						<td bgcolor="#FFFFFF" style="text-align:left;"  ><%=comStat %></td>
					</tr>
					<tr  bgcolor="<%= adminColor("tabletop") %>">
						<td nowrap width="152">코맨트그룹코드</td>
						<td bgcolor="#FFFFFF" style="text-align:left;">
							<input type="text" name="eventgroup_code" value = "<%=eventgroup_code%>">
						</td>
					</tr>

					<tr  bgcolor="<%= adminColor("tabletop") %>">
						<td nowrap width="152">코맨트 이미지</td>
						<td bgcolor="#FFFFFF" style="text-align:left;">
							<input type="button" class="button" size="30" value="이미지 넣기" onclick="jsImgInput('imgdiv2','commentimgName','comment','2000','800','true');"/>
								<input type="hidden" name="commentimgName" value="<%= commentImgName %>">
								<div align="right" id="imgdiv2"><% IF commentImgName<>"" THEN %><img src="<%= commentImgUrl %>" width="25" height="25" style="cursor:pointer" onclick="showimage('<%= commentImgUrl %>');"><% End IF %></div>
						</td>
					</tr>
					<tr  bgcolor="<%= adminColor("tabletop") %>">
						<td nowrap width="152">기간</td>
						<td bgcolor="#FFFFFF" style="text-align:left;">
							<input type="text" name="event_start" size=10 value="<%= event_start %>">
							<a href="javascript:calendarOpen3(frmreg.event_start,'시작일',frmreg.event_start.value)">
							<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
							~<input type="text" name="event_end" size=10  value="<%= event_end %>">
							<a href="javascript:calendarOpen3(frmreg.event_end,'마지막일',frmreg.event_end.value)">
							<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="frmreg.submit();" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="frmreg.reset();" style="cursor:pointer">
					<img src="http://testwebadmin.10x10.co.kr/images/icon_new_registration.gif" border="0" onClick="location.href='/admin/diary2009/DiaryReg.asp';" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
		<form name="imginputfrm" method="post" action="">
			<input type="hidden" name="YearUse" value="2012">
			<input type="hidden" name="divName" value="">
			<input type="hidden" name="orgImgName" value="">
			<input type="hidden" name="inputname" value="">
			<input type="hidden" name="ImagePath" value="">
			<input type="hidden" name="maxFileSize" value="">
			<input type="hidden" name="maxFileWidth" value="">
			<input type="hidden" name="makeThumbYn" value="">
		</form>
	</div>
</div>
<!-- 리스트 끝 -->

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->