<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script>
function SubmitForm()
{

		if (document.f.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("내용을 입력하세요.");
                return;
        }

        document.f.submit();
}
</script>
<script language="JavaScript">
<!--
///////////////////////////////////////////////////////
// htmlarea 불러오기
// Author : Swoo Woong, Seol (swseol@wisenut.co.kr)
//
	_editor_url = "/editer/";                      //URL to hrmlarea files
	var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
	if (navigator.userAgent.indexOf('Mac') >= 0) { win_ie_ver = 0; }
	if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
	if (navigator.userAgent.indexOf('Opera') >= 0) { win_ie_ver = 0; }
	if (win_ie_ver >= 5.5) {
		document.write ('<scr' + 'ipt src="' + _editor_url+ 'editor.js"');
		document.write ('   language="javascript1.2"></scr'+'ipt>');
	} else {
		document.write ('<scr'+'ipt> funtion editor_generate() { return false; } </scr'+'ipt>');
	}


//////////////////////////////////////////////////////////
// htmlarea	conigure
// Author :	Swoo Woong,	Seol (swseol@wisenut.co.kr)
//
var	config = new Object();	  // create	new	config object

config.width = "95%";
config.height =	"200px";
//config.bodyStyle = 'background-color:	white; font-family:	"Verdana"; font-size: x-small;';
config.debug = 0;

// NOTE:  You can remove any of	these blocks and use the default config!

config.toolbar = [
	['fontname'],
	['fontsize'],
	['fontstyle'],
	['linebreak'],
	['bold','italic','underline','separator'],
//	['strikethrough','subscript','superscript','separator'],
	['justifyleft','justifycenter','justifyright','separator'],
	['OrderedList','UnOrderedList','Outdent','Indent','separator'],
	['forecolor','backcolor','separator'],
//	  ['HorizontalRule','Createlink','InsertImage','htmlmode','separator'],
//	['Createlink','htmlmode','separator']
//	  ['about','help','popupeditor'],
	['Createlink','separator']
];

config.fontnames = {
	"굴림":		   "굴림, 굴림체",
	"궁서":		   "궁서, 궁서체",
	"돋움":	   "돋움, 돋움체",
	"바탕":		   "바탕, 바탕체"
};
config.fontsizes = {
	"1 (8 pt)":	 "1",
	"2 (10 pt)": "2",
	"3 (12 pt)": "3",
	"4 (14 pt)": "4",
	"5 (18 pt)": "5",
	"6 (24 pt)": "6",
	"7 (36 pt)": "7"
  };

//config.stylesheet	= "http://www.domain.com/sample.css";
  
config.fontstyles =	[	// make	sure classNames	are	defined	in the page	the	content	is being display as	well in	or they	won't work!
  {	name: "headline",	  className: "headline",  classStyle: "font-family:	arial black, arial;	font-size: 28px; letter-spacing: -2px;"	},
  {	name: "arial red",	  className: "headline2", classStyle: "font-family:	arial black, arial;	font-size: 12px; letter-spacing: -2px; color:red" },
  {	name: "verdana blue", className: "headline4", classStyle: "font-family:	verdana; font-size:	18px; letter-spacing: -2px;	color:blue"	}

// leave classStyle	blank if it's defined in config.stylesheet (above),	like this:
//	{ name:	"verdana blue",	className: "headline4",	classStyle:	"" }  
];

//-->
</script>
<table border="0" cellspacing="1" bgcolor="#99a9bc" width="650" class="a">
<form method="post" name="f" action="offshop_event_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
	<tr>
		<td width="100" align="center" style="color:white">글유형</td>
		<td bgcolor="white" style="padding:0">
			<select name="gubun">
				<option value="">선택</option>
				<option value="00">전체</option>
				<option value="01">대학로점</option>
				<option value="02">잠실점</option>
			</select>
		</td>
	</tr>	
	<tr>
		<td width="100" align="center" style="color:white">제목</td>
		<td bgcolor="white" style="padding:0">
				<input name="title" style="width:450" maxlength="40" style="border:1 solid" value="">
		</td>
	</tr>
	<tr>
		<td width="100" align="center" style="color:white">내용</td>
		<td bgcolor="white" style="padding:0">
				<textarea name="contents" cols="50" rows="15"></textarea>
				<script language="javascript1.2">
					editor_generate('contents',config);
				</script>
		</td>
	</tr>
	<tr>
		<td width="100" align="center" style="color:white">종료일</td>
		<td bgcolor="white" style="padding:0">
				<input type="text" name="enddate" style="width:200" maxlength="40" style="border:1 solid" value="">(2004-07-30)
		</td>
	</tr>
	<tr>
		<td style="padding:0" colspan="2" align="right" bgcolor="white">
			<input type="button" value="Save" onclick="SubmitForm()" style="background-color:#dddddd; height:25; border:1 solid buttonface">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->