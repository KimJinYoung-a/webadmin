// ** I18N

// DyCalendar KO language
// Author: Mihai Bazon, <mihai_bazon@yahoo.com>
// Translation: Yourim Yi <yyi@yourim.net>
// Encoding: UTF-8
// lang : ko
// Distributed under the same terms as the calendar itself.

// For translators: please use UTF-8 if possible.  We strongly believe that
// Unicode is the answer to a real internationalized world.  Also please
// include your contact information in the header, as can be seen above.

// full day names

DyCalendar._DN = new Array
("�Ͽ���",
 "������",
 "ȭ����",
 "������",
 "�����",
 "�ݿ���",
 "�����",
 "�Ͽ���");

// Please note that the following array of short day names (and the same goes
// for short month names, _SMN) isn't absolutely necessary.  We give it here
// for exemplification on how one can customize the short day names, but if
// they are simply the first N letters of the full name you can simply say:
//
//   DyCalendar._SDN_len = N; // short day name length
//   DyCalendar._SMN_len = N; // short month name length
//
// If N = 3 then this is not needed either since we assume a value of 3 if not
// present, to be compatible with translation files that were written before
// this feature.

// short day names
DyCalendar._SDN = new Array
("��",
 "��",
 "ȭ",
 "��",
 "��",
 "��",
 "��",
 "��");

// full month names
DyCalendar._MN = new Array
("1��",
 "2��",
 "3��",
 "4��",
 "5��",
 "6��",
 "7��",
 "8��",
 "9��",
 "10��",
 "11��",
 "12��");

// short month names
DyCalendar._SMN = new Array
("1",
 "2",
 "3",
 "4",
 "5",
 "6",
 "7",
 "8",
 "9",
 "10",
 "11",
 "12");

// tooltips
DyCalendar._TT = {};
DyCalendar._TT["INFO"] = "calendar �Ұ�";

DyCalendar._TT["ABOUT"] =
"��¥ ����:\n" +
"- ������ �����Ϸ��� \xab, \xbb ��ư�� ����մϴ�\n" +
"- ���� �����Ϸ��� " + String.fromCharCode(0x2039) + ", " + String.fromCharCode(0x203a) + " ��ư�� ��������\n" +
"- ��� ������ ������ �� ������ ������ �����Ͻ� �� �ֽ��ϴ�.";
DyCalendar._TT["ABOUT_TIME"] = "\n\n" +
"�ð� ����:\n" +
"- ���콺�� ������ �ð��� �����մϴ�\n" +
"- Shift Ű�� �Բ� ������ �����մϴ�\n" +
"- ���� ���¿��� ���콺�� �����̸� �� �� ������ ���� ���մϴ�.\n";

DyCalendar._TT["PREV_YEAR"] = "���� �� (��� ������ ���)";
DyCalendar._TT["PREV_MONTH"] = "���� �� (��� ������ ���)";
DyCalendar._TT["GO_TODAY"] = "���� ��¥��";
DyCalendar._TT["NEXT_MONTH"] = "���� �� (��� ������ ���)";
DyCalendar._TT["NEXT_YEAR"] = "���� �� (��� ������ ���)";
DyCalendar._TT["SEL_DATE"] = "��¥�� �����ϼ���";
DyCalendar._TT["DRAG_TO_MOVE"] = "���콺 �巡�׷� �̵� �ϼ���";
DyCalendar._TT["PART_TODAY"] = " (����)";

DyCalendar._TT["DAY_FIRST"] = "%s ���� ǥ��";

DyCalendar._TT["WEEKEND"] = "0,6";

DyCalendar._TT["CLOSE"] = "�ݱ�";
DyCalendar._TT["TODAY"] = "����";
DyCalendar._TT["TIME_PART"] = "(Shift-)Ŭ�� �Ǵ� �巡�� �ϼ���";

// date formats
DyCalendar._TT["DEF_DATE_FORMAT"] = "%Y-%m-%d";
DyCalendar._TT["TT_DATE_FORMAT"] = "%b/%e [%a]";

DyCalendar._TT["WK"] = "��";
DyCalendar._TT["TIME"] = "��:";
