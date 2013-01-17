#!/usr/bin/perl

# Author : ping.bao.cn@gmail.com

use POSIX qw(strftime);
use Time::HiRes qw(gettimeofday);
use File::Basename;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use utf8;
use Encode;
use Encode::CN;
use strict;
use Getopt::Long;

my $cur_dir;
my $USER_XLS_PATH;
my $USER_XLS_SHEET_NAME;
my $USER_XLS_HEADER_ROW;
my $USER_XLS_DATE_RANGE;
my $DATA_XLS_AGENT_SUMMARY;
my $DATA_XLS_AGENT_SHEET_NAME;
my $DATA_XLS_HEADER_ROW;
my $DEST_XLS_SIGNATRUE;
my $DEST_XLS_IGNORE_TL;

my $AGENT_XLS_PATH;
my $AGENT_XLS_SHEET_NAME;
my $AGENT_XLS_HEADER_ROW;
my $ATT_XLS_PATH;
my $ATT_XLS_SHEET_NAME;
my $ATT_XLS_HEADER_ROW;
my $OUTBOUND_XLS_PATH;
my $OUTBOUND_XLS_SHEET_NAME;
my $OUTBOUND_XLS_HEADER_ROW;
my $CALL_XLS_PATH;
my $CALL_XLS_SHEET_NAME;
my $CALL_XLS_HEADER_ROW;

my $SHIFTS_XLS_PATH;
my $SHIFTS_XLS_SHEET_NAME;
my $SHIFTS_XLS_HEADER_ROW;
my $SHIFTS_XLS_IGNORE_SEC;

my $BLANK_XLS_PATH;
my $BLANK_XLS_SHEET_NAME;
my $BLANK_XLS_ORANGE;
my $BLANK_XLS_GREEN;
my $BLANK_XLS_PURPLE;

use constant AGENT_STATUS_LIST => ('AcdAgentNotAnswering','ACW','ACW - Manual','At a training session','At Lunch','Available','Away from desk','Call Back','Coaching','Do Not disturb','Follow Up','Gone Home','In a meeting','Out of the office','Technical Support','Working at Home','Total');
use constant ATT_STATUS_LIST => ('Agent','n Abandoned Acd','Call Ans','Avg AnsTime','Avg ACW','Avg Handle Time','NetAvg AnsTime','Avg Hold Time','Total AnsTime');
use constant CALL_STATUS_LIST => ('Agent','外拨电话量','外拨接起电话量','净通话总时长','Total Handle Time','净平均通话时长','平均Hold时长','平均 Handle 时长','外拨无人接听电话量','无人接听总时长');
use constant MON_STATUS_LIST => ('TB ID','Alias','TL','Batch','Call Answer','AHT','工作时间','Utilization','Call Back','外拨电话量','外拨接起电话量','Total Handle Time','平均 Handle 时长','无人接听总时长','外呼通话时长','外呼利用率');
use constant BLANK_STATUS_LIST => ('任务ID','个案编号','任务来源','会员昵称','任务状态','任务标题','优先级','创建人','创建日期','处理人','TL','处理时间','计划完成时间','提醒时间','业务类型','问题类型','是否有会员名');

$cur_dir = dirname($0);
chdir($cur_dir);

my $xls_type;
GetOptions ("t|type=s" => \$xls_type, "v|version|V!");

if (our $opt_v) {
	print "\nVersion : 1.0.3\n";
	exit;	
}
if ($xls_type !~ /^(agent)|(callback)|(shifts)|(blank)$/) {
	print "\nusage : -t [agent|callback|shifts|blank]";
	print "\n        -v\n";
	exit;
}
print strftime("%Y-%m-%d %H:%M:%S", localtime())," : STARTING ...\n";

if( ! -e 'config.ini') {
	print 'ERROR : FILE NOT FOUND config.ini';
	<STDIN>;
	exit;
}

open FH, 'config.ini';
while (<FH>)
{
	# 状态
	
	if (/^USER_XLS_PATH(.*)=(.*)/){
		$USER_XLS_PATH = $2;
		$USER_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^USER_XLS_SHEET_NAME(.*)=(.*)/){
		$USER_XLS_SHEET_NAME = $2;
		$USER_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^USER_XLS_HEADER_ROW(.*)=(.*)/){
		$USER_XLS_HEADER_ROW = $2;
		$USER_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^USER_XLS_DATE_RANGE(.*)=(.*)/){
		$USER_XLS_DATE_RANGE = $2;
		$USER_XLS_DATE_RANGE =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^DATA_XLS_AGENT_SUMMARY(.*)=(.*)/){
		$DATA_XLS_AGENT_SUMMARY = $2;
		$DATA_XLS_AGENT_SUMMARY =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^DATA_XLS_AGENT_SHEET_NAME(.*)=(.*)/){
		$DATA_XLS_AGENT_SHEET_NAME = $2;
		$DATA_XLS_AGENT_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^DATA_XLS_HEADER_ROW(.*)=(.*)/){
		$DATA_XLS_HEADER_ROW = $2;
		$DATA_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^DEST_XLS_SIGNATRUE(.*)=(.*)/){
		$DEST_XLS_SIGNATRUE = $2;
		$DEST_XLS_SIGNATRUE =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^DEST_XLS_IGNORE_TL(.*)=(.*)/){
		$DEST_XLS_IGNORE_TL = $2;
		$DEST_XLS_IGNORE_TL =~ s/(^\s*)|(\s*$)//g;
	}

	# 回拨
	
	if (/^AGENT_XLS_PATH(.*)=(.*)/){
		$AGENT_XLS_PATH = $2;
		$AGENT_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^AGENT_XLS_SHEET_NAME(.*)=(.*)/){
		$AGENT_XLS_SHEET_NAME = $2;
		$AGENT_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^AGENT_XLS_HEADER_ROW(.*)=(.*)/){
		$AGENT_XLS_HEADER_ROW = $2;
		$AGENT_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^ATT_XLS_PATH(.*)=(.*)/){
		$ATT_XLS_PATH = $2;
		$ATT_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^ATT_XLS_SHEET_NAME(.*)=(.*)/){
		$ATT_XLS_SHEET_NAME = $2;
		$ATT_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^ATT_XLS_HEADER_ROW(.*)=(.*)/){
		$ATT_XLS_HEADER_ROW = $2;
		$ATT_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^OUTBOUND_XLS_PATH(.*)=(.*)/){
		$OUTBOUND_XLS_PATH = $2;
		$OUTBOUND_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^OUTBOUND_XLS_SHEET_NAME(.*)=(.*)/){
		$OUTBOUND_XLS_SHEET_NAME = $2;
		$OUTBOUND_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^OUTBOUND_XLS_HEADER_ROW(.*)=(.*)/){
		$OUTBOUND_XLS_HEADER_ROW = $2;
		$OUTBOUND_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^CALL_XLS_PATH(.*)=(.*)/){
		$CALL_XLS_PATH = $2;
		$CALL_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^CALL_XLS_SHEET_NAME(.*)=(.*)/){
		$CALL_XLS_SHEET_NAME = $2;
		$CALL_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^CALL_XLS_HEADER_ROW(.*)=(.*)/){
		$CALL_XLS_HEADER_ROW = $2;
		$CALL_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^SHIFTS_XLS_PATH(.*)=(.*)/){
		$SHIFTS_XLS_PATH = $2;
		$SHIFTS_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	
	# 通宵
	
	if (/^SHIFTS_XLS_SHEET_NAME(.*)=(.*)/){
		$SHIFTS_XLS_SHEET_NAME = $2;
		$SHIFTS_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^SHIFTS_XLS_HEADER_ROW(.*)=(.*)/){
		$SHIFTS_XLS_HEADER_ROW = $2;
		$SHIFTS_XLS_HEADER_ROW =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^SHIFTS_XLS_IGNORE_SEC(.*)=(.*)/){
		$SHIFTS_XLS_IGNORE_SEC = $2;
		$SHIFTS_XLS_IGNORE_SEC =~ s/(^\s*)|(\s*$)//g;
	}
	
	# 空白会员
	
	if (/^BLANK_XLS_PATH(.*)=(.*)/){
		$BLANK_XLS_PATH = $2;
		$BLANK_XLS_PATH =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^BLANK_XLS_SHEET_NAME(.*)=(.*)/){
		$BLANK_XLS_SHEET_NAME = $2;
		$BLANK_XLS_SHEET_NAME =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^BLANK_XLS_ORANGE(.*)=(.*)/){
		$BLANK_XLS_ORANGE = $2;
		$BLANK_XLS_ORANGE =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^BLANK_XLS_GREEN(.*)=(.*)/){
		$BLANK_XLS_GREEN = $2;
		$BLANK_XLS_GREEN =~ s/(^\s*)|(\s*$)//g;
	}
	if (/^BLANK_XLS_PURPLE(.*)=(.*)/){
		$BLANK_XLS_PURPLE = $2;
		$BLANK_XLS_PURPLE =~ s/(^\s*)|(\s*$)//g;
	}
}

my $excel = CreateObject Win32::OLE 'Excel.Application' or die ('ERROR : Microsoft Excel NOT FOUND ...');
$excel -> {'EnableEvents'} = 0;

if ($xls_type eq 'agent') {
	if ( ! -e "$USER_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $USER_XLS_PATH";
		&error_found;
	}
	if ( ! -e "$DATA_XLS_AGENT_SUMMARY" ) {
		print "ERROR : FILE NOT FOUND , $DATA_XLS_AGENT_SUMMARY";
		&error_found;
	}
	&agent_report;
} elsif  ($xls_type eq 'callback') {
	if ( ! -e "$AGENT_XLS_PATH " ) {
		print "ERROR : FILE NOT FOUND , $AGENT_XLS_PATH";
		&error_found;
	}
	if ( ! -e "$ATT_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $ATT_XLS_PATH";
		&error_found;
	}
	if ( ! -e "$OUTBOUND_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $OUTBOUND_XLS_PATH";
		&error_found;
	}
	if ( ! -e "$CALL_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $CALL_XLS_PATH";
		&error_found;
	}
	&callback_report;
} elsif  ($xls_type eq 'shifts') {
	if ( ! -e "$SHIFTS_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $SHIFTS_XLS_PATH";
		&error_found;
	}
	&shifts_report;
} elsif  ($xls_type eq 'blank') {
	if ( ! -e "$BLANK_XLS_PATH" ) {
		print "ERROR : FILE NOT FOUND , $BLANK_XLS_PATH";
		&error_found;
	}
	&blank_report;
}

$excel -> Close();
$excel -> Quit();

#################################### GLOBAL AGENT SUMARRY #########################################################################

sub error_found
{
	$excel -> Close();
	$excel -> Quit();
	<STDIN>;
	exit;
}

#  parameter : $XLS_PATH  $SHEET_NAME $HEADER_ROW [01]

sub get_agent_infomation
{
	my $XLS_PATH = shift;
	my $SHEET_NAME = shift;
	my $HEADER_ROW = shift;
	my $SORT_USED = shift;
	
	my $USER_INFO;
	my $AcdAgentNotAnswering;
	my $ACW;
	my $ACW_Manual;
	my $At_a_training_session;
	my $At_Lunch;
	my $Available;
	my $Away_from_desk;
	my $Call_Back;
	my $Coaching;
	my $Do_Not_disturb;
	my $Follow_Up;
	my $Gone_Home;
	my $In_a_meeting;
	my $Out_of_the_office;
	my $Technical_Support;
	my $Working_at_Home;
	my $Total;

	my $workbook;
	my $sheet;
	
	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($XLS_PATH, 1);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($SHEET_NAME);

	# 数量
	my $used_rows = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	my $used_count = $sheet -> {'UsedRange'} ->{'Columns'} -> {'Count'};
	if ($used_count eq '') {
		print 'ERROR : Agent Summary parsed error ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}

	# 映射
	for (my $i=1; $i<=$used_count; $i++)
	{
		my $temp_val = $sheet->Cells($HEADER_ROW,$i)->{'Value'};

		if ($temp_val eq (AGENT_STATUS_LIST)[0]) {
			$AcdAgentNotAnswering = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[1]) {
			$ACW = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[2]) {
			$ACW_Manual = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[3]) {
			$At_a_training_session = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[4]) {
			$At_Lunch = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[5]) {
			$Available = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[6]) {
			$Away_from_desk = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[7]) {
			$Call_Back = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[8]) {
			$Coaching = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[9]) {
			$Do_Not_disturb = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[10]) {
			$Follow_Up = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[11]) {
			$Gone_Home = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[12]) {
			$In_a_meeting = $i;	
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[13]) {
			$Out_of_the_office = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[14]) {
			$Technical_Support = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[15]) {
			$Working_at_Home = $i;
		} elsif ($temp_val eq (AGENT_STATUS_LIST)[16]) {
			$Total = $i;
		}
	}

	for ( my $i=$HEADER_ROW+1; $i<=$used_rows; $i++)
	{
		my $TB_ID = $sheet-> Cells($i,1)->{'Value'};
		if ($TB_ID eq '') {
			next;
		}
		if ($SORT_USED eq 1)
		{
			$USER_INFO ->{$i} -> {$TB_ID} -> {'AcdAgentNotAnswering'} = $sheet-> Cells($i,$AcdAgentNotAnswering)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'ACW'} = $sheet-> Cells($i,$ACW)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'ACW_Manual'} = $sheet-> Cells($i,$ACW_Manual)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'At_a_training_session'} = $sheet-> Cells($i,$At_a_training_session)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'At_Lunch'} = $sheet-> Cells($i,$At_Lunch)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Available'} = $sheet-> Cells($i,$Available)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Away_from_desk'} = $sheet-> Cells($i,$Away_from_desk)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Call_Back'} = $sheet-> Cells($i,$Call_Back)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Coaching'} = $sheet-> Cells($i,$Coaching)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Do_Not_disturb'} = $sheet-> Cells($i,$Do_Not_disturb)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Follow_Up'} = $sheet-> Cells($i,$Follow_Up)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Gone_Home'} = $sheet-> Cells($i,$Gone_Home)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'In_a_meeting'} = $sheet-> Cells($i,$In_a_meeting)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Out_of_the_office'} = $sheet-> Cells($i,$Out_of_the_office)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Technical_Support'} = $sheet-> Cells($i,$Technical_Support)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Working_at_Home'} = $sheet-> Cells($i,$Working_at_Home)->{'Value'};
			$USER_INFO ->{$i} -> {$TB_ID} -> {'Total'} = $sheet-> Cells($i,$Total)->{'Value'};
		} else {
			$USER_INFO -> {$TB_ID} -> {'AcdAgentNotAnswering'} = $sheet-> Cells($i,$AcdAgentNotAnswering)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'ACW'} = $sheet-> Cells($i,$ACW)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'ACW_Manual'} = $sheet-> Cells($i,$ACW_Manual)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'At_a_training_session'} = $sheet-> Cells($i,$At_a_training_session)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'At_Lunch'} = $sheet-> Cells($i,$At_Lunch)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Available'} = $sheet-> Cells($i,$Available)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Away_from_desk'} = $sheet-> Cells($i,$Away_from_desk)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Call_Back'} = $sheet-> Cells($i,$Call_Back)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Coaching'} = $sheet-> Cells($i,$Coaching)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Do_Not_disturb'} = $sheet-> Cells($i,$Do_Not_disturb)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Follow_Up'} = $sheet-> Cells($i,$Follow_Up)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Gone_Home'} = $sheet-> Cells($i,$Gone_Home)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'In_a_meeting'} = $sheet-> Cells($i,$In_a_meeting)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Out_of_the_office'} = $sheet-> Cells($i,$Out_of_the_office)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Technical_Support'} = $sheet-> Cells($i,$Technical_Support)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Working_at_Home'} = $sheet-> Cells($i,$Working_at_Home)->{'Value'};
			$USER_INFO -> {$TB_ID} -> {'Total'} = $sheet-> Cells($i,$Total)->{'Value'};
		}
	}
	$workbook -> Close({SaveChanges => 0});
	return $USER_INFO;
}

#################################### AGENT #############################################

sub get_user_list
{
	my $USER;
	my $workbook;
	my $sheet;

	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($USER_XLS_PATH);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($USER_XLS_SHEET_NAME);
	
	# 获取工作表行数
	my $used_count = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	if ($used_count eq "") {
		print 'ERROR : User Not Found , Maybe The Sheet Name Is Wrong ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}
	
	# 获取日期无前辍 0
	my $current_date;
	my $current_time;
	$current_time = sprintf("%d",strftime("%H", localtime()));
	if ( $current_time eq 0 ) {
		$current_date = sprintf("%d",strftime("%d", localtime(time - 86400)));
	} else {
		$current_date = sprintf("%d",strftime("%d", localtime()));
	}
	
	# 确定日期
	my @date_range = map { s/\s*//g; $_} split(',', $USER_XLS_DATE_RANGE);
	my $current_column;
	foreach my $a (@date_range)
	{
		if ($sheet-> Range("${a}${USER_XLS_HEADER_ROW}")->{'Value'} eq $current_date) {
			$current_column = $a;
		}
	}
	if (!defined($current_column)) {
		print 'ERROR : DATE Column Not Found , Maybe The Sheet Name Is Wrong ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}
	
	for (my $i=$USER_XLS_HEADER_ROW; $i<=$used_count; $i++)
	{
		my $TB_ID = $sheet-> Range("D$i")->{'Value'};
		if ($TB_ID eq '') {
			next;
		}
		my $ALIAS = $sheet-> Range("F$i")->{'Value'};
		my $TL = $sheet-> Range("B$i")->{'Value'};
		my $BATCH = $sheet-> Range("H$i")->{'Value'};	
		my $NUM = $sheet-> Range("$current_column$i")->{'Value'};
		
		chomp $TB_ID;
		$USER -> {$TB_ID} -> {'ALIAS'} = $ALIAS;
		$USER -> {$TB_ID} -> {'TL'} = $TL;
		$USER -> {$TB_ID} -> {'BATCH'} = $BATCH;
		$USER -> {$TB_ID} -> {'NUM'} = $NUM;
	}
	$workbook -> Close({SaveChanges => 0});
	return $USER;
}

sub analytics
{
	my $ACW_NUM = shift;
	my $TMP;
	my $RESULT_TEXT;
	if ( $ACW_NUM > 1199 ) {
		$TMP = $ACW_NUM/60;
		$TMP =~ s/^([0-9]+)\..*/$1/;
		$RESULT_TEXT = $TMP.encode('gbk','分钟Acw影响，请所属TL跟进');
	}
	return $RESULT_TEXT;
}

sub agent_report
{	
	my @KEYS;
	my $VAL1;
	my $VAL2;
	my $VAL3;
	my $VAL4;
	my $err_status;
	my $usr_text;
	my $s_time;
	my $e_time;
	
	$s_time = gettimeofday();
	my $user_info = get_user_list;
	$e_time = gettimeofday();
	
	if (defined($user_info)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse user infomation successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}
	
	$s_time = gettimeofday();
	my $user_data = get_agent_infomation($DATA_XLS_AGENT_SUMMARY,$DATA_XLS_AGENT_SHEET_NAME,$DATA_XLS_HEADER_ROW,1);
	$e_time = gettimeofday();
	
	if (defined($user_data)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse agent summary successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}
	
	$excel -> {'Visible'} = 1;
	
	my $workbook_rs;
	my $sheet_rs;
	my $xls_doc_rs;
	
	# 生成
	$workbook_rs = $excel -> Workbooks -> Add();
	$sheet_rs =  $workbook_rs -> Worksheets(1);
	
	$xls_doc_rs = "$cur_dir\\Agent Summary Report.xlsx";
	eval {
		unlink $xls_doc_rs if (-e "$xls_doc_rs");
	};
	
	# 格式
	my @tmp_sheet_title = ('','AcdAgentNotAnswering','ACW','ACW - Manual','At a training session','At Lunch','Available','Away from desk','Call Back','Coaching','Do Not disturb','Gone Home','In a meeting','Out of the office','Technical Support','Working at Home','Total','Total -  gone home','Total -  gone home - Lunch time','出勤小时数','花名','TL','批次','班次','ATT','+ACW','+Available','+ACW - Manual','跟进结果','跟进人');
	@tmp_sheet_title = map {encode('gbk', $_)} @tmp_sheet_title;
	$sheet_rs -> Range('A1:AD1') -> {'Value'} = [@tmp_sheet_title];

	$sheet_rs -> Range('A1:AD1') -> {RowHeight} = 40;
	
	# 白字加粗，灰背景
	$sheet_rs -> Range('B1:P1') -> {Interior} -> {Color} = 0xb1b1b1;
	$sheet_rs -> Range('B1:P1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('B1:P1') -> {Font} -> {Bold} = 1;
	
	# 白字加粗，红背景
	$sheet_rs -> Range('Q1:S1') -> {Interior} -> {Color} = 0x0000FF;
	$sheet_rs -> Range('Q1:S1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('Q1:S1') -> {Font} -> {Bold} = 1;

	# 黑字加粗，灰背景
	$sheet_rs -> Range('T1:X1') -> {Interior} -> {Color} = 0xb1b1b1;
	$sheet_rs -> Range('T1:X1') -> {Font} -> {Color} = 0x000000;
	$sheet_rs -> Range('T1:X1') -> {Font} -> {Bold} = 1;
	
	# 白字加粗，红背景
	$sheet_rs -> Range('Y1:AD1') -> {Interior} -> {Color} = 0x0000FF;
	$sheet_rs -> Range('Y1:AD1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('Y1:AD1') -> {Font} -> {Bold} = 1;
	
	$sheet_rs -> Range('A1:AD1') -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range('A1:AD1') -> {Borders} -> {Color} = 0x000000;
	
	# 数据格式
	$sheet_rs -> Columns('T') -> {'NumberFormatLocal'} = "0.00";
	$sheet_rs -> Columns('Y') -> {'NumberFormatLocal'} = "0.00%";
	$sheet_rs -> Columns('Z') -> {'NumberFormatLocal'} = "0.00%";
	$sheet_rs -> Columns('AA') -> {'NumberFormatLocal'} = "0.00%";
	$sheet_rs -> Columns('AB') -> {'NumberFormatLocal'} = "0.00%";
		
	# 处理
	my @ignore_list = map { s/\s*//g; $_} split(',', $DEST_XLS_IGNORE_TL);
	my $k;
	my $v;
	my $tag = 2;
	@KEYS = sort { $a <=> $b } keys($user_data);
	foreach my $key (@KEYS){
		while (($k, $v)=each $user_data->{$key})
		{
			# 处理 TL IGNORE 列表
			if ( grep { $_ eq $user_info->{$k}->{'TL'} } @ignore_list ){
				print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Ignore $k ...\n";
				next;
			}
			if (! defined($user_info->{$k}->{"TL"})){
				print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Ignore TL $k ...\n";
				next;
			}
			
			$err_status = 0;
			# print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Filling $k ...\n";
			
			my $ATT_ADD = $v->{'Total'} - $v->{'Gone_Home'} - $v->{'At_Lunch'};
			my $ACW_ADD = $v->{'ACW'} + $v->{'Do_Not_disturb'};
			
			# 处理 ATT IGNORE 情况
			if ($ATT_ADD eq 0) {
				print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Ignore ATT $k ...\n";
				next;	
			} 

			$sheet_rs -> Range("A$tag")->{'Value'} = $k;
			$sheet_rs -> Range("B$tag")->{'Value'} = $v->{'AcdAgentNotAnswering'};
			$sheet_rs -> Range("C$tag")->{'Value'} = $v->{'ACW'};
			$sheet_rs -> Range("D$tag")->{'Value'} = $v->{'ACW_Manual'};
			$sheet_rs -> Range("E$tag")->{'Value'} = $v->{'At_a_training_session'};
			$sheet_rs -> Range("F$tag")->{'Value'} = $v->{'At_Lunch'};
			$sheet_rs -> Range("G$tag")->{'Value'} = $v->{'Available'};
			$sheet_rs -> Range("H$tag")->{'Value'} = $v->{'Away_from_desk'};
			$sheet_rs -> Range("I$tag")->{'Value'} = $v->{'Call_Back'};
			$sheet_rs -> Range("J$tag")->{'Value'} = $v->{'Coaching'};
			$sheet_rs -> Range("K$tag")->{'Value'} = $v->{'Do_Not_disturb'};
			$sheet_rs -> Range("L$tag")->{'Value'} = $v->{'Gone_Home'};
			$sheet_rs -> Range("M$tag")->{'Value'} = $v->{'In_a_meeting'};
			$sheet_rs -> Range("N$tag")->{'Value'} = $v->{'Out_of_the_office'};
			$sheet_rs -> Range("O$tag")->{'Value'} = $v->{'Technical_Support'};
			$sheet_rs -> Range("P$tag")->{'Value'} = $v->{'Working_at_Home'};
			$sheet_rs -> Range("Q$tag")->{'Value'} = $v->{'Total'};
			$sheet_rs -> Range("R$tag")->{'Value'} = $v->{'Total'} - $v->{'Gone_Home'};
			$sheet_rs -> Range("S$tag")->{'Value'} = $ATT_ADD;
			$sheet_rs -> Range("T$tag")->{'Value'} = $ATT_ADD / 3600;
			
			if (defined($user_info->{$k}->{'ALIAS'})){
				$sheet_rs -> Range("U$tag")->{'Value'} = $user_info->{$k}->{'ALIAS'};
			} else {
				$sheet_rs -> Range("U$tag")->{'Value'} = 'N/A';
			}
			if (defined($user_info->{$k}->{'TL'})){
				$sheet_rs -> Range("V$tag")->{'Value'} = $user_info->{$k}->{'TL'};
			} else {
				$sheet_rs -> Range("V$tag")->{'Value'} = 'N/A';
			}
			if (defined($user_info->{$k}->{'BATCH'})){
				$sheet_rs -> Range("W$tag")->{'Value'} = $user_info->{$k}->{'BATCH'};
			} else {
				$sheet_rs -> Range("W$tag")->{'Value'} = 'N/A';
			}
			if (defined($user_info->{$k}->{'NUM'})){
				$sheet_rs -> Range("X$tag")->{'Value'} = $user_info->{$k}->{'NUM'};
			} else {
				$sheet_rs -> Range("X$tag")->{'Value'} = 'N/A';
			}
			if ($ATT_ADD eq 0) {
				$VAL1 = $VAL2 = $VAL3 = $VAL4 = 0;
				$sheet_rs -> Range("Y$tag:AA$tag") -> {Interior} -> {Color} = 0xcc99ff;
				$sheet_rs -> Range("Y$tag")->{'Value'} = '#DIV/0!';
				$sheet_rs -> Range("Z$tag")->{'Value'} = '#DIV/0!';
				$sheet_rs -> Range("AA$tag")->{'Value'} = '#DIV/0!';
				$sheet_rs -> Range("AB$tag")->{'Value'} = '#DIV/0!';
			} else {
			
				$VAL1 = $v->{'Do_Not_disturb'} / $ATT_ADD;
				$VAL2 = $ACW_ADD / $ATT_ADD;
				$VAL3 = ($v->{'ACW'} + $v->{'Available'} + $v->{'Do_Not_disturb'}) / $ATT_ADD;
				$VAL4 = ($v->{'ACW'} + $v->{'ACW_Manual'} + $v->{'Available'} + $v->{'Do_Not_disturb'}) / $ATT_ADD;
				
				$sheet_rs -> Range("Y$tag")->{'Value'} = $VAL1;
				$sheet_rs -> Range("Z$tag")->{'Value'} = $VAL2;
				$sheet_rs -> Range("AA$tag")->{'Value'} = $VAL3;
				$sheet_rs -> Range("AB$tag")->{'Value'} = $VAL4;
				
				if ($VAL1 < 0.8) {
					$err_status = 1;
					$sheet_rs -> Range("Y$tag") -> {Interior} -> {Color} = 0xcc99ff;
				}
				if ($VAL2 < 0.82) {
					$err_status = 1;
					$sheet_rs -> Range("Z$tag") -> {Interior} -> {Color} = 0xcc99ff;
				}
				if ($VAL3 < 0.85) {
					$err_status = 1;
					$sheet_rs -> Range("AA$tag") -> {Interior} -> {Color} = 0xcc99ff;
				}
				if ($VAL4 < 0.85) {
					$err_status = 1;
					$sheet_rs -> Range("AB$tag") -> {Interior} -> {Color} = 0xcc99ff;
				}
			}
			
			if ( $err_status eq 1 )
			{
				$sheet_rs -> Range("AD$tag")->{'Value'} = [[$DEST_XLS_SIGNATRUE]];
				$usr_text = analytics($v->{'ACW'});
				if ($usr_text ne '') {
					if ( $v->{'ACW'} > 1199 )
					{
						$sheet_rs -> Range("AC$tag") -> {Interior} -> {Color} =  0x00FFFF;
						$sheet_rs -> Range("AC$tag") -> {Font} -> {Color} = 0x0000FF;
					}
				} else {
					my @vals = ($v->{'AcdAgentNotAnswering'}, $v->{'ACW'}, $v->{'ACW_Manual'}, $v->{'At_a_training_session'}, $v->{'Available'}, $v->{'Away_from_desk'}, $v->{'Call_Back'}, $v->{'Coaching'}, $v->{'Out_of_the_office'}, $v->{'Technical_Support'}, $v->{'Working_at_Home'});
					my @val = sort { $a <=> $b } @vals;
					if ($v->{'AcdAgentNotAnswering'} eq $val[-1]) {
						$usr_text = encode('gbk','ACD 影响');
					}
					if ($v->{'ACW'} eq $val[-1]) {
						$usr_text = encode('gbk','ACW 影响');
					}
					if ($v->{'ACW_Manual'} eq $val[-1]) {
						$usr_text = encode('gbk','ACW - Manual 影响');
					}
					if ($v->{'At_a_training_session'} eq $val[-1]) {
						$usr_text = encode('gbk','At a training session 影响');
					}
					if ($v->{'Available'} eq $val[-1]) {
						$usr_text = encode('gbk','Available 影响');
					}
					if ($v->{'Away_from_desk'} eq $val[-1]) {
						$usr_text = encode('gbk','Away 影响');
					}
					if ($v->{'Call_Back'} eq $val[-1]) {
						$usr_text = encode('gbk','Call Back 影响');
					}
					if ($v->{'Coaching'} eq $val[-1]) {
						$usr_text = encode('gbk','Coaching 影响');
					}
					if ($v->{'Out_of_the_office'} eq $val[-1]) {
						$usr_text = encode('gbk','Office 影响');
					}
					if ($v->{'Technical_Support'} eq $val[-1]) {
						$usr_text = encode('gbk','Technical Support 影响');
					}
					if ($v->{'Working_at_Home'} eq $val[-1]) {
						$usr_text = encode('gbk','Working at Home 影响，请所属TL跟进');
						$sheet_rs -> Range("AC$tag") -> {Interior} -> {Color} =  0x00FFFF;
						$sheet_rs -> Range("AC$tag") -> {Font} -> {Color} = 0x0000FF;
					}
				}
			} else {
				$usr_text = '';
			}
			$sheet_rs -> Range("AC$tag")->{'Value'} = [[$usr_text]];
	
			# STYLE			
			$sheet_rs -> Range("A$tag") -> {Interior} -> {Color} = 0xb2b2b2;
			$sheet_rs -> Range("A$tag:AD$tag") -> {RowHeight} = 20;
			$sheet_rs -> Range("A$tag:AD$tag") -> {Borders} -> {LineStyle} = 1;
			$sheet_rs -> Range("A$tag:AD$tag") -> {Borders} -> {Color} = 0x000000;
			$tag++;
		}
	}
	
	# 样式
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> {HorizontalAlignment} = xlHAlignCenter;
	$sheet_rs -> Columns -> {VerticalAlignment} = xlCenter;
	$sheet_rs -> Rows(1) -> {WrapText} = 1;
	
	# 单元格宽度
	$sheet_rs -> Columns -> {ColumnWidth} = 7;
	$sheet_rs -> Columns('AC') -> {ColumnWidth} = 25;
	# $sheet_rs -> Columns -> AutoFit();
	
	# 隐藏
	# $sheet_rs -> Columns("B:M") -> {ColumnWidth} = '0';

	$sheet_rs -> SaveAs ($xls_doc_rs);
	print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Save Result To $xls_doc_rs";
}

#################################### CALL BACK ###########################################

sub get_sample_user_list
{
	my $workbook;
	my $sheet;
	
	my $USER_INFO;
	my $USER_ID;
	
	my $TB_ID;
	my $ALIAS;
	my $TL;
	my $BATCH;

	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($CALL_XLS_PATH);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($CALL_XLS_SHEET_NAME);
	
	# 行数列数
	my $used_rows = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	my $used_count = $sheet -> {'UsedRange'} ->{'Columns'} -> {'Count'};
	
	if ($used_rows eq '') {
		print 'ERROR : User Not Found , Maybe The Sheet Name Is Wrong ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}
	
	# 映射关系
	for (my $i=1; $i<=$used_count; $i++)
	{
		my $temp_val = $sheet->Cells($CALL_XLS_HEADER_ROW,$i)->{'Value'};
		
		if ($temp_val eq (MON_STATUS_LIST)[0]) {
			$TB_ID = $i;
		} elsif ($temp_val eq (MON_STATUS_LIST)[1]) {
			$ALIAS = $i;
		} elsif ($temp_val eq (MON_STATUS_LIST)[2]) {
			$TL = $i;
		} elsif ($temp_val eq (MON_STATUS_LIST)[3]) {
			$BATCH = $i;
		}
	}
	
	# 用户数据
	for ( my $i=$CALL_XLS_HEADER_ROW+1; $i<=$used_rows; $i++)
	{
		$USER_ID = $sheet-> Cells($i,$TB_ID)->{'Value'};
		if ($USER_ID eq '') {
			next;
		}
		$USER_INFO ->{$i} -> {$USER_ID} -> {'ALIAS'} = $sheet-> Cells($i,$ALIAS)->{'Value'};
		$USER_INFO ->{$i} -> {$USER_ID} -> {'TL'} = $sheet-> Cells($i,$TL)->{'Value'};
		$USER_INFO ->{$i} -> {$USER_ID} -> {'BATCH'} = $sheet-> Cells($i,$BATCH)->{'Value'};
	}
	$workbook -> Close({SaveChanges => 0});
	return $USER_INFO;
}

sub get_att_info
{
	my $workbook;
	my $sheet;

	my $ATT_INFO;
	my $USER_ID;
	
	my $TB_ID;
	my $CALL_ANS;
	my $AVG_HANDLE_TIME;

	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($ATT_XLS_PATH);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($ATT_XLS_SHEET_NAME);
		
	# 行数列数
	my $used_rows = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	my $used_count = $sheet -> {'UsedRange'} ->{'Columns'} -> {'Count'};
	
	if ($used_rows eq '') {
		print 'ERROR : User Not Found , Maybe The Sheet Name Is Wrong ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}
	
	# 映射关系
	for (my $i=1; $i<=$used_count; $i++)
	{
		my $temp_val = $sheet->Cells($ATT_XLS_HEADER_ROW,$i)->{'Value'};
		
		if ($temp_val eq (ATT_STATUS_LIST)[0]) {
			$TB_ID = $i;
		} elsif ($temp_val eq (ATT_STATUS_LIST)[2]) {
			$CALL_ANS = $i;
		} elsif ($temp_val eq (ATT_STATUS_LIST)[5]) {
			$AVG_HANDLE_TIME = $i;
		}
	}
	
	# 用户数据
	for ( my $i=$ATT_XLS_HEADER_ROW+1; $i<=$used_rows; $i++)
	{
		$USER_ID = $sheet-> Cells($i,$TB_ID)->{'Value'};
		if ($USER_ID eq '') {
			next;
		}
		$ATT_INFO -> {$USER_ID} -> {'CALL_ANS'} = $sheet-> Cells($i,$CALL_ANS)->{'Value'};
		$ATT_INFO -> {$USER_ID} -> {'AVG_HANDLE_TIME'} = $sheet-> Cells($i,$AVG_HANDLE_TIME)->{'Value'};
	}
	$workbook -> Close({SaveChanges => 0});
	return $ATT_INFO;
}

sub get_call_info
{
	my $workbook;
	my $sheet;

	my $CALL_INFO;
	my $USER_ID;
	
	my $TB_ID;
	my $CALL_OUT;
	my $CALL_NUM;
	my $TOTAL_HANDLE_TIME;
	my $AVG_HANDLE_TIME;
	my $NOANSWER_TIME;

	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($OUTBOUND_XLS_PATH);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($OUTBOUND_XLS_SHEET_NAME);
	
	# 行数列数
	my $used_rows = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	my $used_count = $sheet -> {'UsedRange'} ->{'Columns'} -> {'Count'};
	
	if ($used_rows eq '') {
		print 'ERROR : User Not Found , Maybe The Sheet Name Is Wrong ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}
	
	# 映射关系
	for (my $i=1; $i<=$used_count; $i++)
	{
		my $temp_val = decode('gbk',$sheet->Cells($OUTBOUND_XLS_HEADER_ROW,$i)->{'Value'});
		Encode::_utf8_on($temp_val);
		if ($temp_val eq (CALL_STATUS_LIST)[0]) {
			$TB_ID = $i;
		} elsif ($temp_val eq (CALL_STATUS_LIST)[1]) {
			$CALL_OUT = $i;
		} elsif ($temp_val eq (CALL_STATUS_LIST)[2]) {
			$CALL_NUM = $i;
		} elsif ($temp_val eq (CALL_STATUS_LIST)[4]) {
			$TOTAL_HANDLE_TIME = $i;
		} elsif ($temp_val eq (CALL_STATUS_LIST)[7]) {
			$AVG_HANDLE_TIME = $i;
		} elsif ($temp_val eq (CALL_STATUS_LIST)[9]) {
			$NOANSWER_TIME = $i;
		}
	}
	
	# 用户数据
	for ( my $i=$OUTBOUND_XLS_HEADER_ROW+1; $i<=$used_rows; $i++)
	{
		$USER_ID = $sheet-> Cells($i,$TB_ID)->{'Value'};
		if ($USER_ID eq '') {
			next;
		}
		$CALL_INFO -> {$USER_ID} -> {'CALL_OUT'} = $sheet-> Cells($i,$CALL_OUT)->{'Value'};
		$CALL_INFO -> {$USER_ID} -> {'CALL_NUM'} = $sheet-> Cells($i,$CALL_NUM)->{'Value'};
		$CALL_INFO -> {$USER_ID} -> {'TOTAL_HANDLE_TIME'} = $sheet-> Cells($i,$TOTAL_HANDLE_TIME)->{'Value'};
		$CALL_INFO -> {$USER_ID} -> {'AVG_HANDLE_TIME'} = $sheet-> Cells($i,$AVG_HANDLE_TIME)->{'Value'};
		$CALL_INFO -> {$USER_ID} -> {'NOANSWER_TIME'} = $sheet-> Cells($i,$NOANSWER_TIME)->{'Value'};
	}
	$workbook -> Close({SaveChanges => 0});
	return $CALL_INFO;
}

sub callback_report
{
	my $s_time;
	my $e_time;
	
	$s_time = gettimeofday();
	my $user_data = get_sample_user_list();
	$e_time = gettimeofday();
	
	if (defined($user_data)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse user infomation successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	$s_time = gettimeofday();
	my $user_att = get_att_info();
	$e_time = gettimeofday();
	
	if (defined($user_att)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse ATT successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	$s_time = gettimeofday();
	my $user_agent = get_agent_infomation($AGENT_XLS_PATH,$AGENT_XLS_SHEET_NAME,$AGENT_XLS_HEADER_ROW,0);
	$e_time = gettimeofday();
		
	if (defined($user_agent)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse agent summary successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	$s_time = gettimeofday();
	my $user_call = get_call_info();
	$e_time = gettimeofday();
	
	if (defined($user_call)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse outbound call successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	# 处理
	my $workbook_rs;
	my $sheet_rs;
	my $xls_doc_rs;
	my $tag = 2;

	# 生成
	$excel -> {'Visible'} = 1;
	$workbook_rs = $excel -> Workbooks -> Add();
	$sheet_rs =  $workbook_rs -> Worksheets(1);

	$xls_doc_rs = "$cur_dir\\Call Back Report.xlsx";
	eval {
		unlink $xls_doc_rs if (-e "$xls_doc_rs");
	};
	my @tmp_sheet_title = map {encode('gbk',$_)} (MON_STATUS_LIST);
	$sheet_rs -> Range('A1:P1') -> {'Value'} = [@tmp_sheet_title];

	# 格式
	$sheet_rs -> Range('A1:P1') -> {'RowHeight'} = 40;

	# 边框
	$sheet_rs -> Range('A1:P1') -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range('A1:P1') -> {Borders} -> {Color} = 0x000000;

	# 样式
	$sheet_rs -> Range('A1:P1') -> {Interior} -> {Color} = 0x800080;
	$sheet_rs -> Range('A1:P1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('A1:P1') -> {Font} -> {Bold} = 1;
		
	$sheet_rs -> Range('E1:H1') -> {Interior} -> {Color} = 0x669933;
	$sheet_rs -> Range('E1:H1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('E1:H1') -> {Font} -> {Bold} = 1;

	$sheet_rs -> Range('I1:P1') -> {Interior} -> {Color} = 0x996666;
	$sheet_rs -> Range('I1:P1') -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range('I1:P1') -> {Font} -> {Bold} = 1;

	# 注释
	$sheet_rs -> Range('G1')->AddComment(encode('gbk','TaoBao Agent Summary 工作时间 = Total - Gone Home - At Lunch'));

	# 换行
	$sheet_rs -> Rows(1) -> {WrapText} = 1;

	# 单元格
	$sheet_rs -> Columns -> {ColumnWidth} = 9;

	# 数据
	$sheet_rs -> Columns('H') -> {'NumberFormatLocal'} = "0.00%";
	$sheet_rs -> Columns('P') -> {'NumberFormatLocal'} = "0.00%";

	# 文本
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;

	my @KEYS = sort { $a <=> $b } keys($user_data);
	foreach my $key (@KEYS){
		my @DATA;
		while ((my $k, my $v) = each $user_data->{$key})
		{
			if ($k eq 'Total')
			{
				next;
			}
			my $A_VAL = $k;
			my $B_VAL = $v->{'ALIAS'};
			my $C_VAL = $v->{'TL'};
			my $D_VAL = $v->{'BATCH'};
		
			my $E_VAL = $user_att->{$k}->{'CALL_ANS'};
			my $F_VAL = $user_att->{$k}->{'AVG_HANDLE_TIME'};
			my $G_VAL = $user_agent->{$k}->{'Total'} - $user_agent->{$k}->{'Gone_Home'} - $user_agent->{$k}->{'At_Lunch'};
				
			my $H_VAL;
			if ($E_VAL eq '' or $G_VAL eq 0) {
				$H_VAL = '';
			} else {
				$H_VAL = $E_VAL * $F_VAL / $G_VAL;
			}
			my $I_VAL = $user_agent->{$k}->{'Call_Back'};
			my $J_VAL = $user_call->{$k}->{'CALL_OUT'};
			my $K_VAL = $user_call->{$k}->{'CALL_NUM'};
			my $L_VAL = $user_call->{$k}->{'TOTAL_HANDLE_TIME'};
			my $M_VAL = $user_call->{$k}->{'AVG_HANDLE_TIME'};
			my $N_VAL = $user_call->{$k}->{'NOANSWER_TIME'};
			my $O_VAL = $L_VAL + $N_VAL;
			my $P_VAL;
			if ($I_VAL eq '') {
				$P_VAL = '#DIV/0!';
			} else {
				$P_VAL = $O_VAL / $I_VAL;
			}
			push(@DATA,$A_VAL);
			push(@DATA,$B_VAL);
			push(@DATA,$C_VAL);
			push(@DATA,$D_VAL);
			push(@DATA,$E_VAL);
			push(@DATA,$F_VAL);
			push(@DATA,$G_VAL);
			push(@DATA,$H_VAL);
			push(@DATA,$I_VAL);
			push(@DATA,$J_VAL);
			push(@DATA,$K_VAL);
			push(@DATA,$L_VAL);
			push(@DATA,$M_VAL);
			push(@DATA,$N_VAL);
			push(@DATA,$O_VAL);
			push(@DATA,$P_VAL);
			
			# print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Filling $k ...\n";
			
			# 填充
			$sheet_rs -> Range("A$tag:P$tag")->{'Value'} =[[@DATA]];
			
			# 样式
			$sheet_rs -> Range("A$tag:D$tag") -> {Interior} -> {Color} = 0x00ffff;
			$sheet_rs -> Range("H$tag") -> {Interior} -> {Color} = 0x00ffff;
			$sheet_rs -> Range("O$tag:P$tag") -> {Interior} -> {Color} = 0x00ffff;

			# 边框
			$sheet_rs -> Range("A$tag:P$tag") -> {Borders} -> {LineStyle} = 1;
			$sheet_rs -> Range("A$tag:P$tag") -> {Borders} -> {Color} = 0x000000;
		}
		$tag++;
	}
	my $line1 = $tag-1;
	my $line2 = $tag-2;
	$sheet_rs -> Range("A$line1")->{'Value'} =[['Total']];
	$sheet_rs -> Range("E$line1")->{'Value'} = "=SUM(E2:E$line2)";
	$sheet_rs -> Range("F$line1")->{'Value'} = "=SUMPRODUCT(E2:E$line2,F2:F$line2)/E$line1";
	$sheet_rs -> Range("G$line1")->{'Value'} = "=SUM(G2:G$line2)";
	$sheet_rs -> Range("H$line1")->{'Value'} = "=AVERAGE(H2:H$line2)";
	$sheet_rs -> Range("I$line1")->{'Value'} = "=SUM(I2:I$line2)";
	$sheet_rs -> Range("J$line1")->{'Value'} = "=SUM(J2:J$line2)";
	$sheet_rs -> Range("K$line1")->{'Value'} = "=SUM(K2:K$line2)";
	$sheet_rs -> Range("L$line1")->{'Value'} = "=SUM(L2:L$line2)";
	$sheet_rs -> Range("M$line1")->{'Value'} = "=AVERAGE(M2:M$line2)";
	$sheet_rs -> Range("N$line1")->{'Value'} = "=SUM(N2:N$line2)";
	$sheet_rs -> Range("O$line1")->{'Value'} = "=SUM(O2:O$line2)";
	$sheet_rs -> Range("P$line1")->{'Value'} = "=O$line1/I$line1";

	# 全局
	$sheet_rs -> Columns -> {Font} -> {Name} = '微软雅黑';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> {HorizontalAlignment} = xlHAlignCenter;
	$sheet_rs -> Columns -> {VerticalAlignment} = xlCenter;

	# 样式
	$sheet_rs -> Range("A$line1:P$line1") -> {Interior} -> {Color} = 0x00ffff;
	$sheet_rs -> Range("A$line1:O$line1") -> {Font} -> {Color} = 0x0000FF;
	$sheet_rs -> Range("A$line1:O$line1") -> {Font} -> {Bold} = 1;

	# 格式
	$sheet_rs -> Range("F$line1") -> {'NumberFormatLocal'} = "0.00";
	$sheet_rs -> Range("M$line1") -> {'NumberFormatLocal'} = "0.00";

	# 边框
	$sheet_rs -> Range("A$line1:D$line1") -> Merge();
	$sheet_rs -> Range("A$line1:D$line1") -> {HorizontalAlignment} = xlLeft;
	$sheet_rs -> Range("A$line1:D$line1") -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range("A$line1:D$line1") -> {Borders} -> {Color} = 0x000000;

	$sheet_rs -> Range("E$line1:P$line1") -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range("E$line1:P$line1") -> {Borders} -> {Color} = 0x000000;

	$sheet_rs -> SaveAs ($xls_doc_rs);
	print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Save Result To $xls_doc_rs";
}


#################################### SHIFTS ###########################################

sub shifts_report
{
	my @KEYS;
	my $VAL1;
	my $VAL2;
	my $VAL3;
	my $VAL4;
	my $err_status;
	my $usr_text;
	
	my $s_time;
	my $e_time;
	
	$s_time = gettimeofday();
	my $user_data = get_agent_infomation($SHIFTS_XLS_PATH,$SHIFTS_XLS_SHEET_NAME,$SHIFTS_XLS_HEADER_ROW,1);
	$e_time = gettimeofday();

	if (defined($user_data)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse agent summary successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	$excel -> {'Visible'} = 1;

	my $workbook_rs;
	my $sheet_rs;
	my $xls_doc_rs;

	# 生成
	$workbook_rs = $excel -> Workbooks -> Add();
	$sheet_rs =  $workbook_rs -> Worksheets(1);
	$xls_doc_rs = "$cur_dir\\Agent Summary Shifts.xlsx";
	eval {
		unlink $xls_doc_rs if (-e "$xls_doc_rs");
	};

	# 格式
	my @sheet_title = ('','AcdAgentNotAnswering','ACW','ACW - Manual','At a training session','At Lunch','Available','Away from desk','Call Back','Coaching','Do Not disturb','Follow Up','Gone Home','In a meeting','Out of the office','Technical Support','Working at Home','Total','workingtime','利用率','（+ACW）利用率');
	@sheet_title = map {encode('gbk', $_)} @sheet_title;
	$sheet_rs -> Range('A1:U1') -> {'Value'} = [[@sheet_title]];

	# 白字加粗，灰背景
	$sheet_rs -> Range("B1:Q1") -> {Interior} -> {Color} = 0xb1b1b1;
	$sheet_rs -> Range("B1:Q1") -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range("B1:Q1") -> {Font} -> {Bold} = 1;

	# 白字加粗，红背景
	$sheet_rs -> Range("R1:U1") -> {Interior} -> {Color} = 0x0000FF;
	$sheet_rs -> Range("R1:U1") -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range("R1:U1") -> {Font} -> {Bold} = 1;

	# 边框
	$sheet_rs -> Range('A1:U1') -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range('A1:U1') -> {Borders} -> {Color} = 0x000000;

	# 数据格式
	$sheet_rs -> Columns('T') -> {'NumberFormatLocal'} = "0.00%";
	$sheet_rs -> Columns('U') -> {'NumberFormatLocal'} = "0.00%";

	# 处理
	my $k;
	my $v;
	my $tag = 2;
	@KEYS = sort { $a <=> $b } keys($user_data);
	foreach my $key (@KEYS){
		while (($k, $v)=each $user_data->{$key})
		{
			if (($k eq 'Total') or ($k =~ /TB9/)) {
				print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Ignore $k ...\n";
				next;
			}
			$err_status = 0;
			# print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Filling $k ...\n";

			my $TOTAL_ADD = $v->{'AcdAgentNotAnswering'} + $v->{'ACW'} + $v->{'ACW_Manual'} + $v->{'Available'} + $v->{'Away_from_desk'} + $v->{'Do_Not_disturb'};

			if ($TOTAL_ADD < $SHIFTS_XLS_IGNORE_SEC){
				next;
			}
			$sheet_rs -> Range("A$tag")->{'Value'} = $k;
			$sheet_rs -> Range("B$tag")->{'Value'} = $v->{'AcdAgentNotAnswering'};
			$sheet_rs -> Range("C$tag")->{'Value'} = $v->{'ACW'};
			$sheet_rs -> Range("D$tag")->{'Value'} = $v->{'ACW_Manual'};
			$sheet_rs -> Range("E$tag")->{'Value'} = $v->{'At_a_training_session'};
			$sheet_rs -> Range("F$tag")->{'Value'} = $v->{'At_Lunch'};
			$sheet_rs -> Range("G$tag")->{'Value'} = $v->{'Available'};
			$sheet_rs -> Range("H$tag")->{'Value'} = $v->{'Away_from_desk'};
			$sheet_rs -> Range("I$tag")->{'Value'} = $v->{'Call_Back'};
			$sheet_rs -> Range("J$tag")->{'Value'} = $v->{'Coaching'};
			$sheet_rs -> Range("K$tag")->{'Value'} = $v->{'Do_Not_disturb'};
			$sheet_rs -> Range("L$tag")->{'Value'} = $v->{'Follow_Up'};
			$sheet_rs -> Range("M$tag")->{'Value'} = $v->{'Gone_Home'};
			$sheet_rs -> Range("N$tag")->{'Value'} = $v->{'In_a_meeting'};
			$sheet_rs -> Range("O$tag")->{'Value'} = $v->{'Out_of_the_office'};
			$sheet_rs -> Range("P$tag")->{'Value'} = $v->{'Technical_Support'};
			$sheet_rs -> Range("Q$tag")->{'Value'} = $v->{'Working_at_Home'};
			$sheet_rs -> Range("R$tag")->{'Value'} = $v->{'Total'};
			$sheet_rs -> Range("S$tag")->{'Value'} = $TOTAL_ADD;
			$sheet_rs -> Range("T$tag")->{'Value'} = $v->{'Do_Not_disturb'} / $TOTAL_ADD;
			$sheet_rs -> Range("U$tag")->{'Value'} = ($v->{'ACW'} + $v->{'Do_Not_disturb'}) / $TOTAL_ADD;

			# STYLE
			$sheet_rs -> Range("A$tag") -> {Interior} -> {Color} = 0xb2b2b2;
			$sheet_rs -> Range("A$tag") -> {Font} -> {Color} = 0xFFFFFF;
			$sheet_rs -> Range("A$tag") -> {Font} -> {Bold} = 1;
	
			$sheet_rs -> Range("A$tag:U$tag") -> {Borders} -> {LineStyle} = 1;
			$sheet_rs -> Range("A$tag:U$tag") -> {Borders} -> {Color} = 0x000000;
			$tag++;
		}
	}
	
	# 附加
	my $line1 = $tag;
	my $line2 = $tag-1;
	$sheet_rs -> Range("A$line1")->{'Value'} =[['Total']];
	$sheet_rs -> Range("B$line1")->{'Value'} = "=SUM(B2:B$line2)";
	$sheet_rs -> Range("C$line1")->{'Value'} = "=SUM(C2:C$line2)";
	$sheet_rs -> Range("D$line1")->{'Value'} = "=SUM(D2:D$line2)";
	$sheet_rs -> Range("E$line1")->{'Value'} = "=SUM(E2:E$line2)";
	$sheet_rs -> Range("F$line1")->{'Value'} = "=SUM(F2:F$line2)";
	$sheet_rs -> Range("G$line1")->{'Value'} = "=SUM(G2:G$line2)";
	$sheet_rs -> Range("H$line1")->{'Value'} = "=SUM(H2:H$line2)";
	$sheet_rs -> Range("I$line1")->{'Value'} = "=SUM(I2:I$line2)";
	$sheet_rs -> Range("J$line1")->{'Value'} = "=SUM(J2:J$line2)";
	$sheet_rs -> Range("K$line1")->{'Value'} = "=SUM(K2:K$line2)";
	$sheet_rs -> Range("L$line1")->{'Value'} = "=SUM(L2:L$line2)";
	$sheet_rs -> Range("M$line1")->{'Value'} = "=SUM(M2:M$line2)";
	$sheet_rs -> Range("N$line1")->{'Value'} = "=SUM(N2:N$line2)";
	$sheet_rs -> Range("O$line1")->{'Value'} = "=SUM(O2:O$line2)";
	$sheet_rs -> Range("P$line1")->{'Value'} = "=SUM(P2:P$line2)";
	$sheet_rs -> Range("Q$line1")->{'Value'} = "=SUM(Q2:Q$line2)";
	$sheet_rs -> Range("R$line1")->{'Value'} = "=SUM(R2:R$line2)";
	
	$sheet_rs -> Range("S$line1")->{'Value'} = "=B$line1+C$line1+D$line1+G$line1+H$line1+K$line1";
	$sheet_rs -> Range("T$line1")->{'Value'} = "=K$line1/S$line1";
	$sheet_rs -> Range("U$line1")->{'Value'} = "=(C$line1+K$line1)/S$line1";
	
	# STYLE

	$sheet_rs -> Range("A$line1:U$line1") -> {Borders} -> {LineStyle} = 1;
	$sheet_rs -> Range("A$line1:U$line1") -> {Borders} -> {Color} = 0x000000;
	
	$sheet_rs -> Range("A$line1") -> {Interior} -> {Color} = 0x0000FF;
	$sheet_rs -> Range("A$line1") -> {Font} -> {Color} = 0xFFFFFF;
	$sheet_rs -> Range("A$line1") -> {Font} -> {Bold} = 1;
	
	my $tmp_tag = $tag + 1;
	$sheet_rs -> Range("S$tmp_tag")->{'Value'} = "=S$line1/28800";
	$sheet_rs -> Range("S$tmp_tag")->{'NumberFormatLocal'} = "0.0000000";
	
	# 单元格宽度
	$sheet_rs -> Columns -> {ColumnWidth} = 8;

	# 样式
	$sheet_rs -> Rows -> {RowHeight} = 18;
	$sheet_rs -> Rows(1) -> {WrapText} = 1;
	$sheet_rs -> Range('A1:U1') -> {RowHeight} = 40;
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> {HorizontalAlignment} = xlHAlignCenter;
	$sheet_rs -> Columns -> {VerticalAlignment} = xlCenter;

	# 隐藏
	# $sheet_rs -> Columns("B:M") -> {ColumnWidth} = '0';

	# 保存
	# $sheet_rs -> Columns -> AutoFit();
	$sheet_rs -> SaveAs ($xls_doc_rs);
	print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Save Result To $xls_doc_rs";
}

#################################### BLANK ###########################################

sub blank_report
{
	my $TASK_INFO;
	my @KEYS;
	my $s_time = gettimeofday();
	
	my $TASK_ID;
	my $ITEM_ID;
	my $TASK_SRC;
	my $NICKNAME;
	my $TASK_STATUS;
	my $TASK_TITLE;
	my $TASK_PER;
	my $TASK_CREATER;
	my $TASK_CREATE_DATE;
	my $TASK_EXECUTER;
	my $TASK_TL;
	my $TASK_PROCESS_DATE;
	my $TASK_PLAN_DATE;
	my $TASK_NOTIFY_DATE;
	my $TASK_BUESSINESS_TYPE;
	my $TASK_ISSUESS_TYPE;
	my $TASK_HAS_USERID;

	my $workbook;
	my $sheet;
	my $data;
	my $tl;
	my $creater;
	my $str_ref;
	my $tmp_type;
	my $tmp_buniess_type;
	my $tmp_nickneme;
	my $UUID;
	
	$excel -> {'Visible'} = 0;
	$workbook = $excel -> Workbooks -> Open($BLANK_XLS_PATH, 1);
	$workbook -> {'ReadOnly'} = 1;
	$sheet = $workbook -> Worksheets($BLANK_XLS_SHEET_NAME);

	# 数量
	my $used_rows = $sheet -> {'UsedRange'} ->{'Rows'} -> {'Count'};
	my $used_count = $sheet -> {'UsedRange'} ->{'Columns'} -> {'Count'};
	if ($used_count eq '') {
		print 'ERROR : username blank sheet parsed error ...';
		$workbook -> Close({SaveChanges => 0});
		&error_found;
	}

	# 映射
	
	for (my $i=1; $i<=$used_count; $i++)
	{
		my $temp_val = decode('gbk', $sheet->Cells(1,$i)->{'Value'});
		Encode::_utf8_on($temp_val);
		if ($temp_val eq (BLANK_STATUS_LIST)[0]) {
			$TASK_ID = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[1]) {
			$ITEM_ID = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[2]) {
			$TASK_SRC = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[3]) {
			$NICKNAME = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[4]) {
			$TASK_STATUS = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[5]) {
			$TASK_TITLE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[6]) {
			$TASK_PER = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[7]) {
			$TASK_CREATER = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[8]) {
			$TASK_CREATE_DATE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[9]) {
			$TASK_EXECUTER = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[10]) {
			$TASK_TL = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[11]) {
			$TASK_PROCESS_DATE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[12]) {
			$TASK_PLAN_DATE = $i;	
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[13]) {
			$TASK_NOTIFY_DATE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[14]) {
			$TASK_BUESSINESS_TYPE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[15]) {
			$TASK_ISSUESS_TYPE = $i;
		} elsif ($temp_val eq (BLANK_STATUS_LIST)[16]) {
			$TASK_HAS_USERID = $i;
		} elsif ($temp_val eq '有无会员名') {
			$TASK_HAS_USERID = $i;
		}
	}

	for ( my $i=1+1; $i<=$used_rows; $i++)
	{
	
		$tl = $sheet-> Cells($i,$TASK_TL)->{'Value'};
		$tmp_nickneme = $sheet-> Cells($i,$NICKNAME)->{'Value'};
		
		# 处理 tl 为空的情况
		
		$str_ref = scalar($sheet-> Cells($i,$TASK_TL)->{'Value'});
		if ( "$str_ref" =~ /^-[0-9]+$/){
			$tl = '#N/A';
		}

		# 有会员名记数
		if ( $tmp_nickneme =~ /\S/){
			$data->{$tl}->{0} += 1;
			next;
		}
		
		# 会员编号为空
		$UUID = $sheet-> Cells($i,1)->{'Value'};
		($UUID eq '') and next;
		
		$tmp_buniess_type = $sheet-> Cells($i,$TASK_BUESSINESS_TYPE)->{'Value'};
		$tmp_type = decode('gbk', $tmp_buniess_type);
		Encode::_utf8_on($tmp_type);
		
		if ($tmp_type =~ /无效电话/) {
			$data->{$tl}->{2} += 1;
			next;
		}
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_ID'} = $sheet-> Cells($i,$TASK_ID)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'ITEM_ID'} = $sheet-> Cells($i,$ITEM_ID)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_SRC'} = $sheet-> Cells($i,$TASK_SRC)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'NICKNAME'} = $tmp_nickneme;
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_STATUS'} = $sheet-> Cells($i,$TASK_STATUS)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_TITLE'} = $sheet-> Cells($i,$TASK_TITLE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_PER'} = $sheet-> Cells($i,$TASK_PER)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_CREATER'} = $sheet-> Cells($i,$TASK_CREATER)->{'Value'};;
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_CREATE_DATE'} = $sheet-> Cells($i,$TASK_CREATE_DATE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_EXECUTER'} = $sheet-> Cells($i,$TASK_EXECUTER)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_TL'} = $tl;
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_PROCESS_DATE'} = $sheet-> Cells($i,$TASK_PROCESS_DATE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_PLAN_DATE'} = $sheet-> Cells($i,$TASK_PLAN_DATE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_NOTIFY_DATE'} = $sheet-> Cells($i,$TASK_NOTIFY_DATE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_BUESSINESS_TYPE'} = $sheet-> Cells($i,$TASK_BUESSINESS_TYPE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_ISSUESS_TYPE'} = $sheet-> Cells($i,$TASK_ISSUESS_TYPE)->{'Value'};
		$TASK_INFO ->{$i} -> {$UUID} -> {'TASK_HAS_USERID'} = $sheet-> Cells($i,$TASK_HAS_USERID)->{'Value'};
	}
	$workbook -> Close({SaveChanges => 0});
	
	my $e_time = gettimeofday();
	if (defined($TASK_INFO)) {
		print strftime("%Y-%m-%d %H:%M:%S", localtime())," : parse successfully ...";
		printf ("(%0.3fs)\n",($e_time-$s_time));
	} else {
		&error_found;
	}

	$excel -> {'Visible'} = 1;

	my $workbook_rs;
	my $sheet_rs;
	my $xls_doc_rs;

	$xls_doc_rs = "$cur_dir\\Blank Username Report.xlsx";
	eval {
		unlink $xls_doc_rs if (-e "$xls_doc_rs");
	};

	# 正解注解
	my $k;
	my $v;
	my $tag;
	my @orange = split("#", $BLANK_XLS_ORANGE);
	my @green = split("#", $BLANK_XLS_GREEN);
	my @purple = split("#", $BLANK_XLS_PURPLE);
	@KEYS = sort { $a <=> $b } keys($TASK_INFO);
	
	# 表头
	my @sheet_title = map {encode('gbk', $_)} (BLANK_STATUS_LIST);
	my @person_title = ('花名','组别','非无效电话空白会员名','错误备注','错误占比');
	my @person_title_gbk = map {encode('gbk',$_)} @person_title;
	my @data_title = ('组别','有会员名','无效电话','非无效电话空白会员名','正确备注','错误备注','错误占比','错误备注抽检数');
	my @data_title_gbk = map {encode('gbk',$_)} @data_title;
	
	# 生成
	$workbook_rs = $excel -> Workbooks -> Add();
	$sheet_rs =  $workbook_rs -> Worksheets(2);
	$sheet_rs->{'NAME'} = encode('gbk','正确注解');

	$sheet_rs -> Range('A1:Q1') -> {'Value'} = [@sheet_title];
	
	# 边框
	#$sheet_rs -> Range('A1:Q1') -> {Borders} -> {LineStyle} = 1;
	#$sheet_rs -> Range('A1:Q1') -> {Borders} -> {Color} = 0x000000;
	
	# 数据格式
	$sheet_rs -> Columns('A') -> {'NumberFormatLocal'} = "@";
	$sheet_rs -> Columns('B') -> {'NumberFormatLocal'} = "@";
	#$sheet_rs -> Columns('H') -> {'NumberFormatLocal'} = "0.0%";
	
	$tag = 2;
	foreach my $key (@KEYS){
		while (($k, $v)=each $TASK_INFO->{$key})
		{
			if ((map {$v->{'TASK_TITLE'} =~ /$_/i} @orange) or (map {$v->{'TASK_TITLE'} =~ /$_/i} @green) or (map {$v->{'TASK_TITLE'} =~ /$_/i} @purple) or ($v->{'TASK_TITLE'} =~ /^\s*$/)) {
				$sheet_rs -> Range("A$tag")->{'Value'} = $k;
				$sheet_rs -> Range("B$tag")->{'Value'} = $v->{'ITEM_ID'};
				$sheet_rs -> Range("C$tag")->{'Value'} = $v->{'TASK_SRC'};
				$sheet_rs -> Range("D$tag")->{'Value'} = $v->{'NICKNAME'};
				$sheet_rs -> Range("E$tag")->{'Value'} = $v->{'TASK_STATUS'};
				$sheet_rs -> Range("F$tag")->{'Value'} = $v->{'TASK_TITLE'};
				$sheet_rs -> Range("G$tag")->{'Value'} = $v->{'TASK_PER'};
				$sheet_rs -> Range("H$tag")->{'Value'} = $v->{'TASK_CREATER'};
				$sheet_rs -> Range("I$tag")->{'Value'} = $v->{'TASK_CREATE_DATE'};
				$sheet_rs -> Range("J$tag")->{'Value'} = $v->{'TASK_EXECUTER'};
				$sheet_rs -> Range("K$tag")->{'Value'} = $v->{'TASK_TL'};
				$sheet_rs -> Range("L$tag")->{'Value'} = $v->{'TASK_PROCESS_DATE'};
				$sheet_rs -> Range("M$tag")->{'Value'} = $v->{'TASK_PLAN_DATE'};
				$sheet_rs -> Range("N$tag")->{'Value'} = $v->{'TASK_NOTIFY_DATE'};
				$sheet_rs -> Range("O$tag")->{'Value'} = $v->{'TASK_BUESSINESS_TYPE'};
				$sheet_rs -> Range("P$tag")->{'Value'} = $v->{'TASK_ISSUESS_TYPE'};
				$sheet_rs -> Range("Q$tag")->{'Value'} = $v->{'TASK_HAS_USERID'};
				if (map {$v->{'TASK_TITLE'} =~ /$_/i} @orange) {
					$sheet_rs -> Range("F$tag") -> {Interior} -> {Color} = 0x00C0FF;
				} elsif (map {$v->{'TASK_TITLE'} =~ /$_/i} @green) {
					$sheet_rs -> Range("F$tag") -> {Interior} -> {Color} = 0x59BB9B;
				} elsif (map {$v->{'TASK_TITLE'} =~ /$_/i} @purple) {
					$sheet_rs -> Range("F$tag") -> {Interior} -> {Color} = 0xA03070;
				}
				$data->{$v->{'TASK_TL'}}->{1} += 1;
				$data->{$v->{'TASK_TL'}}->{9}->{$v->{'TASK_CREATER'}}->{1} += 1;	
				$tag++;
			} else {
				$data->{$v->{'TASK_TL'}}->{3} += 1;
				$data->{$v->{'TASK_TL'}}->{9}->{$v->{'TASK_CREATER'}}->{3} += 1;
			}
		}
	}
	
	# 0 有会员名
	# 1 正确注解
	# 2 无效电话
	# 3 错误注解
	# 9 -> 1 正确注解
	# 9 -> 2 无效电话
	# 9 -> 3 错误注解
	
	# 样式
	$sheet_rs -> Rows -> {RowHeight} = 18;
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> AutoFilter();
	
	# $sheet_rs -> Rows(1) -> {WrapText} = 1;
	# $sheet_rs -> Range('A1:U1') -> {RowHeight} = 40;

	# $sheet_rs -> Columns -> {HorizontalAlignment} = xlHAlignCenter;
	# $sheet_rs -> Columns -> {VerticalAlignment} = xlCenter;

	# 错误注解
	$sheet_rs =  $workbook_rs -> Worksheets(3);
	$sheet_rs->{'NAME'} = encode('gbk','错误注解');
	$sheet_rs -> Range('A1:Q1') -> {'Value'} = [@sheet_title];
	$sheet_rs -> Columns('A') -> {'NumberFormatLocal'} = "@";
	$sheet_rs -> Columns('B') -> {'NumberFormatLocal'} = "@";
	
	$tag = 2;
	foreach my $key (@KEYS){
		while (($k, $v)=each $TASK_INFO->{$key})
		{
			if ((map {$v->{'TASK_TITLE'} =~ /$_/i} @orange) or (map {$v->{'TASK_TITLE'} =~ /$_/i} @green) or (map {$v->{'TASK_TITLE'} =~ /$_/i} @purple) or ($v->{'TASK_TITLE'} =~ /^\s*$/)) {
				next;
			}
			$sheet_rs -> Range("A$tag")->{'Value'} = $k;
			$sheet_rs -> Range("B$tag")->{'Value'} = $v->{'ITEM_ID'};
			$sheet_rs -> Range("C$tag")->{'Value'} = $v->{'TASK_SRC'};
			$sheet_rs -> Range("D$tag")->{'Value'} = $v->{'NICKNAME'};
			$sheet_rs -> Range("E$tag")->{'Value'} = $v->{'TASK_STATUS'};
			$sheet_rs -> Range("F$tag")->{'Value'} = $v->{'TASK_TITLE'};
			$sheet_rs -> Range("G$tag")->{'Value'} = $v->{'TASK_PER'};
			$sheet_rs -> Range("H$tag")->{'Value'} = $v->{'TASK_CREATER'};
			$sheet_rs -> Range("I$tag")->{'Value'} = $v->{'TASK_CREATE_DATE'};
			$sheet_rs -> Range("J$tag")->{'Value'} = $v->{'TASK_EXECUTER'};
			$sheet_rs -> Range("K$tag")->{'Value'} = $v->{'TASK_TL'};
			$sheet_rs -> Range("L$tag")->{'Value'} = $v->{'TASK_PROCESS_DATE'};
			$sheet_rs -> Range("M$tag")->{'Value'} = $v->{'TASK_PLAN_DATE'};
			$sheet_rs -> Range("N$tag")->{'Value'} = $v->{'TASK_NOTIFY_DATE'};
			$sheet_rs -> Range("O$tag")->{'Value'} = $v->{'TASK_BUESSINESS_TYPE'};
			$sheet_rs -> Range("P$tag")->{'Value'} = $v->{'TASK_ISSUESS_TYPE'};
			$sheet_rs -> Range("Q$tag")->{'Value'} = $v->{'TASK_HAS_USERID'};
			$tag++;
		}
	}
	$sheet_rs -> Rows -> {RowHeight} = 18;
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> AutoFilter();
		
	# 个人汇总
	$sheet_rs =  $workbook_rs -> Worksheets(1);
	$sheet_rs->{'NAME'} = encode('gbk','个人汇总');
	$sheet_rs -> Range('A1:E1') -> {'Value'} = [@person_title_gbk];
	$sheet_rs -> Columns('E') -> {'NumberFormatLocal'} = "0.00%";
	
	
	my $tmp_sort_res;
	my $tmp_res;
	$tag = 1;
	while (($k, $v)=each $data)
	{
		if(!$v->{9}){
			next;
		}
		while ((my $q, my $r)=each $v->{9}) {
			$tmp_sort_res->{$tag}->{0} = $k;
			$tmp_sort_res->{$tag}->{1} = $q;
			$tmp_sort_res->{$tag}->{2} = $r->{1}+$r->{3};
			$tmp_sort_res->{$tag}->{3} = $r->{3};
			$tmp_res->{$tag} = sprintf("%f",$r->{3}/($r->{1}+$r->{3}));
			$tag++;
		}
	}
	my @KEYS = sort {$tmp_res->{$b} <=> $tmp_res->{$a}} keys($tmp_res);
	
	$tag = 2;
	foreach my $key (@KEYS)
	{
		$sheet_rs -> Range("A$tag")->{'Value'} = $tmp_sort_res->{$key}->{0};
		$sheet_rs -> Range("B$tag")->{'Value'} = $tmp_sort_res->{$key}->{1};
		$sheet_rs -> Range("C$tag")->{'Value'} = $tmp_sort_res->{$key}->{2};
		$sheet_rs -> Range("D$tag")->{'Value'} = sprintf("%u",$tmp_sort_res->{$key}->{3});
		$sheet_rs -> Range("E$tag")->{'Value'} = $tmp_res->{$key};
		$tag++;
	}
	$sheet_rs -> Rows -> {RowHeight} = 18;
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> AutoFilter();
	$sheet_rs -> Columns -> AutoFit();
	
	# 全局统计
	$sheet_rs =  $workbook_rs -> Worksheets -> Add();
	$sheet_rs->{'NAME'} = strftime("%Y-%m-%d", localtime());
	$sheet_rs -> Range('A1:H1') -> {'Value'} = [@data_title_gbk];
	$sheet_rs -> Columns('G') -> {'NumberFormatLocal'} = "0.00%";
	
	my $tmp_sort_ref;
	while (($k, $v) = each $data) {
		my $tmp_sort_cnt = $v->{'1'}+$v->{'3'};
		if ($tmp_sort_cnt == 0){
			$tmp_sort_ref->{$k} = 0;
		} else {
			$tmp_sort_ref->{$k} = sprintf("%f", $v->{'3'}/$tmp_sort_cnt);
		}	
	}
	my @KEYS = sort {$tmp_sort_ref->{$b} <=> $tmp_sort_ref->{$a}} keys($tmp_sort_ref);
	$tag = 2;
	foreach my $key (@KEYS){
		$sheet_rs -> Range("A$tag")->{'Value'} = $key;
		$sheet_rs -> Range("B$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'0'});
		$sheet_rs -> Range("C$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'2'});
		$sheet_rs -> Range("D$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'3'}+$data->{$key}->{'1'});
		$sheet_rs -> Range("E$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'1'});
		$sheet_rs -> Range("F$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'3'});
		$sheet_rs -> Range("G$tag")->{'Value'} = $tmp_sort_ref->{$key};
		$sheet_rs -> Range("H$tag")->{'Value'} = sprintf("%u",$data->{$key}->{'3'});
		$tag++;
	}
	my $tag_1 = $tag - 1;
	$sheet_rs -> Range("A$tag")->{'Value'} = encode('gbk','汇总');
	$sheet_rs -> Range("B$tag")->{'Value'} = "=SUM(B2:B$tag_1)";
	$sheet_rs -> Range("C$tag")->{'Value'} = "=SUM(C2:C$tag_1)";
	$sheet_rs -> Range("D$tag")->{'Value'} = "=SUM(D2:D$tag_1)";
	$sheet_rs -> Range("E$tag")->{'Value'} = "=SUM(E2:E$tag_1)";
	$sheet_rs -> Range("F$tag")->{'Value'} = "=SUM(F2:F$tag_1)";
	$sheet_rs -> Range("G$tag")->{'Value'} = "=F$tag/D$tag";
	$sheet_rs -> Range("H$tag")->{'Value'} = "=SUM(H2:H$tag_1)";
	
	$sheet_rs -> Rows -> {RowHeight} = 18;
	$sheet_rs -> Columns -> {Font} -> {Name} = 'Calibri';
	$sheet_rs -> Columns -> {Font} -> {Size} = 10;
	$sheet_rs -> Columns -> AutoFit();
	
	$sheet_rs -> SaveAs ($xls_doc_rs);
	print strftime("%Y-%m-%d %H:%M:%S", localtime())," : Save Result To $xls_doc_rs";
}