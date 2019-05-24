<!--#include file="connection.asp" -->
<%



a="SELECT table_name, column_name, data_type, data_length FROM USER_TAB_COLUMNS WHERE table_name = 'KEHADIRAN'"
set qa = objekConn.Execute(a)

do while not qa.eof 
	response.Write(qa("column_name") & ", "& qa("data_type") &", "& qa("data_length") &"<br>")
qa.movenext
loop

'f="CREATE VIEW sispot_kehadiran AS SELECT no_pekerja,tarikh,time1,time2,time3,time4,komen from kehadiran.kehadiran where tarikh>=to_date('01012015','mmddyyyy')"
'objekConn.Execute(f)


a="select nvl(time1,'0') time1,nvl(time2,'0') time2,nvl(time3,'0') time3,nvl(time4,'0') time4,komen "
a = a & "from kehadiran.sispot_kehadiran where no_pekerja=19185 and tarikh like to_date('12192015','mmddyyyy')"
set qa = objekConn.Execute(a)
do while not qa.eof 
	response.Write(qa("time1") & ", "& qa("time2") &", "& qa("time3") &", "& qa("time4") &", "& qa("komen") &"<br>")
qa.movenext
loop
%>