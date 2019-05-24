<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>


<html>
<head>
<title>Sistem Pengurusan OT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">


<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.id.value=="")
	{
		alert("Sila masukkan id");
		f.id.focus();
		return false
	}
	if (f.bulan.value=="")
	{
		alert("Sila masukkan bulan");
		f.bulan.focus();
		return false
	}
	if (f.tahun.value=="")
	{
		alert("Sila masukkan tahun");
		f.tahun.focus();
		return false
	}
	
}
//  End -->
</script>
</head>


<body leftmargin="0" topmargin="0">


<%response.cookies("amenu") = "sp13.asp"%>
<% 
proses = request.form("proses")
if proses="" then 
	papar
else
    id = request.form("id")
    bulan = request.form("bulan")
	tahun = request.form("tahun")
	
		if proses="Simpan" then 
			simpan
		elseif proses="Hapus" then 
			hapus
		end if
end if




sub simpan()
proses=request("proses")

if proses="Simpan" then
rekod=""
bln = mid(bulan,3,2)
thn = mid(tahun,5,4)


    q0="select * from payroll.selenggara_kerja where tahun like '"& thn &"' and bulan like '"& bln &"' and id like '"& id &"'  "
    set r0=objConn.Execute(q0)
	if r0.eof then
		q2="Insert into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan &"','"& tahun &"')"
		objConn.Execute(q2) 
		rekod="simpan"
		msg="Rekod id "& id &" bagi bulan "& bulan &" / tahun "& tahun &" selesai ditambah"
	else
		rekod="duplicate"
		msg="Rekod tahun untuk id "& id &"  , bulan "& bln &"  ,tahun "& thn &" telah wujud.<br>Proses tidak diteruskan"
	end if
	papar
	if rekod<> "" then %><br>
    
    
    <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
      <TBODY> 
      <tr>
        <td align="center"><%=msg%></td>
      </tr>
    </tbody>
    </table>
    
    
<% end if
end if
end sub


sub papar()
q1="select rowid ,id as id , tahun as tahun, bulan as bulan from payroll.selenggara_kerja order by tahun , bulan  asc"
Set rs1 = objConn.Execute(q1)
id = rs1 ("id")
%>
<br>



<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all width="55%" align="center"> 
  <tr class="hd"> 
    <td width="40%" align="center" class="hd">Selenggara Kelulusan Tuntutan Melebihi Sepertiga</td>
  </tr>
  
   <form name="test0" method="post" action="sp13.asp" onSubmit="return check(this)">
   <tr valign="middle"> 
    <td>&nbsp;&nbsp;&nbsp;&nbsp;No Pekerja : <input type="text" name="id" size="6" maxlength="6">
    &nbsp;&nbsp;
     <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->
    </td>
    </tr>
    
    
  <tr valign="middle"> 
    <td>&nbsp;&nbsp;&nbsp;&nbsp;Bulan      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: <input type="text" name="bulan" size="2" maxlength="2">
    &nbsp;&nbsp;
    <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	--></td>
    </tr>
    
   <tr valign="middle"> 
    <td>&nbsp;&nbsp;&nbsp;&nbsp;Tahun &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: <input type="text" name="tahun" size="4" maxlength="4">
    &nbsp;&nbsp;
    <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <input type="submit" name="proses" value="Simpan">
    </form>
  </tr>
</table>



<br><br>
<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY>
  <tr>
  	<td width="2%" class="hd" align="center"><B>Bil</B></td>
    <td width="21%" align="center" class="hd"><B>Pekerja</B></td>
     <td width="19%" align="center" class="hd"><B>Jabatan</B></td>
    <td width="15%" align="center" class="hd"><B>Bulan</B></td>
    <td width="14%" align="center" class="hd"><B>Tahun</B></td>   
    <td width="29%" align="center" class="hd"><B>Proses Selenggara Data</B>
    </td>
  </tr>
  </tbody> 
</table>

<%  
bil=0
Do While Not rs1.eof 
bil=bil + 1
id = rs1 ("id")
mm = "select nama , lokasi from payroll.paymas where no_pekerja='"& id & "' "
Set rsmm = objConn.Execute(mm)
'response.write mm
nama = rsmm ("nama")
lokasi = rsmm ("lokasi")


nn = "select keterangan from iabs.jabatan where kod='"& lokasi & "' "
Set rsnn = objConn.Execute (nn)

ket = rsnn ("keterangan")

%>

<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
  	<td width="2%" align="center"><%=bil%></td>
    <td width="21%" align="left"><%=rs1("id")%>-<%=rsmm("nama")%></td>
    <td width="19%" align="left">(<%=rsmm("lokasi")%>) <%=ket%></td>
    <td width="15%" align="center"><%=rs1("bulan")%></td>
    <td width="14%" align="center"><%=rs1("tahun")%></td>
    <form name="test_hapus<%=bil%>"  method="post" action="sp13.asp"> 
	<input type="hidden"  name="rowid"   value="<%=rs1("rowid")%>">
    <input type="hidden"  name="id"      value="<%=rs1("id")%><">
    <input type="hidden"  name="bulan"   value="<%=rs1("bulan")%>">
    <input type="hidden"  name="tahun"   value="<%=rs1("tahun")%>">    
    <td width="29%" align="center">
      <input type="submit" name="proses" value="Hapus" onClick="return confirm('Anda Pasti Hapus Rekod?')">
    </td>
    </form>
  </tr>
   <% 
  rs1.movenext
  loop 
  %>
  </tbody> 
</table>



<% end sub %>
<%sub hapus()
a=request("rowid")
id=request("id")
bulan=request("bulan")
tahun=request("tahun")
If a<>"" then
q1="Delete from payroll.selenggara_kerja where rowid like '"& a &"'"
objConn.Execute(q1)
papar()%>

<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod data pekerja <%=id%> bagi bulan <%=bulan%> / <%=tahun%>  telah dihapus.</td>
  </tr>
</tbody></table>


<%
end if
end sub%>
<%sub edit()
rowid=request("rowid")
id=request("id")
bulan=request("bulan")
tahun=request("tahun")
tahun_asal=request("tahun_asal")


q1="Update payroll.selenggara_kerja set tahun = to_date('"& tahun &"','yyyy') , bulan = to_date('"& bulan &"','dd') where rowid like '"& rowid &"'"
papar()%>
<br>


<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
    <td align="center"> Rekod asal no pekerja <%=id_asal%> selesai diedit kepada <%=id%>.</td>
  </tr>
</tbody></table>
<%end sub%>
</body>
</html>