sub group1_5555_tank_straps_straps()

if isworkbookopen("excel.xlsx") = false then
          set wbk = workbooks.open("excel.xlsx")
else
   exit sub

end if

set ws1 = sheets("tanks")

 dim sqlstring as string
 dim cnn as adodb.connection
 set cnn = new adodb.connection

   cnn.connectionstring = "provider=oraoledb.oracle;data source=server;user id=username;password=password;"
   cnn.connectiontimeout = 90
   cnn.open
   if cnn.state = adstateopen then
   
   else
      msgbox "sorry. connection failed"
      exit sub
   end if

dim x as integer
for x = 2 to 4000
if ws1.range("d" & x).value <> "" then
with ws1
   tankid = ws1.range("k" & x).value
   recorddate = ws1.range("d" & x).value
   divisionid = 95
   if ws1.range("f" & x).value <> "" then
   dailyvolume = ws1.range("f" & x).value
   elseif ws1.range("t" & x) = "yes" then
   dailyvolume = 0
   elseif ws1.range("r" & x).value <= 1325 then
   dailyvolume = (-0.0019 * ((ws1.range("e" & x).value + ws1.range("p" & x).value) ^ 3)) + (0.272 * ((ws1.range("e" & x).value + ws1.range("p" & x).value) ^ 2)) + (5.7178 * (ws1.range("e" & x).value + ws1.range("p" & x).value)) - 19.303
    elseif ws1.range("r" & x).value <= 2200 then
   dailyvolume = (-0.0015 * ((ws1.range("e" & x) + ws1.range("p" & x).value) ^ 3)) + (0.2885 * ((ws1.range("e" & x) + ws1.range("p" & x).value) ^ 2)) + (7.1159 * (ws1.range("e" & x) + ws1.range("p" & x).value)) - 27.885
      elseif ws1.range("r" & x).value <= 4400 then
dailyvolume = (-0.0014 * ((ws1.range("e" & x) + ws1.range("p" & x).value) ^ 3)) + (0.3283 * ((ws1.range("e" & x) + ws1.range("p" & x).value) ^ 2)) + (11.742 * (ws1.range("e" & x) + ws1.range("p" & x).value)) - 50.125
     else: dailyvolume = ""
   end if

   addedvolume = ws1.range("g" & x).value
   tanksightglassreading = ws1.range("e" & x).value
   if ws1.range("t" & x) = "yes" then
   comments = "empty;" & ws1.range("h" & x).value
   else
   comments = ws1.range("h" & x).value
   end if
   username = ws1.range("i" & x).value
   chemicalid = ws1.range("j" & x).value
end with

recorddate = format(range("d" & x).value, "yyyy/mm/dd")

        
        sqlstring = "insert into volume vol (vol.reading, vol.tank_id, vol.record_dt, vol.division_id, vol.daily_volume, vol.added_volume,vol.comments, vol.username, vol.chemical_id) values('" & tanksightglassreading & "','" & tankid & "','" & recorddate & "','" & divisionid & "','" & dailyvolume & "','" & addedvolume & "','" & comments & "','" & username & "','" & chemicalid & "')"       
        ws1.range("z10") = sqlstring
        cnn.execute sqlstring, , adcmdtext

end if
        
 next x
   cnn.close
   set cnn = nothing
  
activeworkbook.refreshall
application.calculateuntilasyncqueriesdone
     
        range("z10").select
    selection.clearcontents
      range("a1").select
     
    workbooks("group1_5555_tank_straps.xlsx").save
    workbooks("group1_5555_tank_straps.xlsx").close
   
end sub
