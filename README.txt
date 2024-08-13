Detta repo innehåller två skript för att konvertera SIE-filer. Jag har sparat källkoden - som vi inte har skapat - i filen "sie_converter_source.py" och den nya modifierade koden som "sie_converter_source". 

Source: Denna kod är skriven för att hantera SIE-filer av formatet SIE4. Vi har aldrig provat med tidigare versioner men då SIE1 - SIE4 till min vetskap endast bygger på datan och inte ändrar formatet så är det möjligt att det kommer funka. 
Vad denna kod inte är skriven för att klara av vissa SIE filer där verifikationsserien inte anges. Detta löser vi i den modifierade versionen. Det är mycket möjligt att det finns andra buggar som vi ännu inte har upptäckt. 

Modified: Denna kod är våran interna modifikation av källkoden och det är här vi gör all påbyggnad. 

Kommentarer: Det finns en rad förbättringar som skulle kunna förbättra källkoden då logiken är väldigt statiskt. Detta skulle kräva att vi sätter oss in oss djupare i hur SIE-filer funkar vilket skulle vara tidskrävande. 

Structure_data: Ibland blir det för mycket data i TRANSACTION filen för att Excel ska kunna hantera detta. Detta skript delar denna fil i fyra lika stora delar så att de går att öppna i Excel. Jag tycker inte att detta är optimalt och vi skulle kunna göra vidare förberedelser av datan i kod. Detta är ett ämne som senare bör diskuteras. 

Skillnaden mellan BIG Travel .SIE filer och Mockberg .SI filer: 

Big Travel #VER
#VER "" VINV000017943 20230101 
{
#TRANS 2640 {2 "912"} 83130.00 
#TRANS 2443 {2 "912"} -1519105.00 
#TRANS 4052 {2 "912"} 1435975.00 
}

OBS: "" anges som verifikationsserie

Mockberg #VER
#VER     AP	173325	20240103	""
{
   #TRANS   1931	{}	    -7205.74	20240103	"CHK - 303938"
   #TRANS   2440	{}	     7306.29	20240103	"CHK - 303938"
   #TRANS   3960	{}	     -100.55	20240103	"CHK - 303938"
}

OBS: AP anges som verifikationsserie
