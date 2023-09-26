// ---------- Kombinera csv-filer med Power Query i Excel ----------



// Steg 0. Ladda ned rapporterna AD-8 och AD-2 från 33900.edu.stockholm.se, Rapporter.



// Steg 1. Öppna .csv-filerna med Anteckningar och radera allt innan första kolumnen, d.v.s. så att filerna börjar med "COS-nummmer;" respektive "Efternamn;", spara.



// Steg 2. Starta Excel, gå till "Data", Hämta data > Starta Power Query-redigeraren, Ny källa > Andra källor > Tom fråga, Avancerad redigerare, klistra där in Power Query-koden nedan.



// Steg 3. Ändra i koden följande:



//			a) byt sökvägarna på raderna som börjar med Source0 = och Source1 =, så att Power Query hittar de nedladdade rapporterna från steg 1.
//				Tips: Shifttangent+högerklick > Kopiera som sökväg kopierar hela filens sökväg med citationstecken.

//			b) välj ut vilka kolumner från filerna som rapporten ska visa. Kolumnerna måste ha unika namn, heter de lika så ändra detta i ena csv-filens första rad i Anteckningar.

//			c) bestäm i vilken ordning kolumnerna ska visas.

//			d) kopiera den färdiga frågan och spara i en textfil till framtiden.
//				Tips: Hela det här dokumentet är kommenterat och går att spara i en Power Query-fråga.



// Steg 4. Tryck Klart, Stäng och läs in.



// Power Query koden nedan:



let



	//	a) Peka ut första csv-filen:

	Source0 = Csv.Document(File.Contents("C:\sökväg\till\AD-8 Export Förteckning över Windows-datorer.csv"), [Delimiter = ";", Encoding = 1200, QuoteStyle = QuoteStyle.None]),
	HeaderRow0 = Table.PromoteHeaders(Source0, [PromoteAllScalars=true]),



	//	a) Peka ut andra csv-filen:

	Source1 = Csv.Document(File.Contents("C:\sökväg\till\AD-2 Export Förteckning över användarkonton.csv"), [Delimiter = ";", Encoding = 1200, QuoteStyle = QuoteStyle.None]),
	HeaderRow1 = Table.PromoteHeaders(Source1, [PromoteAllScalars=true]),



	//	b) Välj vilka kolumner vi ska ha från "Source0" (all tillgängliga kolumner står på första raden i csv-filen):

	SelectColumns0 = Table.SelectColumns(HeaderRow0, {"COS-nummer", "Datorobjekt skapat", "Värdnamn (DNS)", "Operativsystem", "Primäranvändare (UID)"}),



	//	b) Välj vilka kolumner vi ska ha från "Source1" (all tillgängliga kolumner står på första raden i csv-filen):

	SelectColumns1 = Table.SelectColumns(HeaderRow1, {"Efternamn", "Förnamn", "Hemkatalog", "Kontot skapat", "AnvändarID (UID)", "AnvändarID (UPN)", "Verksamhet", "Klass"}),



	//	c) Här bestäms i vilken ordning som kolumnerna ska visas, från vänster till höger

	ColumnOrder = {"COS-nummer", "AnvändarID (UID)", "Förnamn", "Efternamn", "AnvändarID (UPN)", "Klass", "Värdnamn (DNS)", "Operativsystem", "Hemkatalog", "Datorobjekt skapat"},
	ReorderedColumns = Table.ReorderColumns(Expanded, ColumnOrder),



	//	Detta avsnitt sammanfogar tabellerna genom kolumnerna "Primäranvändare (UID)" och "AnvändarID (UID)":

	Merged = Table.NestedJoin(SelectColumns0, "Primäranvändare (UID)", SelectColumns1, "AnvändarID (UID)", "MergedData"),
	Expanded = Table.ExpandTableColumn(Merged, "MergedData", Table.ColumnNames(SelectColumns1)),
	Custom = ReorderedColumns



in
	Custom



	// ---------- Slut på instruktion och Power Query-kod, uppdaterad 2023-09-26 ----------
	
