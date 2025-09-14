# INTRODUCERE ÎN EXCEL VBA
## Prelegere de 4 ore cu exemple practice

---

## CUPRINS

**Ora 1: Fundamentele VBA**
- Ce este VBA și de ce să îl folosim
- Editorul VBA și înregistrarea macro-urilor
- Variabile și tipuri de date
- Primul nostru macro

**Ora 2: Structuri de control**
- Instrucțiuni condiționale (If-Then-Else)
- Bucle (For, While, Do)
- Select Case pentru decizii multiple

**Ora 3: Lucrul cu foi de calcul**
- Obiecte și colecții în Excel VBA
- Manipularea celulelor și range-urilor
- Formatarea și stilizarea datelor

**Ora 4: Aplicații avansate**
- Funcții definite de utilizator
- Gestionarea erorilor
- Proiect final: Analizator de vânzări

---

# ORA 1: FUNDAMENTELE VBA

## Ce este VBA?

**Visual Basic for Applications (VBA)** este un limbaj de programare dezvoltat de Microsoft pentru automatizarea sarcinilor în aplicațiile Office, inclusiv Excel.

### Avantajele VBA:
- Automatizarea sarcinilor repetitive
- Crearea de funcții personalizate
- Interfețe interactive pentru utilizatori
- Procesarea rapidă a volumelor mari de date

---

## Accesarea editorului VBA

**Pașii pentru deschiderea editorului VBA:**
1. Apăsați `Alt + F11` în Excel
2. Sau accesați Developer → Visual Basic
3. Pentru activarea Developer tab: File → Options → Customize Ribbon → Developer

**Părțile principale ale editorului VBA:**
- **Project Explorer**: Arată structura proiectului
- **Code Window**: Zona pentru scrierea codului
- **Properties Window**: Proprietățile obiectelor selectate
- **Immediate Window**: Pentru testarea rapidă a codului

---

## Primul nostru macro

Să începem cu un exemplu simplu care afișează un mesaj:

```vba
Sub PrimulMeuMacro()
    ' Acesta este primul nostru macro
    ' Linia de mai jos afiseaza un mesaj simplu
    MsgBox "Bun venit la cursul de VBA!"
End Sub
```

**Pentru rularea macro-ului:**
- Apăsați `F5` în editorul VBA
- Sau folosiți Developer → Macros → Run

---

## Variabile și tipuri de date

Variabilele sunt containere pentru stocarea datelor. În VBA, putem declara variabile cu instrucțiunea `Dim`.

### Tipurile principale de date:

```vba
Sub TipuriDeDate()
    ' Declararea variabilelor cu tipuri specifice
    Dim numeClient As String        ' Text (siruri de caractere)
    Dim varsta As Integer          ' Numere intregi (-32,768 la 32,767)
    Dim pret As Double            ' Numere cu zecimale
    Dim esteActiv As Boolean      ' Valori True/False
    Dim dataComanda As Date       ' Date si ore
    
    ' Atribuirea valorilor
    numeClient = "Alfreds Futterkiste"  ' Folosim ghilimele pentru text
    varsta = 25
    pret = 123.45
    esteActiv = True
    dataComanda = #12/31/2024#    ' Datele se incadreaza in #
    
    ' Afisarea valorilor in ferestra Immediate (Ctrl+G)
    Debug.Print "Client: " & numeClient
    Debug.Print "Varsta: " & varsta
    Debug.Print "Pret: " & pret
    Debug.Print "Este activ: " & esteActiv
    Debug.Print "Data comenzii: " & dataComanda
End Sub
```

---

## Lucrul cu datele din fișierul Northwind

Să vedem cum putem accesa datele din fișierul nostru Northwind:

```vba
Sub CitireClientNorthwind()
    ' Declararea variabilelor pentru stocarea informatiilor clientului
    Dim numeCompanie As String
    Dim oras As String
    Dim tara As String
    Dim ws As Worksheet
    
    ' Setarea referintei catre foaia Customers
    Set ws = ThisWorkbook.Worksheets("Customers")
    
    ' Citirea datelor primului client (linia 2, coloana A, B, C, D)
    numeCompanie = ws.Cells(2, 2).Value  ' Coloana B - CompanyName
    oras = ws.Cells(2, 3).Value          ' Coloana C - City  
    tara = ws.Cells(2, 4).Value          ' Coloana D - Country
    
    ' Afisarea informatiilor intr-un mesaj
    MsgBox "Primul client din baza de date:" & vbCrLf & _
           "Companie: " & numeCompanie & vbCrLf & _
           "Oras: " & oras & vbCrLf & _
           "Tara: " & tara
           
    ' vbCrLf = linie noua (carriage return + line feed)
    ' _ = continuarea liniei pe linia urmatoare
End Sub
```

---

## Exercițiu practic Ora 1

Creați un macro care să:
1. Citească numele primilor 3 angajați din foaia Employees
2. Calculeze vechimea lor în companie (diferența dintre data angajării și ziua de azi)
3. Afișeze rezultatele într-un mesaj

```vba
Sub ExercitiiAngajati()
    Dim ws As Worksheet
    Dim i As Integer
    Dim nume As String
    Dim prenume As String
    Dim dataAngajare As Date
    Dim vechime As Integer
    Dim mesaj As String
    
    ' Setarea referintei catre foaia Employees
    Set ws = ThisWorkbook.Worksheets("Employees")
    
    mesaj = "Informatii angajati:" & vbCrLf & vbCrLf
    
    ' Parcurgerea primilor 3 angajati (liniile 2, 3, 4)
    For i = 2 To 4
        ' Citirea datelor angajatului
        nume = ws.Cells(i, 2).Value      ' Coloana B - LastName
        prenume = ws.Cells(i, 3).Value   ' Coloana C - FirstName  
        dataAngajare = ws.Cells(i, 5).Value ' Coloana E - HireDate
        
        ' Calcularea vechimii in ani
        vechime = Year(Date) - Year(dataAngajare)
        
        ' Adaugarea informatiilor la mesaj
        mesaj = mesaj & prenume & " " & nume & ": " & vechime & " ani" & vbCrLf
    Next i
    
    ' Afisarea mesajului final
    MsgBox mesaj
End Sub
```

---

# ORA 2: STRUCTURI DE CONTROL

## Instrucțiuni condiționale (If-Then-Else)

Instrucțiunile condiționale ne permit să executăm cod diferit în funcție de anumite condiții:

```vba
Sub AnalizaPretProdus()
    Dim ws As Worksheet
    Dim pretProdus As Double
    Dim numeProdus As String
    Dim categoriePret As String
    
    ' Setarea referintei catre foaia Products
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Citirea informatiilor despre primul produs (linia 2)
    numeProdus = ws.Cells(2, 2).Value   ' Coloana B - ProductName
    pretProdus = ws.Cells(2, 4).Value   ' Coloana D - UnitPrice
    
    ' Clasificarea produsului dupa pret
    If pretProdus < 10 Then
        categoriePret = "Ieftin"
    ElseIf pretProdus < 50 Then
        categoriePret = "Mediu"
    ElseIf pretProdus < 100 Then
        categoriePret = "Scump"
    Else
        categoriePret = "Foarte scump"
    End If
    
    ' Afisarea rezultatului
    MsgBox "Produsul '" & numeProdus & "' costa " & pretProdus & " si este " & categoriePret
End Sub
```

---

## Bucle For - parcurgerea datelor

Bucla For este utilă când știm exact de câte ori vrem să repetăm o operațiune:

```vba
Sub NumarareProdusePerCategorie()
    Dim ws As Worksheet
    Dim i As Integer
    Dim ultimaLinie As Integer
    Dim categorieCautata As String
    Dim categorieGasita As String
    Dim contor As Integer
    
    ' Setarea referintei catre foaia Products
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Gasirea ultimei linii cu date
    ultimaLinie = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Categoria pe care o cautam
    categorieCautata = "Beverages"
    contor = 0
    
    ' Parcurgerea tuturor produselor (incepand din linia 2)
    For i = 2 To ultimaLinie
        ' Citirea categoriei produsului curent
        categorieGasita = ws.Cells(i, 6).Value  ' Coloana F - CategoryName
        
        ' Verificarea daca categoria se potriveste
        If categorieGasita = categorieCautata Then
            contor = contor + 1
            ' Afisarea numelui produsului in fereastra Immediate
            Debug.Print "Produs gasit: " & ws.Cells(i, 2).Value
        End If
    Next i
    
    ' Afisarea rezultatului final
    MsgBox "Au fost gasite " & contor & " produse din categoria " & categorieCautata
End Sub
```

---

## Bucle While - condiții de continuare

Bucla While continuă să se execute cât timp o condiție este adevărată:

```vba
Sub CautarePrimulClientDinGermany()
    Dim ws As Worksheet
    Dim liniaCurenta As Integer
    Dim taraGasita As String
    Dim clientGasit As Boolean
    Dim numeCompanie As String
    
    ' Setarea referintei catre foaia Customers
    Set ws = ThisWorkbook.Worksheets("Customers")
    
    liniaCurenta = 2    ' Incedem de la prima linie cu date
    clientGasit = False
    
    ' Continuam sa cautam pana gasim un client din Germania
    While Not clientGasit And liniaCurenta <= 100  ' Limitam cautarea la primele 100 linii
        ' Citirea tarii clientului curent
        taraGasita = ws.Cells(liniaCurenta, 4).Value  ' Coloana D - Country
        
        ' Verificarea daca am gasit Germania
        If taraGasita = "Germany" Then
            clientGasit = True
            numeCompanie = ws.Cells(liniaCurenta, 2).Value  ' Coloana B - CompanyName
        Else
            liniaCurenta = liniaCurenta + 1  ' Trecem la urmatorul client
        End If
    Wend
    
    ' Afisarea rezultatului
    If clientGasit Then
        MsgBox "Primul client din Germania este: " & numeCompanie & " (linia " & liniaCurenta & ")"
    Else
        MsgBox "Nu s-a gasit niciun client din Germania in primele 100 de linii"
    End If
End Sub
```

---

## Select Case - decizii multiple

Select Case este o alternativă elegantă la multe instrucțiuni If-ElseIf:

```vba
Sub ClassificareVanzariPerAngajat()
    Dim ws As Worksheet
    Dim idAngajat As Integer
    Dim numeAngajat As String
    Dim performanta As String
    
    ' Setarea referintei catre foaia Orders pentru a numara comenzile per angajat
    Set ws = ThisWorkbook.Worksheets("Orders")
    
    ' Sa analizam performanta pentru angajatul cu ID 1
    idAngajat = 1
    
    ' Numararea comenzilor pentru acest angajat
    Dim numarComenzi As Integer
    numarComenzi = Application.WorksheetFunction.CountIf(ws.Range("I:I"), idAngajat)
    
    ' Clasificarea performantei in functie de numarul de comenzi
    Select Case numarComenzi
        Case Is < 50
            performanta = "Performanta scazuta"
        Case 50 To 100
            performanta = "Performanta medie"
        Case 101 To 200
            performanta = "Performanta buna"
        Case Is > 200
            performanta = "Performanta excelenta"
        Case Else
            performanta = "Date nedisponibile"
    End Select
    
    ' Gasirea numelui angajatului din foaia Employees
    Dim wsEmp As Worksheet
    Set wsEmp = ThisWorkbook.Worksheets("Employees")
    numeAngajat = wsEmp.Cells(idAngajat + 1, 3).Value & " " & wsEmp.Cells(idAngajat + 1, 2).Value
    
    ' Afisarea rezultatului
    MsgBox "Angajatul " & numeAngajat & " are " & numarComenzi & " comenzi - " & performanta
End Sub
```

---

## Exercițiu practic Ora 2

Creați un macro care să analizeze toate comenzile din 2021 și să clasifice lunile după volumul de vânzări:

```vba
Sub AnalizeVanzariPeLuni2021()
    Dim ws As Worksheet
    Dim i As Integer
    Dim ultimaLinie As Integer
    Dim dataComanda As Date
    Dim luna As Integer
    Dim anulComanda As Integer
    Dim valoareComanda As Double
    Dim totalLuni(1 To 12) As Double  ' Array pentru totalurile pe luni
    Dim numeLuni As String
    
    ' Setarea referintei catre foaia Orders
    Set ws = ThisWorkbook.Worksheets("Orders")
    
    ' Gasirea ultimei linii cu date
    ultimaLinie = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Parcurgerea tuturor comenzilor
    For i = 2 To ultimaLinie
        ' Citirea datei comenzii
        dataComanda = ws.Cells(i, 2).Value  ' Coloana B - Order Date
        anulComanda = Year(dataComanda)
        
        ' Verificam daca comanda este din 2021
        If anulComanda = 2021 Then
            luna = Month(dataComanda)
            ' Calcularea valorii comenzii (UnitPrice * Quantity)
            valoareComanda = ws.Cells(i, 6).Value * ws.Cells(i, 7).Value  ' Coloana F * G
            ' Adaugarea la totalul lunii respective
            totalLuni(luna) = totalLuni(luna) + valoareComanda
        End If
    Next i
    
    ' Analizarea si afisarea rezultatelor
    For i = 1 To 12
        ' Determinarea numelui lunii
        Select Case i
            Case 1: numeLuni = "Ianuarie"
            Case 2: numeLuni = "Februarie"  
            Case 3: numeLuni = "Martie"
            Case 4: numeLuni = "Aprilie"
            Case 5: numeLuni = "Mai"
            Case 6: numeLuni = "Iunie"
            Case 7: numeLuni = "Iulie"
            Case 8: numeLuni = "August"
            Case 9: numeLuni = "Septembrie"
            Case 10: numeLuni = "Octombrie"
            Case 11: numeLuni = "Noiembrie"
            Case 12: numeLuni = "Decembrie"
        End Select
        
        ' Afisarea rezultatului pentru luna curenta
        If totalLuni(i) > 0 Then
            Debug.Print numeLuni & ": " & Format(totalLuni(i), "Currency")
        End If
    Next i
    
    MsgBox "Analiza completa! Vezi rezultatele in fereastra Immediate (Ctrl+G)"
End Sub
```

---

# ORA 3: LUCRUL CU FOI DE CALCUL

## Obiecte și colecții în Excel VBA

Excel VBA folosește un model de obiecte ierarhic. Să înțelegem principalele obiecte:

```vba
Sub ExempluObiecte()
    Dim wb As Workbook          ' Cartea de lucru (fisierul Excel)
    Dim ws As Worksheet         ' Foaia de calcul
    Dim rng As Range           ' O celula sau un grup de celule
    
    ' Referinte catre obiecte
    Set wb = ThisWorkbook              ' Cartea de lucru curenta
    Set ws = wb.Worksheets("Orders")   ' Foaia Orders din cartea curenta
    Set rng = ws.Range("A1:D1")       ' Intervalul A1:D1 din foaia Orders
    
    ' Afisarea unor proprietati
    Debug.Print "Nume carte: " & wb.Name
    Debug.Print "Nume foaie: " & ws.Name  
    Debug.Print "Adresa interval: " & rng.Address
    Debug.Print "Numar celule in interval: " & rng.Count
End Sub
```

---

## Manipularea celulelor și range-urilor

### Metode de referențiere a celulelor:

```vba
Sub MetodeReferentiere()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Metoda 1: Cells(rand, coloana) - indexul porneste de la 1
    Debug.Print "Celula B2: " & ws.Cells(2, 2).Value
    
    ' Metoda 2: Range cu adresa
    Debug.Print "Celula B2: " & ws.Range("B2").Value
    
    ' Metoda 3: Range cu interval
    Dim pretProduse As Range
    Set pretProduse = ws.Range("D2:D10")  ' Preturile primelor 9 produse
    Debug.Print "Suma preturilor: " & Application.WorksheetFunction.Sum(pretProduse)
    
    ' Metoda 4: Referinta relativa fata de o celula
    Dim celulaStart As Range
    Set celulaStart = ws.Range("A1")
    Debug.Print "Celula din dreapta lui A1: " & celulaStart.Offset(0, 1).Value
    Debug.Print "Celula de sub A1: " & celulaStart.Offset(1, 0).Value
End Sub
```

---

## Scrierea datelor în celule

```vba
Sub ScriereDateInCelule()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Gasirea primei linii goale pentru adaugarea unui produs nou
    Dim ultimaLinie As Integer
    ultimaLinie = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Adaugarea unui produs nou in urmatoarea linie libera
    ws.Cells(ultimaLinie, 1).Value = 100                    ' ProductID nou
    ws.Cells(ultimaLinie, 2).Value = "Produs Nou VBA"       ' ProductName
    ws.Cells(ultimaLinie, 3).Value = 1                      ' CategoryID
    ws.Cells(ultimaLinie, 4).Value = 29.99                  ' UnitPrice
    ws.Cells(ultimaLinie, 5).Value = 24.99                  ' UnitCost
    ws.Cells(ultimaLinie, 6).Value = "Beverages"            ' CategoryName
    
    ' Alternativ, putem scrie un intreg rand dintr-o data
    Dim dateProdusNou As Variant
    dateProdusNou = Array(101, "Alt Produs VBA", 2, 15.50, 12.40, "Condiments")
    ws.Range("A" & (ultimaLinie + 1) & ":F" & (ultimaLinie + 1)).Value = dateProdusNou
    
    MsgBox "Au fost adaugate 2 produse noi in foaia Products!"
End Sub
```

---

## Formatarea și stilizarea datelor

```vba
Sub FormatareDateVanzari()
    Dim ws As Worksheet
    Dim rngAntet As Range
    Dim rngDate As Range
    
    ' Cream o noua foaie pentru raportul nostru
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Raport Formatat"
    
    ' Crearea antetului
    Set rngAntet = ws.Range("A1:E1")
    rngAntet.Value = Array("Produs", "Categoria", "Pret Unitar", "Pret Total", "Status")
    
    ' Formatarea antetului
    With rngAntet
        .Font.Bold = True                    ' Text ingrosat
        .Font.Size = 12                      ' Dimensiunea fontului
        .Font.Color = RGB(255, 255, 255)     ' Text alb
        .Interior.Color = RGB(0, 100, 200)   ' Fundal albastru
        .HorizontalAlignment = xlCenter       ' Aliniere centru
        .Borders.LineStyle = xlContinuous    ' Borduri continue
    End With
    
    ' Adaugarea unor date de exemplu si formatarea lor
    Dim i As Integer
    Dim wsProducts As Worksheet
    Set wsProducts = ThisWorkbook.Worksheets("Products")
    
    For i = 2 To 6  ' Primele 5 produse
        ' Copierea datelor din foaia Products
        ws.Cells(i, 1).Value = wsProducts.Cells(i, 2).Value  ' ProductName
        ws.Cells(i, 2).Value = wsProducts.Cells(i, 6).Value  ' CategoryName
        ws.Cells(i, 3).Value = wsProducts.Cells(i, 4).Value  ' UnitPrice
        ws.Cells(i, 4).Value = wsProducts.Cells(i, 4).Value * 10  ' Pret total simulat
        
        ' Determinarea statusului in functie de pret
        If ws.Cells(i, 3).Value > 50 Then
            ws.Cells(i, 5).Value = "Premium"
            ws.Cells(i, 5).Interior.Color = RGB(255, 200, 200)  ' Fundal roz deschis
        ElseIf ws.Cells(i, 3).Value > 20 Then
            ws.Cells(i, 5).Value = "Standard"
            ws.Cells(i, 5).Interior.Color = RGB(255, 255, 200)  ' Fundal galben deschis
        Else
            ws.Cells(i, 5).Value = "Economic"
            ws.Cells(i, 5).Interior.Color = RGB(200, 255, 200)  ' Fundal verde deschis
        End If
        
        ' Formatarea preturilor ca valuta
        ws.Cells(i, 3).NumberFormat = "$#,##0.00"
        ws.Cells(i, 4).NumberFormat = "$#,##0.00"
    Next i
    
    ' Ajustarea automata a latimii coloanelor
    ws.Columns("A:E").AutoFit
    
    ' Adaugarea unei borduri pentru toate datele
    Set rngDate = ws.Range("A1:E6")
    rngDate.Borders.LineStyle = xlContinuous
    rngDate.Borders.Weight = xlThin
    
    MsgBox "Raportul formatat a fost creat cu succes!"
End Sub
```

---

## Lucrul cu mai multe foi simultan

```vba
Sub AnalizeaVanzariCompletaFoi()
    Dim wsOrders As Worksheet
    Dim wsProducts As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsRaport As Worksheet
    Dim i As Integer
    Dim ultimaLinie As Integer
    
    ' Setarea referintelor catre foile existente
    Set wsOrders = ThisWorkbook.Worksheets("Orders")
    Set wsProducts = ThisWorkbook.Worksheets("Products")  
    Set wsCustomers = ThisWorkbook.Worksheets("Customers")
    
    ' Crearea unei foi noi pentru raport
    Set wsRaport = ThisWorkbook.Worksheets.Add
    wsRaport.Name = "Analiza Completa"
    
    ' Crearea antetului pentru raportul detaliat
    wsRaport.Range("A1:F1").Value = Array("ID Comanda", "Client", "Produs", "Cantitate", "Pret Unitar", "Valoare Totala")
    
    ' Formatarea antetului
    With wsRaport.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ultimaLinie = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row
    Dim liniaCurenta As Integer
    liniaCurenta = 2
    
    ' Parcurgerea comenzilor si completarea raportului cu date din toate foile
    For i = 2 To 21  ' Analizăm primele 20 de comenzi pentru exemplu
        Dim idComanda As Integer
        Dim idClient As String
        Dim idProdus As Integer
        Dim cantitate As Integer
        Dim pretUnitar As Double
        Dim numeClient As String
        Dim numeProdus As String
        
        ' Citirea datelor din foaia Orders
        idComanda = wsOrders.Cells(i, 1).Value    ' OrderID
        idClient = wsOrders.Cells(i, 4).Value     ' CustomerID
        idProdus = wsOrders.Cells(i, 5).Value     ' ProductID
        cantitate = wsOrders.Cells(i, 7).Value    ' Quantity
        pretUnitar = wsOrders.Cells(i, 6).Value   ' UnitPrice
        
        ' Cautarea numelui clientului in foaia Customers
        numeClient = Application.VLookup(idClient, wsCustomers.Range("A:B"), 2, False)
        If IsError(numeClient) Then numeClient = "Client necunoscut"
        
        ' Cautarea numelui produsului in foaia Products
        numeProdus = Application.VLookup(idProdus, wsProducts.Range("A:B"), 2, False)
        If IsError(numeProdus) Then numeProdus = "Produs necunoscut"
        
        ' Completarea raportului
        wsRaport.Cells(liniaCurenta, 1).Value = idComanda
        wsRaport.Cells(liniaCurenta, 2).Value = numeClient
        wsRaport.Cells(liniaCurenta, 3).Value = numeProdus
        wsRaport.Cells(liniaCurenta, 4).Value = cantitate
        wsRaport.Cells(liniaCurenta, 5).Value = pretUnitar
        wsRaport.Cells(liniaCurenta, 6).Value = cantitate * pretUnitar  ' Valoarea totala
        
        ' Formatarea valorilor monetare
        wsRaport.Cells(liniaCurenta, 5).NumberFormat = "$#,##0.00"
        wsRaport.Cells(liniaCurenta, 6).NumberFormat = "$#,##0.00"
        
        liniaCurenta = liniaCurenta + 1
    Next i
    
    ' Ajustarea coloanelor si adaugarea bordurilor
    wsRaport.Columns("A:F").AutoFit
    wsRaport.Range("A1:F" & (liniaCurenta - 1)).Borders.LineStyle = xlContinuous
    
    MsgBox "Analiza completa a fost creata cu succes!"
End Sub
```

---

## Exercițiu practic Ora 3

Creați un raport de vânzări pe categorii care să:
1. Calculeze totalul vânzărilor pentru fiecare categorie de produse
2. Creeze un grafic pentru a vizualiza datele
3. Formateze frumos raportul

```vba
Sub RaportVanzariPerCategorie()
    Dim wsOrders As Worksheet
    Dim wsProducts As Worksheet
    Dim wsRaport As Worksheet
    Dim i As Integer
    Dim ultimaLinie As Integer
    Dim categorii As Object  ' Dictionary pentru stocarea totalurilor per categorie
    Dim idProdus As Integer
    Dim categorie As String
    Dim valoareComanda As Double
    
    ' Setarea referintelor
    Set wsOrders = ThisWorkbook.Worksheets("Orders")
    Set wsProducts = ThisWorkbook.Worksheets("Products")
    Set wsRaport = ThisWorkbook.Worksheets.Add
    wsRaport.Name = "Vanzari per Categorie"
    
    ' Crearea unui dictionary pentru categorii (simulam cu Collection)
    Set categorii = CreateObject("Scripting.Dictionary")
    
    ' Parcurgerea comenzilor pentru calcularea totalurilor
    ultimaLinie = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinie
        idProdus = wsOrders.Cells(i, 5).Value  ' ProductID
        valoareComanda = wsOrders.Cells(i, 6).Value * wsOrders.Cells(i, 7).Value  ' UnitPrice * Quantity
        
        ' Gasirea categoriei produsului
        categorie = Application.VLookup(idProdus, wsProducts.Range("A:F"), 6, False)
        If Not IsError(categorie) Then
            ' Adaugarea la totalul categoriei
            If categorii.Exists(categorie) Then
                categorii(categorie) = categorii(categorie) + valoareComanda
            Else
                categorii.Add categorie, valoareComanda
            End If
        End If
    Next i
    
    ' Crearea raportului
    wsRaport.Range("A1:B1").Value = Array("Categorie", "Total Vanzari")
    
    ' Formatarea antetului
    With wsRaport.Range("A1:B1")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(0, 150, 0)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Completarea datelor
    Dim chei As Variant
    chei = categorii.Keys
    
    For i = 0 To categorii.Count - 1
        wsRaport.Cells(i + 2, 1).Value = chei(i)
        wsRaport.Cells(i + 2, 2).Value = categorii(chei(i))
        wsRaport.Cells(i + 2, 2).NumberFormat = "$#,##0.00"
        
        ' Colorarea alternativa a randurilor
        If i Mod 2 = 0 Then
            wsRaport.Range("A" & (i + 2) & ":B" & (i + 2)).Interior.Color = RGB(240, 240, 240)
        End If
    Next i
    
    ' Ajustarea coloanelor
    wsRaport.Columns("A:B").AutoFit
    
    ' Adaugarea bordurilor
    wsRaport.Range("A1:B" & (categorii.Count + 1)).Borders.LineStyle = xlContinuous
    
    ' Crearea unui grafic
    Dim grafic As Chart
    Set grafic = wsRaport.Shapes.AddChart2(, xlColumnClustered).Chart
    grafic.SetSourceData wsRaport.Range("A1:B" & (categorii.Count + 1))
    grafic.HasTitle = True
    grafic.ChartTitle.Text = "Vanzari per Categorie de Produse"
    
    MsgBox "Raportul cu grafic a fost creat cu succes!"
End Sub
```

---

# ORA 4: APLICAȚII AVANSATE

## Funcții definite de utilizator (UDF)

Putem crea propriile funcții care să poată fi folosite în foi de calcul ca orice altă funcție Excel:

```vba
' Functie pentru calcularea profitului pe baza pretiului si costului
Function CalculProfit(pretVanzare As Double, costProdus As Double) As Double
    ' Calcularea profitului ca diferenta intre pret si cost
    CalculProfit = pretVanzare - costProdus
End Function

' Functie pentru calcularea marjei de profit in procente
Function MarjaProfit(pretVanzare As Double, costProdus As Double) As Double
    ' Verificarea impartirii la zero
    If costProdus = 0 Then
        MarjaProfit = 0
    Else
        ' Calcularea marjei: (Pret - Cost) / Cost * 100
        MarjaProfit = ((pretVanzare - costProdus) / costProdus) * 100
    End If
End Function

' Functie pentru clasificarea performantei angajatilor
Function ClasificarePerformanta(numarComenzi As Integer) As String
    ' Clasificarea pe baza numarului de comenzi procesate
    Select Case numarComenzi
        Case Is < 30
            ClasificarePerformanta = "Sub medie"
        Case 30 To 60
            ClasificarePerformanta = "Medie"
        Case 61 To 100
            ClasificarePerformanta = "Buna"
        Case Is > 100
            ClasificarePerformanta = "Excelenta"
        Case Else
            ClasificarePerformanta = "Date insuficiente"
    End Select
End Function

' Exemplu de folosire a functiilor definite mai sus
Sub TestFunctiiPersonalizate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Testarea functiilor cu date din primul produs
    Dim pret As Double
    Dim cost As Double
    
    pret = ws.Cells(2, 4).Value  ' UnitPrice primul produs
    cost = ws.Cells(2, 5).Value  ' UnitCost primul produs
    
    Debug.Print "Pret: " & pret
    Debug.Print "Cost: " & cost
    Debug.Print "Profit: " & CalculProfit(pret, cost)
    Debug.Print "Marja profit: " & MarjaProfit(pret, cost) & "%"
End Sub
```

---

## Gestionarea erorilor

Gestionarea erorilor este esențială pentru crearea de aplicații robuste:

```vba
Sub CitireDateCuGestionareErori()
    ' Activarea gestionarii erorilor
    On Error GoTo GestionareEroare
    
    Dim ws As Worksheet
    Dim numeClient As String
    Dim pretProdus As Double
    Dim dataComanda As Date
    
    ' Incercam sa accesam o foaie care ar putea sa nu existe
    Set ws = ThisWorkbook.Worksheets("FoaieInexistenta")
    
    ' Acest cod nu va fi executat daca foaia nu exista
    numeClient = ws.Cells(1, 1).Value
    
    MsgBox "Datele au fost citite cu succes: " & numeClient
    
    ' Iesirea normala din procedura
    Exit Sub
    
GestionareEroare:
    ' Codul care se executa cand apare o eroare
    Select Case Err.Number
        Case 9  ' Subscript out of range (foaia nu exista)
            MsgBox "Eroare: Foaia specificata nu exista!" & vbCrLf & _
                   "Eroare nr: " & Err.Number & vbCrLf & _
                   "Descriere: " & Err.Description
                   
            ' Incercam sa folosim o foaie care exista
            Set ws = ThisWorkbook.Worksheets("Orders")
            MsgBox "Am comutat la foaia Orders ca alternativa"
            Resume Next  ' Continua cu urmatoarea linie dupa cea care a cauzat eroarea
            
        Case 1004  ' Application-defined or object-defined error
            MsgBox "Eroare de aplicatie: " & Err.Description
            
        Case Else
            MsgBox "Eroare neasteptata:" & vbCrLf & _
                   "Numar: " & Err.Number & vbCrLf & _
                   "Descriere: " & Err.Description
    End Select
    
    ' Resetarea sistemului de gestionare a erorilor
    Err.Clear
End Sub
```

---

## Lucrul cu evenimente

Evenimentele ne permit să executăm cod automat când se întâmplă anumite acțiuni:

```vba
' Acest cod trebuie pus in modulul foii (nu in modulul standard)
' Faceti dublu-click pe foaia dorita in Project Explorer pentru a accesa modulul ei

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Acest cod se executa automat cand se modifica o celula in foaie
    
    ' Verificam daca modificarea a fost in coloana preturilor (coloana D in Products)
    If Not Intersect(Target, Me.Columns("D:D")) Is Nothing Then
        ' Calculam automat marja de profit
        Dim linia As Integer
        linia = Target.Row
        
        ' Verificam daca avem si costul produsului
        If Me.Cells(linia, 5).Value <> "" Then
            Dim pret As Double
            Dim cost As Double
            pret = Me.Cells(linia, 4).Value
            cost = Me.Cells(linia, 5).Value
            
            ' Calculam si afisam marja in coloana urmatoare
            If cost > 0 Then
                Me.Cells(linia, 7).Value = ((pret - cost) / cost) * 100
                Me.Cells(linia, 7).NumberFormat = "0.00%"
            End If
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Acest cod se executa cand se schimba selectia
    ' Putem sa evidentiem randul si coloana curente
    
    ' Resetam colorarea anterioara
    Cells.Interior.ColorIndex = xlNone
    
    ' Evidentiem randul si coloana curente
    Target.EntireRow.Interior.Color = RGB(255, 255, 200)
    Target.EntireColumn.Interior.Color = RGB(200, 255, 200)
End Sub
```

---

## Proiect final: Analizator complet de vânzări

```vba
Sub AnalizatorComplectVanzari()
    ' Acest macro combina toate conceptele invatate pentru a crea
    ' un analizator complet de vanzari
    
    On Error GoTo GestionareEroare  ' Activarea gestionarii erorilor
    
    Dim wsOrders As Worksheet
    Dim wsProducts As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsRaport As Worksheet
    
    ' Setarea referintelor
    Set wsOrders = ThisWorkbook.Worksheets("Orders")
    Set wsProducts = ThisWorkbook.Worksheets("Products")
    Set wsCustomers = ThisWorkbook.Worksheets("Customers")
    
    ' Crearea foii de raport
    On Error Resume Next  ' Ignora eroarea daca foaia exista deja
    Set wsRaport = ThisWorkbook.Worksheets("Dashboard Vanzari")
    If wsRaport Is Nothing Then
        Set wsRaport = ThisWorkbook.Worksheets.Add
        wsRaport.Name = "Dashboard Vanzari"
    Else
        wsRaport.Cells.Clear  ' Sterge continutul anterior
    End If
    On Error GoTo GestionareEroare  ' Reactiveaza gestionarea completa a erorilor
    
    ' Crearea antetului principal
    wsRaport.Range("A1:F1").Merge
    wsRaport.Range("A1").Value = "DASHBOARD ANALIZA VANZARI NORTHWIND"
    With wsRaport.Range("A1")
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 100, 200)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Sectiunea 1: Statistici generale
    wsRaport.Range("A3").Value = "STATISTICI GENERALE"
    wsRaport.Range("A3").Font.Bold = True
    
    Dim totalComenzi As Integer
    Dim totalClienti As Integer
    Dim totalProduse As Integer
    
    totalComenzi = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row - 1
    totalClienti = wsCustomers.Cells(wsCustomers.Rows.Count, 1).End(xlUp).Row - 1
    totalProduse = wsProducts.Cells(wsProducts.Rows.Count, 1).End(xlUp).Row - 1
    
    wsRaport.Range("A4").Value = "Total comenzi: " & totalComenzi
    wsRaport.Range("A5").Value = "Total clienti: " & totalClienti
    wsRaport.Range("A6").Value = "Total produse: " & totalProduse
    
    ' Sectiunea 2: Top 5 produse dupa valoare vanzari
    wsRaport.Range("A8").Value = "TOP 5 PRODUSE - VALOARE VANZARI"
    wsRaport.Range("A8").Font.Bold = True
    
    wsRaport.Range("A9:C9").Value = Array("Produs", "Categoria", "Valoare Totala")
    wsRaport.Range("A9:C9").Font.Bold = True
    wsRaport.Range("A9:C9").Interior.Color = RGB(200, 200, 200)
    
    ' Calcularea vanzarilor pe produse (versiune simplificata)
    Dim produsVanzari As Object
    Set produsVanzari = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim ultimaLinieOrders As Integer
    ultimaLinieOrders = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row
    
    ' Calcularea valorii totale pentru fiecare produs
    For i = 2 To ultimaLinieOrders
        Dim idProdus As Integer
        Dim valoare As Double
        
        idProdus = wsOrders.Cells(i, 5).Value  ' ProductID
        valoare = wsOrders.Cells(i, 6).Value * wsOrders.Cells(i, 7).Value  ' UnitPrice * Quantity
        
        If produsVanzari.Exists(idProdus) Then
            produsVanzari(idProdus) = produsVanzari(idProdus) + valoare
        Else
            produsVanzari.Add idProdus, valoare
        End If
    Next i
    
    ' Gasirea si afisarea top 5 produse (implementare simplificata)
    Dim linieCurenta As Integer
    linieCurenta = 10
    Dim contor As Integer
    contor = 0
    
    Dim chei As Variant
    chei = produsVanzari.Keys
    
    ' Sortarea simpla (bubble sort pentru primele 5)
    Dim j As Integer, temp As Variant, tempValoare As Double
    For i = 0 To UBound(chei)
        For j = i + 1 To UBound(chei)
            If produsVanzari(chei(i)) < produsVanzari(chei(j)) Then
                temp = chei(i)
                chei(i) = chei(j)
                chei(j) = temp
            End If
        Next j
        
        ' Afisarea primelor 5
        If i < 5 Then
            Dim numeProdus As String
            Dim categorieProdus As String
            
            numeProdus = Application.VLookup(chei(i), wsProducts.Range("A:B"), 2, False)
            categorieProdus = Application.VLookup(chei(i), wsProducts.Range("A:F"), 6, False)
            
            If Not IsError(numeProdus) And Not IsError(categorieProdus) Then
                wsRaport.Cells(linieCurenta, 1).Value = numeProdus
                wsRaport.Cells(linieCurenta, 2).Value = categorieProdus
                wsRaport.Cells(linieCurenta, 3).Value = produsVanzari(chei(i))
                wsRaport.Cells(linieCurenta, 3).NumberFormat = "$#,##0.00"
                linieCurenta = linieCurenta + 1
            End If
        End If
    Next i
    
    ' Sectiunea 3: Analiza pe tari
    wsRaport.Range("E3").Value = "ANALIZA PE TARI - TOP 3"
    wsRaport.Range("E3").Font.Bold = True
    
    wsRaport.Range("E4:F4").Value = Array("Tara", "Numar Clienti")
    wsRaport.Range("E4:F4").Font.Bold = True
    wsRaport.Range("E4:F4").Interior.Color = RGB(200, 200, 200)
    
    ' Numararea clientilor pe tari
    Dim tariClienti As Object
    Set tariClienti = CreateObject("Scripting.Dictionary")
    
    Dim ultimaLinieCustomers As Integer
    ultimaLinieCustomers = wsCustomers.Cells(wsCustomers.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinieCustomers
        Dim tara As String
        tara = wsCustomers.Cells(i, 4).Value  ' Country
        
        If tariClienti.Exists(tara) Then
            tariClienti(tara) = tariClienti(tara) + 1
        Else
            tariClienti.Add tara, 1
        End If
    Next i
    
    ' Afisarea primelor 3 tari
    Dim cheiTari As Variant
    cheiTari = tariClienti.Keys
    
    ' Sortare simpla pentru primele 3
    For i = 0 To UBound(cheiTari)
        For j = i + 1 To UBound(cheiTari)
            If tariClienti(cheiTari(i)) < tariClienti(cheiTari(j)) Then
                temp = cheiTari(i)
                cheiTari(i) = cheiTari(j)
                cheiTari(j) = temp
            End If
        Next j
    Next i
    
    For i = 0 To 2  ' Primele 3 tari
        If i <= UBound(cheiTari) Then
            wsRaport.Cells(5 + i, 5).Value = cheiTari(i)
            wsRaport.Cells(5 + i, 6).Value = tariClienti(cheiTari(i))
        End If
    Next i
    
    ' Formatarea finala
    wsRaport.Columns("A:F").AutoFit
    wsRaport.Range("A9:C14").Borders.LineStyle = xlContinuous
    wsRaport.Range("E4:F7").Borders.LineStyle = xlContinuous
    
    ' Adaugarea datei si orei generate
    wsRaport.Range("A16").Value = "Raport generat la: " & Now()
    wsRaport.Range("A16").Font.Italic = True
    
    MsgBox "Dashboard-ul complet de analiza a vanzarilor a fost generat cu succes!" & vbCrLf & _
           "Verificati foaia 'Dashboard Vanzari' pentru rezultate."
    
    ' Activarea foii de raport
    wsRaport.Activate
    
    Exit Sub
    
GestionareEroare:
    MsgBox "A aparut o eroare in procesarea datelor:" & vbCrLf & _
           "Eroare: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Va rugam verificati datele si incercati din nou."
    Err.Clear
End Sub
```

---

## Exercițiu final și recapitulare

### Exercițiu de consolidare:

Creați un sistem complet de gestionare care să:
1. Permită adăugarea de noi produse cu validare
2. Calculeze automat profitul și marja 
3. Creeze rapoarte personalizate
4. Gestioneze erorile elegant

```vba
Sub SistemGestionareComplet()
    ' Combinarea tuturor conceptelor intr-un sistem integrat
    
    Dim raspuns As String
    raspuns = InputBox("Ce operatiune doriti sa efectuati?" & vbCrLf & _
                      "1 - Adauga produs nou" & vbCrLf & _
                      "2 - Genereaza raport vanzari" & vbCrLf & _
                      "3 - Calculeaza profitabilitatea" & vbCrLf & _
                      "Introduceti numarul operatiunii:")
    
    Select Case raspuns
        Case "1"
            Call AdaugaProdusNou
        Case "2" 
            Call GenereazaRaportVanzari
        Case "3"
            Call CalculeazaProfitabilitatea
        Case Else
            MsgBox "Optiune invalida! Va rugam selectati 1, 2 sau 3."
    End Select
End Sub

Sub AdaugaProdusNou()
    On Error GoTo GestionareEroare
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Products")
    
    ' Colectarea informatiilor despre produsul nou
    Dim numeProdus As String
    Dim pret As Double
    Dim cost As Double
    Dim categorie As String
    
    numeProdus = InputBox("Introduceti numele produsului nou:")
    If numeProdus = "" Then Exit Sub
    
    pret = CDbl(InputBox("Introduceti pretul produsului:"))
    If pret <= 0 Then
        MsgBox "Pretul trebuie sa fie pozitiv!"
        Exit Sub
    End If
    
    cost = CDbl(InputBox("Introduceti costul produsului:"))
    If cost <= 0 Then
        MsgBox "Costul trebuie sa fie pozitiv!"
        Exit Sub
    End If
    
    If cost >= pret Then
        MsgBox "Atentie: Costul este mai mare sau egal cu pretul! Produsul nu va fi profitabil."
    End If
    
    categorie = InputBox("Introduceti categoria produsului:")
    If categorie = "" Then categorie = "Necategorisit"
    
    ' Adaugarea produsului in foaie
    Dim ultimaLinie As Integer
    ultimaLinie = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    Dim idNou As Integer
    idNou = ws.Cells(ultimaLinie - 1, 1).Value + 1  ' ID automat
    
    ws.Cells(ultimaLinie, 1).Value = idNou
    ws.Cells(ultimaLinie, 2).Value = numeProdus
    ws.Cells(ultimaLinie, 3).Value = 1  ' CategoryID implicit
    ws.Cells(ultimaLinie, 4).Value = pret
    ws.Cells(ultimaLinie, 5).Value = cost
    ws.Cells(ultimaLinie, 6).Value = categorie
    
    MsgBox "Produsul '" & numeProdus & "' a fost adaugat cu succes!" & vbCrLf & _
           "ID produs: " & idNou & vbCrLf & _
           "Profit estimat: " & Format(pret - cost, "Currency")
    
    Exit Sub
    
GestionareEroare:
    MsgBox "Eroare la adaugarea produsului: " & Err.Description
    Err.Clear
End Sub
```

---

## Recapitulare și resurse suplimentare

### Ce am învățat în aceste 4 ore:

**Ora 1 - Fundamentele:**
- Editorul VBA și macro-urile
- Variabile și tipuri de date
- Lucrul cu obiecte Excel

**Ora 2 - Structuri de control:**
- Instrucțiuni If-Then-Else
- Bucle For, While, Do
- Select Case pentru decizii complexe

**Ora 3 - Manipularea datelor:**
- Range-uri și celule
- Formatarea automată
- Lucrul cu mai multe foi

**Ora 4 - Aplicații avansate:**
- Funcții personalizate (UDF)
- Gestionarea erorilor
- Evenimente și automatizare

### Concepte cheie de reținut:

1. **Întotdeauna declarați variabilele** cu `Dim`
2. **Folosiți `Set` pentru obiecte** și `=` pentru valori
3. **Gestionați erorile** cu `On Error GoTo`
4. **Comentați codul** pentru înțelegere ulterioară
5. **Testați pas cu pas** folosind `Debug.Print` și `F8`

### Pentru dezvoltare ulterioară:

- Studiați evenimentele Workbook și Application
- Învățați să creați formulare (UserForms)
- Explorați conectarea la baze de date externe
- Practicați automatizarea altor aplicații Office
- Cercetați despre clase și module avansate

### Resurse utile:
- Microsoft VBA Documentation
- Excel VBA Programming For Dummies
- Comunități online: Stack Overflow, Reddit r/excel
- Canale YouTube specializate în Excel VBA

---

**Mulțumim pentru participarea la acest curs introductiv de Excel VBA!**

**Succes în automatizarea sarcinilor voastre!**


## Cod in curs

### Exemplu

```vba
Sub TipuriDeDate()
    ' Declararea variabilelor cu tipuri specifice
    Dim numeClient As String        ' Text (siruri de caractere)
    Dim varsta As Integer          ' Numere intregi (-32,768 la 32,767)
    Dim pret As Double            ' Numere cu zecimale
    Dim esteActiv As Boolean      ' Valori True/False
    Dim dataComanda As Date       ' Date si ore
    
    ' Atribuirea valorilor
    numeClient = "Alfreds Futterkiste"  ' Folosim ghilimele pentru text
    varsta = 25
    pret = 123.45
    esteActiv = True
    dataComanda = #12/31/2024#    ' Datele se incadreaza in #
    
    ' Afisarea valorilor in ferestra Immediate (Ctrl+G)
    Debug.Print "Client: " & numeClient
    Debug.Print "Varsta: " & varsta
    Debug.Print "Pret: " & pret
    Debug.Print "Este activ: " & esteActiv
    Debug.Print "Data comenzii: " & dataComanda
End Sub
```

---
