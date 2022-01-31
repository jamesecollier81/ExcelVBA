Sub CopyColumns()

    Dim fname As Range
    Dim lname As Range
    Dim email As Range
    Dim mphone As Range
    Dim hphone As Range
    Dim wphone As Range
    Dim street As Range
    Dim street2 As Range
    Dim city As Range
    Dim state As Range
    Dim zip As Range
    Dim country As Range
    Dim company As Range
    Dim co_size As Range
    Dim ind As Range
    Dim li As Range
    Dim domain As Range
    
    Dim customer As String
    
    Dim lrow As Long
    Dim i As Long
    
    'get customer instance name
    customer = Sheets("facts").Cells(1, 2).Value
    
    'figure out the last row of exported data
    lrow = ActiveWorkbook.Worksheets("sheet2").Cells(Rows.Count, "B").End(xlUp).Row
    
    'clear sheet for new run
    Sheets("sheet2").Cells.Clear
    
    'for some reason it wasn't working to find `id` so I am just copying the first column since that is where it is
    Sheets("data").Columns(1).Copy Destination:=Sheets("sheet2").Columns(2)
    
    'copy only necessary columns for output
    Set fname = Sheets("data").Rows(1).Find("First Name")
    fname.EntireColumn.Copy Sheets("sheet2").Range("C1")
    
    Set lname = Sheets("data").Rows(1).Find("Last Name")
    lname.EntireColumn.Copy Sheets("sheet2").Range("D1")
    
    Set email = Sheets("data").Rows(1).Find("Email")
    email.EntireColumn.Copy Sheets("sheet2").Range("E1")
    
    Set mphone = Sheets("data").Rows(1).Find("Mobile Phone")
    mphone.EntireColumn.Copy Sheets("sheet2").Range("F1")
    
    Set hphone = Sheets("data").Rows(1).Find("Home Phone")
    hphone.EntireColumn.Copy Sheets("sheet2").Range("G1")
    
    Set wphone = Sheets("data").Rows(1).Find("Work Phone")
    wphone.EntireColumn.Copy Sheets("sheet2").Range("H1")

    Set street = Sheets("data").Rows(1).Find("Street")
    street.EntireColumn.Copy Sheets("sheet2").Range("I1")
    
    Set street2 = Sheets("data").Rows(1).Find("Street 2")
    street2.EntireColumn.Copy Sheets("sheet2").Range("J1")
    
    Set city = Sheets("data").Rows(1).Find("City")
    city.EntireColumn.Copy Sheets("sheet2").Range("K1")
    
    Set state = Sheets("data").Rows(1).Find("State")
    state.EntireColumn.Copy Sheets("sheet2").Range("L1")

    Set zip = Sheets("data").Rows(1).Find("Zipcode")
    zip.EntireColumn.Copy Sheets("sheet2").Range("M1")

    Set country = Sheets("data").Rows(1).Find("Country")
    country.EntireColumn.Copy Sheets("sheet2").Range("N1")
    
    Set company = Sheets("data").Rows(1).Find("Company")
    company.EntireColumn.Copy Sheets("sheet2").Range("O1")
    
    Set co_size = Sheets("data").Rows(1).Find("Company Size")
    co_size.EntireColumn.Copy Sheets("sheet2").Range("P1")

    Set ind = Sheets("data").Rows(1).Find("Company Industry")
    ind.EntireColumn.Copy Sheets("sheet2").Range("Q1")
    
    Set li = Sheets("data").Rows(1).Find("LinkedIn")
    li.EntireColumn.Copy Sheets("sheet2").Range("R1")
    
    Set domain = Sheets("data").Rows(1).Find("Website")
    domain.EntireColumn.Copy Sheets("sheet2").Range("S1")

    'assign column names for other data points
    Sheets("sheet2").Range("A1") = "panoply_id"
    Sheets("sheet2").Range("T1") = "customer_crm"
    Sheets("sheet2").Range("U1") = "customer_revenue"
    Sheets("sheet2").Range("V1") = "purchasing_dept"
    
    'for every row, concatenate the outreach id and the customer instance to generate a unique identifier
    For i = 2 To lrow
        Sheets("sheet2").Range("A" & i) = customer + CStr(Cells(i, 2).Value)
    Next i
End Sub
