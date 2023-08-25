Attribute VB_Name = "WhatChemSymbol"
Sub WhatChemSymbol()

Dim PeriodicTable
Dim givenelement As String
Dim counter As Integer
Dim selrange As Range
Dim selStart, selEnd As Long

Set PeriodicTable = CreateObject("Scripting.Dictionary")

Set selrange = ActiveDocument.Range
selStart = selection.Start
selEnd = selection.End
selrange.SetRange Start:=selStart, End:=selEnd
selrange.Select

If selection.Range.ComputeStatistics(wdStatisticWords) < 1 Then
    givenelement = InputBox("Please input element name:")
    counter = 1
    Else
        givenelement = selrange.Text
        counter = 0
End If

' Creating a dictionary of all elements in the table, with the element names as the keys
With PeriodicTable
    .Add "Hydrogen", "H"
    .Add "Helium", "He"
    .Add "Lithium", "Li"
    .Add "Beryllium", "Be"
    .Add "Boron", "B"
    .Add "Carbon", "C"
    .Add "Nitrogen", "N"
    .Add "Oxygen", "O"
    .Add "Fluorine", "F"
    .Add "Neon", "Ne"
    .Add "Sodium", "Na"
    .Add "Magnesium", "Mg"
    .Add "Aluminium", "Al"
    .Add "Silicon", "Si"
    .Add "Phosphorus", "P"
    .Add "Sulfur", "S"
    .Add "Chlorine", "Cl"
    .Add "Argon", "Ar"
    .Add "Potassium", "K"
    .Add "Calcium", "Ca"
    .Add "Scandium", "Sc"
    .Add "Titanium", "Ti"
    .Add "Vanadium", "V"
    .Add "Chromium", "Cr"
    .Add "Manganese", "Mn"
    .Add "Iron", "Fe"
    .Add "Cobalt", "Co"
    .Add "Nickel", "Ni"
    .Add "Copper", "Cu"
    .Add "Zinc", "Zn"
    .Add "Gallium", "Ga"
    .Add "Germanium", "Ge"
    .Add "Arsenic", "As"
    .Add "Selenium", "Se"
    .Add "Bromine", "Br"
    .Add "Krypton", "Kr"
    .Add "Rubidium", "Rb"
    .Add "Strontium", "Sr"
    .Add "Yttrium", "Y"
    .Add "Zirconium", "Zr"
    .Add "Niobium", "Nb"
    .Add "Molybdenum", "Mo"
    .Add "Technetium", "Tc"
    .Add "Ruthenium", "Ru"
    .Add "Rhodium", "Rh"
    .Add "Palladium", "Pd"
    .Add "Silver", "Ag"
    .Add "Cadmium", "Cd"
    .Add "Indium", "In"
    .Add "Tin", "Sn"
    .Add "Antimony", "Sb"
    .Add "Tellurium", "Te"
    .Add "Iodine", "I"
    .Add "Xenon", "Xe"
    .Add "Caesium", "Cs"
    .Add "Barium", "Ba"
    .Add "Lanthanum", "La"
    .Add "Cerium", "Ce"
    .Add "Praseodymium", "Pr"
    .Add "Neodymium", "Nd"
    .Add "Promethium", "Pm"
    .Add "Samarium", "Sm"
    .Add "Europium", "Eu"
    .Add "Gadolinium", "Gd"
    .Add "Terbium", "Tb"
    .Add "Dysprosium", "Dy"
    .Add "Holmium", "Ho"
    .Add "Erbium", "Er"
    .Add "Thulium", "Tm"
    .Add "Ytterbium", "Yb"
    .Add "Lutetium", "Lu"
    .Add "Hafnium", "Hf"
    .Add "Tantalum", "Ta"
    .Add "Tungsten", "W"
    .Add "Rhenium", "Re"
    .Add "Osmium", "Os"
    .Add "Iridium", "Ir"
    .Add "Platinum", "Pt"
    .Add "Gold", "Au"
    .Add "Mercury", "Hg"
    .Add "Thallium", "Tl"
    .Add "Lead", "Pb"
    .Add "Bismuth", "Bi"
    .Add "Polonium", "Po"
    .Add "Astatine", "At"
    .Add "Radon", "Rn"
    .Add "Francium", "Fr"
    .Add "Radium", "Ra"
    .Add "Actinium", "Ac"
    .Add "Thorium", "Th"
    .Add "Protactinium", "Pa"
    .Add "Uranium", "U"
    .Add "Neptunium", "Np"
    .Add "Plutonium", "Pu"
    .Add "Americium", "Am"
    .Add "Curium", "Cm"
    .Add "Berkelium", "Bk"
    .Add "Californium", "Cf"
    .Add "Einsteinium", "Es"
    .Add "Fermium", "Fm"
    .Add "Mendelevium", "Md"
    .Add "Nobelium", "No"
    .Add "Lawrencium", "Lr"
    .Add "Rutherfordium", "Rf"
    .Add "Dubnium", "Db"
    .Add "Seaborgium", "Sg"
    .Add "Bohrium", "Bh"
    .Add "Hassium", "Hs"
    .Add "Meitnerium", "Mt"
    .Add "Darmstadtium", "Ds"
    .Add "Roentgenium", "Rg"
    .Add "Copernicium", "Cn"
    .Add "Nihonium", "Nh"
    .Add "Flerovium", "Fl"
    .Add "Moscovium", "Mc"
    .Add "Livermorium", "Lv"
    .Add "Tennessine", "Ts"
    .Add "Oganesson", "Og"
End With

Check:
If PeriodicTable.Exists(givenelement) Then
    MsgBox "Element symbol: " & PeriodicTable.Item(givenelement) & vbCrLf & "Full name: " & givenelement
    Exit Sub
    Else
        If counter > 0 Then
            MsgBox "Sorry, element not found."
        Else
            counter = counter + 1
            givenelement = InputBox("Please input element name:")
            GoTo Check
        End If
End If

End Sub

