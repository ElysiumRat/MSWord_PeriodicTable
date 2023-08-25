Attribute VB_Name = "WhatElement"
Sub WhatElement()

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
    givenelement = InputBox("Please input element symbol:")
    counter = 1
    Else
        givenelement = selrange.Text
        counter = 0
End If

' Creating a dictionary of all elements in the table, with the symbols as the keys
With PeriodicTable
    .Add "H", "Hydrogen"
    .Add "He", "Helium"
    .Add "Li", "Lithium"
    .Add "Be", "Beryllium"
    .Add "B", "Boron"
    .Add "C", "Carbon"
    .Add "N", "Nitrogen"
    .Add "O", "Oxygen"
    .Add "F", "Fluorine"
    .Add "Ne", "Neon"
    .Add "Na", "Sodium"
    .Add "Mg", "Magnesium"
    .Add "Al", "Aluminium"
    .Add "Si", "Silicon"
    .Add "P", "Phosphorus"
    .Add "S", "Sulfur"
    .Add "Cl", "Chlorine"
    .Add "Ar", "Argon"
    .Add "K", "Potassium"
    .Add "Ca", "Calcium"
    .Add "Sc", "Scandium"
    .Add "Ti", "Titanium"
    .Add "V", "Vanadium"
    .Add "Cr", "Chromium"
    .Add "Mn", "Manganese"
    .Add "Fe", "Iron"
    .Add "Co", "Cobalt"
    .Add "Ni", "Nickel"
    .Add "Cu", "Copper"
    .Add "Zn", "Zinc"
    .Add "Ga", "Gallium"
    .Add "Ge", "Germanium"
    .Add "As", "Arsenic"
    .Add "Se", "Selenium"
    .Add "Br", "Bromine"
    .Add "Kr", "Krypton"
    .Add "Rb", "Rubidium"
    .Add "Sr", "Strontium"
    .Add "Y", "Yttrium"
    .Add "Zr", "Zirconium"
    .Add "Nb", "Niobium"
    .Add "Mo", "Molybdenum"
    .Add "Tc", "Technetium"
    .Add "Ru", "Ruthenium"
    .Add "Rh", "Rhodium"
    .Add "Pd", "Palladium"
    .Add "Ag", "Silver"
    .Add "Cd", "Cadmium"
    .Add "In", "Indium"
    .Add "Sn", "Tin"
    .Add "Sb", "Antimony"
    .Add "Te", "Tellurium"
    .Add "I", "Iodine"
    .Add "Xe", "Xenon"
    .Add "Cs", "Caesium"
    .Add "Ba", "Barium"
    .Add "La", "Lanthanum"
    .Add "Ce", "Cerium"
    .Add "Pr", "Praseodymium"
    .Add "Nd", "Neodymium"
    .Add "Pm", "Promethium"
    .Add "Sm", "Samarium"
    .Add "Eu", "Europium"
    .Add "Gd", "Gadolinium"
    .Add "Tb", "Terbium"
    .Add "Dy", "Dysprosium"
    .Add "Ho", "Holmium"
    .Add "Er", "Erbium"
    .Add "Tm", "Thulium"
    .Add "Yb", "Ytterbium"
    .Add "Lu", "Lutetium"
    .Add "Hf", "Hafnium"
    .Add "Ta", "Tantalum"
    .Add "W", "Tungsten"
    .Add "Re", "Rhenium"
    .Add "Os", "Osmium"
    .Add "Ir", "Iridium"
    .Add "Pt", "Platinum"
    .Add "Au", "Gold"
    .Add "Hg", "Mercury"
    .Add "Tl", "Thallium"
    .Add "Pb", "Lead"
    .Add "Bi", "Bismuth"
    .Add "Po", "Polonium"
    .Add "At", "Astatine"
    .Add "Rn", "Radon"
    .Add "Fr", "Francium"
    .Add "Ra", "Radium"
    .Add "Ac", "Actinium"
    .Add "Th", "Thorium"
    .Add "Pa", "Protactinium"
    .Add "U", "Uranium"
    .Add "Np", "Neptunium"
    .Add "Pu", "Plutonium"
    .Add "Am", "Americium"
    .Add "Cm", "Curium"
    .Add "Bk", "Berkelium"
    .Add "Cf", "Californium"
    .Add "Es", "Einsteinium"
    .Add "Fm", "Fermium"
    .Add "Md", "Mendelevium"
    .Add "No", "Nobelium"
    .Add "Lr", "Lawrencium"
    .Add "Rf", "Rutherfordium"
    .Add "Db", "Dubnium"
    .Add "Sg", "Seaborgium"
    .Add "Bh", "Bohrium"
    .Add "Hs", "Hassium"
    .Add "Mt", "Meitnerium"
    .Add "Ds", "Darmstadtium"
    .Add "Rg", "Roentgenium"
    .Add "Cn", "Copernicium"
    .Add "Nh", "Nihonium"
    .Add "Fl", "Flerovium"
    .Add "Mc", "Moscovium"
    .Add "Lv", "Livermorium"
    .Add "Ts", "Tennessine"
    .Add "Og", "Oganesson"
End With

Check:
If PeriodicTable.Exists(givenelement) Then
    MsgBox "Element symbol: " & givenelement & vbCrLf & "Full name: " & PeriodicTable.Item(givenelement)
    Exit Sub
    Else
        If counter > 0 Then
            MsgBox "Sorry, element not found."
        Else
            counter = counter + 1
            givenelement = InputBox("Please input element symbol:")
            GoTo Check
        End If
End If

End Sub

