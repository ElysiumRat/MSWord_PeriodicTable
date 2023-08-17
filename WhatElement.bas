Attribute VB_Name = "WhatElement"
Sub WhatElement()

Dim PeriodicTable
Dim seltext As String
Dim counter As Integer
Dim selrange As Range
Dim selStart, selEnd As Long

Set PeriodicTable = CreateObject("Scripting.Dictionary")

Set selrange = ActiveDocument.Range
selStart = selection.Start
selEnd = selection.End
selrange.SetRange Start:=selStart, End:=selEnd
selrange.Select

seltext = selrange.Text

If selection.Range.ComputeStatistics(wdStatisticWords) < 1 Then
    MsgBox "Nothing selected."
    Exit Sub
End If

' Creating a dictionary of all elements in the table, with the symbols as the keys
PeriodicTable.Add "H", "Hydrogen"
PeriodicTable.Add "He", "Helium"
PeriodicTable.Add "Li", "Lithium"
PeriodicTable.Add "Be", "Beryllium"
PeriodicTable.Add "B", "Boron"
PeriodicTable.Add "C", "Carbon"
PeriodicTable.Add "N", "Nitrogen"
PeriodicTable.Add "O", "Oxygen"
PeriodicTable.Add "F", "Fluorine"
PeriodicTable.Add "Ne", "Neon"
PeriodicTable.Add "Na", "Sodium"
PeriodicTable.Add "Mg", "Magnesium"
PeriodicTable.Add "Al", "Aluminium"
PeriodicTable.Add "Si", "Silicon"
PeriodicTable.Add "P", "Phosphorus"
PeriodicTable.Add "S", "Sulfur"
PeriodicTable.Add "Cl", "Chlorine"
PeriodicTable.Add "Ar", "Argon"
PeriodicTable.Add "K", "Potassium"
PeriodicTable.Add "Ca", "Calcium"
PeriodicTable.Add "Sc", "Scandium"
PeriodicTable.Add "Ti", "Titanium"
PeriodicTable.Add "V", "Vanadium"
PeriodicTable.Add "Cr", "Chromium"
PeriodicTable.Add "Mn", "Manganese"
PeriodicTable.Add "Fe", "Iron"
PeriodicTable.Add "Co", "Cobalt"
PeriodicTable.Add "Ni", "Nickel"
PeriodicTable.Add "Cu", "Copper"
PeriodicTable.Add "Zn", "Zinc"
PeriodicTable.Add "Ga", "Gallium"
PeriodicTable.Add "Ge", "Germanium"
PeriodicTable.Add "As", "Arsenic"
PeriodicTable.Add "Se", "Selenium"
PeriodicTable.Add "Br", "Bromine"
PeriodicTable.Add "Kr", "Krypton"
PeriodicTable.Add "Rb", "Rubidium"
PeriodicTable.Add "Sr", "Strontium"
PeriodicTable.Add "Y", "Yttrium"
PeriodicTable.Add "Zr", "Zirconium"
PeriodicTable.Add "Nb", "Niobium"
PeriodicTable.Add "Mo", "Molybdenum"
PeriodicTable.Add "Tc", "Technetium"
PeriodicTable.Add "Ru", "Ruthenium"
PeriodicTable.Add "Rh", "Rhodium"
PeriodicTable.Add "Pd", "Palladium"
PeriodicTable.Add "Ag", "Silver"
PeriodicTable.Add "Cd", "Cadmium"
PeriodicTable.Add "In", "Indium"
PeriodicTable.Add "Sn", "Tin"
PeriodicTable.Add "Sb", "Antimony"
PeriodicTable.Add "Te", "Tellurium"
PeriodicTable.Add "I", "Iodine"
PeriodicTable.Add "Xe", "Xenon"
PeriodicTable.Add "Cs", "Caesium"
PeriodicTable.Add "Ba", "Barium"
PeriodicTable.Add "La", "Lanthanum"
PeriodicTable.Add "Ce", "Cerium"
PeriodicTable.Add "Pr", "Praseodymium"
PeriodicTable.Add "Nd", "Neodymium"
PeriodicTable.Add "Pm", "Promethium"
PeriodicTable.Add "Sm", "Samarium"
PeriodicTable.Add "Eu", "Europium"
PeriodicTable.Add "Gd", "Gadolinium"
PeriodicTable.Add "Tb", "Terbium"
PeriodicTable.Add "Dy", "Dysprosium"
PeriodicTable.Add "Ho", "Holmium"
PeriodicTable.Add "Er", "Erbium"
PeriodicTable.Add "Tm", "Thulium"
PeriodicTable.Add "Yb", "Ytterbium"
PeriodicTable.Add "Lu", "Lutetium"
PeriodicTable.Add "Hf", "Hafnium"
PeriodicTable.Add "Ta", "Tantalum"
PeriodicTable.Add "W", "Tungsten"
PeriodicTable.Add "Re", "Rhenium"
PeriodicTable.Add "Os", "Osmium"
PeriodicTable.Add "Ir", "Iridium"
PeriodicTable.Add "Pt", "Platinum"
PeriodicTable.Add "Au", "Gold"
PeriodicTable.Add "Hg", "Mercury"
PeriodicTable.Add "Tl", "Thallium"
PeriodicTable.Add "Pb", "Lead"
PeriodicTable.Add "Bi", "Bismuth"
PeriodicTable.Add "Po", "Polonium"
PeriodicTable.Add "At", "Astatine"
PeriodicTable.Add "Rn", "Radon"
PeriodicTable.Add "Fr", "Francium"
PeriodicTable.Add "Ra", "Radium"
PeriodicTable.Add "Ac", "Actinium"
PeriodicTable.Add "Th", "Thorium"
PeriodicTable.Add "Pa", "Protactinium"
PeriodicTable.Add "U", "Uranium"
PeriodicTable.Add "Np", "Neptunium"
PeriodicTable.Add "Pu", "Plutonium"
PeriodicTable.Add "Am", "Americium"
PeriodicTable.Add "Cm", "Curium"
PeriodicTable.Add "Bk", "Berkelium"
PeriodicTable.Add "Cf", "Californium"
PeriodicTable.Add "Es", "Einsteinium"
PeriodicTable.Add "Fm", "Fermium"
PeriodicTable.Add "Md", "Mendelevium"
PeriodicTable.Add "No", "Nobelium"
PeriodicTable.Add "Lr", "Lawrencium"
PeriodicTable.Add "Rf", "Rutherfordium"
PeriodicTable.Add "Db", "Dubnium"
PeriodicTable.Add "Sg", "Seaborgium"
PeriodicTable.Add "Bh", "Bohrium"
PeriodicTable.Add "Hs", "Hassium"
PeriodicTable.Add "Mt", "Meitnerium"
PeriodicTable.Add "Ds", "Darmstadtium"
PeriodicTable.Add "Rg", "Roentgenium"
PeriodicTable.Add "Cn", "Copernicium"
PeriodicTable.Add "Nh", "Nihonium"
PeriodicTable.Add "Fl", "Flerovium"
PeriodicTable.Add "Mc", "Moscovium"
PeriodicTable.Add "Lv", "Livermorium"
PeriodicTable.Add "Ts", "Tennessine"
PeriodicTable.Add "Og", "Oganesson"

If PeriodicTable.Exists(seltext) Then
    MsgBox "Element symbol: " & selection & vbCrLf & "Full name: " & PeriodicTable.Item(seltext)
    Else
        MsgBox "Sorry, element not found."
End If

End Sub
