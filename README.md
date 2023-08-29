# MSWord_PeriodicTable

VBA macros for word that allow for input of a chemical symbol, with an output telling you what the full name is (WhatElement), or vice versa (WhatChemSymbol). PeriodicTable does so in both directions, by using an array instead of a dictionary, allowing for more flexible searching.

The search functionality works by first checking against selected text, or, if it can't find anything, by allowing the user to input the element symbol (WhatElement) or name (WhatChemSymbol) before checking. In PeriodicTable, you can input either.

The WhatElement and WhatChemSymbol macros are less useful than the PeriodicTable macro, because of how Dictionary objects work, but it gave me a chance to practice with Dictionary objects for future reference.
