# MSWord_PeriodicTable

A VBA macro for Word that allows for input of a an element's symbol, name, or atomic number to give you all three pieces of information about that element.

The search functionality works by first checking against selected text, or, if it can't find anything, by allowing the user to input the information.

## WhatElement and WhatChemSymbol - older, worse versions

These two macros work in only one direction, inputting a chemical symbol to give a name (WhatElement), or vice versa (WhatChemSymbol), because these use dictionary objects. It's less useful, but I'm keeping then here because I worked on them before doing it with an array, so they're essentially earlier versions that were abandoned because I found out only after making them that a Dictionary only really works in one direction.

Naturally, he WhatElement and WhatChemSymbol macros are less useful than the PeriodicTable macro, but they gave me a chance to practice with Dictionary objects for future reference.
