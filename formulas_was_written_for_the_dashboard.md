# Excel Formulas

```excel
=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)


' Searches for the value in C2 within column A of the "customers" sheet and returns the corresponding value from column B.'

=IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0) = 0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0))


' Uses XLOOKUP to find C2 in column A and return column C’s value. If the result is 0, it returns an empty string instead.'

=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)


' Searches for C2 in column A of the "customers" sheet and retrieves the corresponding value from column G.'

=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0))


' Uses INDEX and MATCH to find the intersection of a row (matching D2 in column A) and a column (matching I1 in row 1) in the "products" sheet.'

=IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica"))))


' Converts coffee type codes into their full names (e.g., "Rob" → "Robusta", "Exc" → "Excelsa", etc.).'

=IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))


' Converts roast level codes into full descriptions (e.g., "M" → "Medium", "L" → "Light", etc.).'

=XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0)


' Searches for the structured reference [@[Customer ID]] in column A of "customers" and returns the corresponding value from column I.'
