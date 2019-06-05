# VBA-Row-Expander
Problem: some databases will, for various reasons, store data you want to treat as multiple rows in your spreadsheet in a single entry. This attempts to expand those entries.
Current idiosyncracies include magic numbers in the Range functions, splitting into multiple sheets based on the first column rather than a user-designated column, and splitting based on a single delimiter in the code rather than in user input (though this does work better for hard-to-type delimiters).
