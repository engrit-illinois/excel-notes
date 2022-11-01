# Summary
Just a repository for some handy excel tricks, mostly related to data munging.  
<br />
<br />

# Entries

### Get unique values from a column, excluding blanks and the column header

If your list with header is in sheet 'Data', column A:  
`=FILTER(UNIQUE(A:A),(UNIQUE(A:A)<>"Header text")*(UNIQUE(A:A)<>0)))`

Here's a more condensed, more readable version using LET:  
`=LET(x,UNIQUE(Data!A:A),SORT(FILTER(x,(x<>"Header text")*(x<>0))))`

To translate, it says:  
"Let X equal UNIQUE(A:A), remove values of "Header text" and 0 from X, return what's left, and sort it alphabetically."  


<br />
<br />

# Notes
- By mseng3. See my other projects here: https://github.com/mmseng/code-compendium.
