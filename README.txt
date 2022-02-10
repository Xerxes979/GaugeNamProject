This is just a script I wrote for my work at Nifco. I was given a large dataset of gauge and nam numbers and asked to break up entries with multiple numbers represented, so that each of the related parts has their own separate row in the spreadsheet. I was also asked to sort them by gauge and nam number. I decided that rather than editing about 800 lines by hand, I would write a script to do it instead. 


In a nutshell: 
2 spreadsheets (dummysheet and dummysheet2) are imported to pandas dataframes. the original is just to read from, and the second starts off as a direct copy. The original is essentially for the sake of redundancy. dummysheet 2 is deleted, re-created from dummysheet, and reprocessed each time the script runs just to make sure the code is not dependent on previous modifications made on the spreadsheet to function correctly. I wanted it to be as free-standing as possible. 

The sheets are copied into dataframes, the column headers are printed for debugging purposes, and any rows with the characters "/(&#" in the gauge number or the nam number are duplicated. 

Then it sorts the spreadsheet by gauge number and resets the indexes to make them accurate after the sort

Then to edit, it goes through and finds every row with one of the mentioned characters in the gauge numbers or the nam numbers, and cuts them based on the usage of the character. 
For example, with the '/' character, the string is split on '/', the first element of that split goes in the first row, and the second element of the split goes in the duplicate row. There are nested if statements that exist to keep extra rows from being created with there are 2 gauge numbers and 2 nam numbers. 

The indexes are reset again for readability and the dataframe is written to dummysheet2.



