<div align="center">

## Very Very Basic encryption and decryption routine \- but effective


</div>

### Description

This is a very simple enocde decode routine that I created for a program where I didn't want to store the password in a plaintext string. This routine then came in handy when creating a client/server program where to stop people sending spurious commands to either the server of the client - I encrypted the time in minutes and it was sent with every bit of data - if the time was out by 3 minutes then the server or client discared the data - this is just one example for this bit of code.

The routine currently runs 100 times in encryption or decryption but is a for loop - so easy to change - just make sure that the encryption and decryption routines run the same number of times.

As an example "Planet Source Code Rocks" becomes "8=EMPT\chipu{$'/49;>?DEHKRY\dehiosuv{}~$+1369;CKQRZ[^`cfglpuw~$%*0479AFKNQSZ]ekry~$*-48:ACHOVW_elrz$D`UbYhrGcifWYr7cXYrFcW_g"

The key to the encryption and decryption is the number at the beginning of the string (which is randomly chosen during encryption - and because it is a printable character gets encrypted itself on every subsequent encryption). This routine is for from unbreakable but is going to throw of the person who just wants to quickly find the password.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-01-19 00:55:02
**By**             |[Andrew Simon Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-simon-jones.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Very\_Very\_1841431192005\.zip](https://github.com/Planet-Source-Code/andrew-simon-jones-very-very-basic-encryption-and-decryption-routine-but-effective__1-58360/archive/master.zip)








