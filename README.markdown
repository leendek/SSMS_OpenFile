SSMS_OpenFile
-------------

To use this addin, stored procedures/functions should be saved in a file with their own name as filename.

Usage
-----
Suppose you have the files procA.sql and procB.sql

procA.sql:

```
CREATE PROCEDURE procA
AS
EXEC procB
```

procB.sql:

```
CREATE PROCEDURE procB
AS
PRINT 'something'
```

If procA.sql is opened in the query editor:

* put the cursor on procB

 * use the shortcut CTRL+k, CTRL+o

 * OR: click Open File2 in the context menu

 * OR: choose Open File2 from Tools > SSMS_OpenFile2 > Open File2

* the addin will look for procB.sql in the folder where procA.sql is in. If not found, it will look for procB.sql in C:\svn.
* open procB.sql