/*
List all files that share a rsidR with Normal.dotm template file.

Normal.dotm may or may not have rsidR values in it. If it does, and you find other files with the same ones,
you can opine that those other files were created from that template.

And because normal.dotm is saved under %AppData%\Microsoft\Templates on a Windows computer,
those values are likely unique to a specific user. An exception to this could be in a corporate environment where 
a default folder layout is created for a new user and within it, it inherits a the same template file.

*/

SELECT file_name, rsid_value
FROM rsids
WHERE rsid_value IN (SELECT rsid_value FROM rsids WHERE file_name LIKE "%\Normal.dotm")
ORDER BY rsid_value